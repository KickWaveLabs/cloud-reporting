#!/usr/bin/env python3
"""Storage backup inventory: Blob Storage, ADLS Gen2, Azure Files, Managed Disks and Elastic SAN.

One row per protectable unit (storage account, file share, disk, Elastic SAN volume) with
native recovery features (soft delete, versioning, point-in-time restore, snapshots), Azure
Backup protection (Recovery Services vault and Backup vault), resilience and sizing fields
for later cost analysis.

Outputs (in --output-dir):
  storage_inventory.csv
  storage_snapshots_by_source.csv
  storage_orphan_backup_items.csv
  storage_backup_policies.csv
  storage_errors.csv
  storage_summary.json
  storage_inventory.xlsx            (with --xlsx)

Read-only: Resource Graph queries and ARM GET calls on the control plane. No keys, SAS or
data-plane reads.

Dependencies:
  pip install -r requirements.txt
"""

import logging
from collections import defaultdict
from datetime import timedelta
from itertools import zip_longest

import bu_common as c
from bu_common import LOG

STORAGE_API = "2023-05-01"
METRICS_API = "2023-10-01"
ELASTIC_SAN_API = "2024-05-01"

STORAGE_WORKLOADS = ("azurefileshare", "microsoft.storage/storageaccounts/blobservices",
                     "microsoft.storage/storageaccounts/adlsblobservices", "microsoft.compute/disks",
                     "microsoft.elasticsan/elasticsans/volumegroups/volumes")
# Error codes that mean "not configured" or "not applicable to this account", not an access gap.
NOT_CONFIGURED_CODES = {"ManagementPolicyNotFound", "FeatureNotSupportedForAccount"}
GEO_REDUNDANCY = {"GRS", "RAGRS", "GZRS", "RAGZRS"}
BLOB_KINDS_WITHOUT_FILES = {"blockblobstorage", "blobstorage"}

SIZING_FIELDS = [
    "kind", "sku", "performance", "redundancy", "access_tier", "capacity_gb", "used_capacity_gb",
    "lifecycle_policy", "storage_account", "storage_account_id", "disk_tier", "disk_iops", "disk_mbps",
    "bursting_enabled", "disk_state", "attached_to", "unattached_days", "os_type", "zones",
    "snapshot_count", "snapshot_total_gb", "snapshot_incremental_count", "newest_snapshot",
    "oldest_snapshot_age_days",
]
NATIVE_FIELDS = [
    "hns_enabled", "blob_soft_delete_days", "container_soft_delete_days", "versioning", "change_feed",
    "pitr_days", "immutability", "last_access_tracking", "share_soft_delete_days", "enabled_protocols",
    "native_recovery_features", "vaulted_containers", "covered_by_vm", "vm_disk_lun",
    "vm_backup_disk_included", "data_role_hint", "data_role_evidence",
]
INVENTORY_FIELDS = (["service_type"] + c.BASE_FIELDS + SIZING_FIELDS + NATIVE_FIELDS + ["backup_method"]
                    + c.PROTECTION_FIELDS + c.RESILIENCE_FIELDS + c.FINDING_FIELDS)

SNAPSHOT_FIELDS = [
    "source_resource_id", "source_type", "source_exists", "source_disk_state", "snapshot_count",
    "snapshot_total_gb", "snapshot_incremental_count", "snapshot_full_count", "newest_snapshot",
    "oldest_snapshot", "oldest_snapshot_age_days", "skus", "snapshot_resource_groups",
]


# ---------------------------------------------------------------------------
# Resource Graph
# ---------------------------------------------------------------------------

QUERIES = {
    "storage_accounts": (
        "resources"
        " | where type =~ 'microsoft.storage/storageaccounts'"
        " | project id, name, resourceGroup, subscriptionId, location, tags, kind, sku = tostring(sku.name),"
        " accessTier = tostring(properties.accessTier), hns = tobool(properties.isHnsEnabled),"
        " immutability = tobool(properties.immutableStorageWithVersioning.enabled),"
        " largeFileShares = tostring(properties.largeFileSharesState),"
        " secondaryLocation = tostring(properties.secondaryLocation)"
    ),
    "disks": (
        "resources"
        " | where type =~ 'microsoft.compute/disks'"
        " | project id, name, resourceGroup, subscriptionId, location, tags, managedBy, zones,"
        " sku = tostring(sku.name), sizeGB = toint(properties.diskSizeGB), tier = tostring(properties.tier),"
        " iops = tolong(properties.diskIOPSReadWrite), mbps = tolong(properties.diskMBpsReadWrite),"
        " bursting = tobool(properties.burstingEnabled), state = tostring(properties.diskState),"
        " osType = tostring(properties.osType), created = tostring(properties.timeCreated),"
        " ownershipUpdated = tostring(properties.LastOwnershipUpdateTime)"
    ),
    "vm_disks": (
        "resources"
        " | where type =~ 'microsoft.compute/virtualmachines'"
        " | project id, name, osDisk = tolower(tostring(properties.storageProfile.osDisk.managedDisk.id)),"
        " dataDisks = properties.storageProfile.dataDisks"
    ),
    "snapshots_by_source": (
        "resources"
        " | where type =~ 'microsoft.compute/snapshots'"
        " | extend source = tolower(tostring(properties.creationData.sourceResourceId)),"
        " created = todatetime(properties.timeCreated), incremental = tobool(properties.incremental)"
        " | summarize snapshotCount = count(), totalGB = sum(toint(properties.diskSizeGB)),"
        " incrementalCount = countif(incremental == true), newest = max(created), oldest = min(created),"
        " skus = make_set(tostring(sku.name), 5), resourceGroups = make_set(tolower(resourceGroup), 5)"
        " by id = iff(isempty(source), 'no-source', source)"
    ),
    "elastic_sans": (
        "resources"
        " | where type =~ 'microsoft.elasticsan/elasticsans'"
        " | project id, name, resourceGroup, subscriptionId, location, tags, sku = tostring(sku.name),"
        " baseSizeTiB = toint(properties.baseSizeTiB), extendedSizeTiB = toint(properties.extendedCapacitySizeTiB)"
    ),
    "function_app_groups": (
        "resources"
        " | where type =~ 'microsoft.web/sites' and kind contains 'functionapp'"
        " | summarize n = count() by id = tolower(strcat('/subscriptions/', subscriptionId,"
        " '/resourcegroups/', resourceGroup))"
    ),
}


# ---------------------------------------------------------------------------
# Throttling visibility and ordering
# ---------------------------------------------------------------------------


class RetryCounter(logging.Handler):
    """Counts ARM retries (429/5xx) logged at debug level by bu_common, without printing them."""

    def __init__(self):
        super().__init__(level=logging.DEBUG)
        self.counts = defaultdict(int)

    def emit(self, record):
        if record.msg.startswith("ARM %s, retry"):
            self.counts[str(record.args[0])] += 1


def install_retry_counter(verbose):
    counter = RetryCounter()
    for handler in logging.getLogger().handlers:
        handler.setLevel(logging.DEBUG if verbose else logging.INFO)
    LOG.setLevel(logging.DEBUG)
    LOG.addHandler(counter)
    return counter


def interleave_by_subscription(rows):
    """Round-robin rows across subscriptions so concurrent calls do not hit one subscription's read limit."""
    groups = defaultdict(list)
    for row in rows:
        groups[row.get("subscriptionId")].append(row)
    ordered = []
    for batch in zip_longest(*groups.values()):
        ordered.extend(row for row in batch if row is not None)
    return ordered


# ---------------------------------------------------------------------------
# ARM detail calls
# ---------------------------------------------------------------------------


def get_optional(ctx, stage, path, api_version, params=None):
    """ARM GET that treats 404 as "not configured" instead of an error."""
    try:
        return ctx.arm_get(path, api_version, params)
    except c.ArmError as error:
        if error.status == 404 or error.code in NOT_CONFIGURED_CODES:
            return None
        ctx.errors.add(stage, path, error.url.split("?")[0].replace(c.ARM_ENDPOINT, ""),
                       error.status, error.code, error.message)
        return {"_error": error.status}


def list_optional(ctx, stage, path, api_version, params=None):
    try:
        return ctx.arm_list(path, api_version, params)
    except c.ArmError as error:
        if error.status == 404 or error.code in NOT_CONFIGURED_CODES:
            return []
        ctx.errors.add(stage, path, error.url.split("?")[0].replace(c.ARM_ENDPOINT, ""),
                       error.status, error.code, error.message)
        return None


def used_capacity_gb(ctx, account_id):
    """Latest hourly UsedCapacity of the account (blobs, files, queues and tables), in GB."""
    end = c.utc_now().replace(minute=0, second=0, microsecond=0)
    start = end - timedelta(hours=6)
    body = get_optional(ctx, "capacity_metric", f"{account_id}/providers/Microsoft.Insights/metrics", METRICS_API, {
        "metricnames": "UsedCapacity",
        "aggregation": "Average",
        "interval": "PT1H",
        "timespan": f"{start:%Y-%m-%dT%H:%M:%SZ}/{end:%Y-%m-%dT%H:%M:%SZ}",
    })
    if not body or "_error" in body:
        return ""
    for metric in body.get("value", []):
        for series in metric.get("timeseries", []):
            for point in reversed(series.get("data", [])):
                if point.get("average") is not None:
                    return round(point["average"] / 1024 ** 3, 2)
    return ""


def account_details(ctx, account, args):
    """Blob service, lifecycle policy, file service, shares and capacity of one account."""
    account_id = account["id"]
    kind = str(account.get("kind") or "").lower()
    detail = {"blob": None, "lifecycle": None, "file_service": None, "shares": None, "capacity": ""}
    if kind != "filestorage":
        detail["blob"] = get_optional(ctx, "blob_service", f"{account_id}/blobServices/default", STORAGE_API)
        if not args.skip_lifecycle:
            detail["lifecycle"] = get_optional(
                ctx, "lifecycle_policy", f"{account_id}/managementPolicies/default", STORAGE_API)
    if kind not in BLOB_KINDS_WITHOUT_FILES:
        shares = list_optional(ctx, "file_shares", f"{account_id}/fileServices/default/shares", STORAGE_API,
                               {"$expand": "snapshots"})
        detail["shares"] = shares
        if shares:
            detail["file_service"] = get_optional(
                ctx, "file_service", f"{account_id}/fileServices/default", STORAGE_API)
            if not args.skip_share_stats:
                for share in shares:
                    if not c.safe_get(share, "properties", "snapshotTime"):
                        stats = get_optional(ctx, "share_stats", share["id"], STORAGE_API, {"$expand": "stats"})
                        share["_usage_bytes"] = c.safe_get(stats or {}, "properties", "shareUsageBytes")
    if not args.skip_capacity:
        detail["capacity"] = used_capacity_gb(ctx, account_id)
    return detail


def elastic_san_volumes(ctx, san):
    """Volumes of one Elastic SAN with their snapshot counts."""
    groups = list_optional(ctx, "elastic_san", f"{san['id']}/volumegroups", ELASTIC_SAN_API) or []
    volumes = []
    for group in groups:
        snapshots = list_optional(ctx, "elastic_san", f"{group['id']}/snapshots", ELASTIC_SAN_API) or []
        by_volume = defaultdict(list)
        for snapshot in snapshots:
            source = c.normalize(c.safe_get(snapshot, "properties", "creationData", "sourceId"))
            by_volume[source].append(snapshot)
        for volume in list_optional(ctx, "elastic_san", f"{group['id']}/volumes", ELASTIC_SAN_API) or []:
            volume["_snapshots"] = by_volume.get(c.normalize(volume["id"]), [])
            volume["_group"] = group.get("name", "")
            volumes.append(volume)
    return volumes


# ---------------------------------------------------------------------------
# Row builders
# ---------------------------------------------------------------------------


def split_sku(sku):
    """Standard_RAGRS -> (Standard, RAGRS)."""
    parts = str(sku or "").split("_", 1)
    return (parts[0], parts[1]) if len(parts) == 2 else ("", parts[0])


def protect(row, index, items, args):
    """Add protection columns and remember whether an active backup is unhealthy."""
    row.update(c.protection_columns(index, items))
    row["_unhealthy"] = any(c.item_is_active(item) and not c.item_is_healthy(item, args.stale_hours)
                            for item in items)


def empty_row(service_type, base):
    row = {key: "" for key in INVENTORY_FIELDS}
    row.update(base)
    row["service_type"] = service_type
    return row


def data_role(row, account, function_groups):
    """Hint about whether the account holds business data; it must be validated with the owner."""
    tags = c.tags_lower(account.get("tags"))
    name = str(account.get("name") or "").lower()
    if row["infrastructure_class"] != "TRADITIONAL":
        return "PLATFORM_MANAGED", row["infrastructure_evidence"]
    if tags.get("ms-resource-usage") == "azure-cloud-shell":
        return "CLOUD_SHELL", "TAG:ms-resource-usage=azure-cloud-shell"
    if "diag" in name:
        return "DIAGNOSTICS_HINT", "NAME:*diag*"
    rg_id = c.resource_group_id(account.get("subscriptionId"), account.get("resourceGroup"))
    if rg_id in function_groups:
        return "FUNCTION_OR_APP_HINT", "RESOURCE_GROUP_HAS_FUNCTION_APP"
    return "PENDING_VALIDATION", ""


def account_common(row, account):
    performance, redundancy = split_sku(account.get("sku"))
    row.update({
        "kind": account.get("kind", ""),
        "sku": account.get("sku", ""),
        "performance": performance,
        "redundancy": redundancy,
        "storage_account": account.get("name", ""),
        "storage_account_id": account.get("id", ""),
        "hns_enabled": c.yes_no(bool(account.get("hns"))),
        "immutability": "YES" if account.get("immutability") else "NO",
    })


def blob_row(ctx, account, detail, index, resilience, managed_rgs, function_groups):
    service_type = "ADLS_GEN2" if account.get("hns") else "BLOB_STORAGE"
    row = empty_row(service_type, c.base_columns(ctx, account, managed_rgs))
    account_common(row, account)
    blob = detail.get("blob") or {}
    properties = blob.get("properties") or {}
    lifecycle = detail.get("lifecycle")
    row["access_tier"] = account.get("accessTier", "")
    row["used_capacity_gb"] = detail.get("capacity", "")
    if "_error" in blob or not blob:
        features = ["DETAIL_NOT_AVAILABLE"]
    else:
        soft = properties.get("deleteRetentionPolicy") or {}
        container_soft = properties.get("containerDeleteRetentionPolicy") or {}
        change_feed = properties.get("changeFeed") or {}
        restore = properties.get("restorePolicy") or {}
        row["blob_soft_delete_days"] = soft.get("days", "") if soft.get("enabled") else 0
        row["container_soft_delete_days"] = container_soft.get("days", "") if container_soft.get("enabled") else 0
        row["versioning"] = c.yes_no(bool(properties.get("isVersioningEnabled")))
        row["change_feed"] = (change_feed.get("retentionInDays") or "INFINITE") if change_feed.get("enabled") else "NO"
        row["pitr_days"] = restore.get("days", "") if restore.get("enabled") else 0
        row["last_access_tracking"] = c.yes_no(bool(c.safe_get(properties, "lastAccessTimeTrackingPolicy", "enable")))
        features = []
        if row["blob_soft_delete_days"]:
            features.append(f"BLOB_SOFT_DELETE={row['blob_soft_delete_days']}d")
        if row["container_soft_delete_days"]:
            features.append(f"CONTAINER_SOFT_DELETE={row['container_soft_delete_days']}d")
        if row["versioning"] == "YES":
            features.append("VERSIONING")
        if row["change_feed"] != "NO":
            features.append("CHANGE_FEED")
        if row["pitr_days"]:
            features.append(f"PITR={row['pitr_days']}d")
    if row["immutability"] == "YES":
        features.append("IMMUTABILITY")
    row["native_recovery_features"] = ";".join(features)
    if lifecycle is None:
        row["lifecycle_policy"] = "NO" if not ctx.args.skip_lifecycle else ""
    elif "_error" in lifecycle:
        row["lifecycle_policy"] = "UNKNOWN"
    else:
        rules = c.safe_get(lifecycle, "properties", "policy", "rules") or []
        row["lifecycle_policy"] = f"YES({len([rule for rule in rules if rule.get('enabled', True)])} rules)"
    row["data_role_hint"], row["data_role_evidence"] = data_role(row, account, function_groups)

    items = [item for item in index.for_resource(account["id"])
             if "blobservices" in c.normalize(item["datasource_type"])]
    row["vaulted_containers"] = ";".join(item["protected_items_extra"] for item in items if item["protected_items_extra"])
    protect(row, index, items, ctx.args)
    row.update(c.resilience_columns(resilience, account["id"]))
    if row["backup_enabled"] == "YES":
        row["backup_method"] = "VAULT_BV"
    elif row["pitr_days"] and row["versioning"] == "YES" and row["change_feed"] != "NO":
        row["backup_method"] = "NATIVE_PITR"
    elif row["blob_soft_delete_days"] or row["container_soft_delete_days"] or row["versioning"] == "YES":
        row["backup_method"] = "NATIVE_SOFT_DELETE_ONLY"
    else:
        row["backup_method"] = "NONE" if "DETAIL_NOT_AVAILABLE" not in features else "UNKNOWN"
    return row


def share_rows(ctx, account, detail, index, resilience, managed_rgs, function_groups):
    shares = detail.get("shares") or []
    snapshots = defaultdict(list)
    for share in shares:
        snapshot_time = c.safe_get(share, "properties", "snapshotTime")
        if snapshot_time:
            snapshots[share["name"].lower()].append(snapshot_time)
    file_service = detail.get("file_service") or {}
    soft = c.safe_get(file_service, "properties", "shareDeleteRetentionPolicy") or {}
    soft_days = soft.get("days", "") if soft.get("enabled") else 0
    if "_error" in file_service:
        soft_days = ""
    account_items = [item for item in index.for_resource(account["id"])
                     if c.normalize(item["workload_type"]) == "azurefileshare"]
    base = c.base_columns(ctx, account, managed_rgs)
    rows = []
    for share in shares:
        properties = share.get("properties") or {}
        if properties.get("snapshotTime"):
            continue
        row = empty_row("AZURE_FILES_SHARE", base)
        account_common(row, account)
        row["name"] = share["name"]
        row["id"] = share["id"]
        times = sorted(snapshots.get(share["name"].lower(), []))
        usage = share.get("_usage_bytes")
        row.update({
            "access_tier": properties.get("accessTier", ""),
            "capacity_gb": properties.get("shareQuota", ""),
            "used_capacity_gb": round(usage / 1024 ** 3, 2) if usage is not None else "",
            "snapshot_count": len(times),
            "newest_snapshot": c.iso(times[-1]) if times else "",
            "oldest_snapshot_age_days": c.days_since(times[0]) if times else "",
            "share_soft_delete_days": soft_days,
            "enabled_protocols": properties.get("enabledProtocols", ""),
        })
        features = []
        if soft_days:
            features.append(f"SHARE_SOFT_DELETE={soft_days}d")
        if times:
            features.append(f"SHARE_SNAPSHOTS={len(times)}")
        row["native_recovery_features"] = ";".join(features)
        row["data_role_hint"], row["data_role_evidence"] = data_role(row, account, function_groups)
        items = [item for item in account_items
                 if str(item["protected_item_name"]).lower() == share["name"].lower()]
        protect(row, index, items, ctx.args)
        row.update(c.resilience_columns(resilience, account["id"]))
        if row["backup_enabled"] == "YES":
            row["backup_method"] = "VAULT_RSV"
        elif times:
            row["backup_method"] = "SNAPSHOT_ONLY"
        elif soft_days:
            row["backup_method"] = "NATIVE_SOFT_DELETE_ONLY"
        else:
            row["backup_method"] = "NONE"
        rows.append(row)
    return rows


def vm_disk_map(vm_rows):
    """{disk_id_lower: (vm_id_lower, vm_name, lun)}; lun is "OS" for the OS disk."""
    mapping = {}
    for vm in vm_rows:
        vm_id = c.normalize(vm["id"])
        if vm.get("osDisk"):
            mapping[vm["osDisk"]] = (vm_id, vm["name"], "OS")
        for disk in vm.get("dataDisks") or []:
            disk_id = c.normalize(c.safe_get(disk, "managedDisk", "id"))
            if disk_id:
                mapping[disk_id] = (vm_id, vm["name"], disk.get("lun"))
    return mapping


def disk_included(item, lun):
    """Whether a VM backup item covers the disk at this LUN. The OS disk is always included."""
    if lun == "OS" or not item.get("disk_lun_list"):
        return True
    luns = {value for value in str(item["disk_lun_list"]).split(";") if value != ""}
    listed = str(lun) in luns
    return listed if item.get("disk_is_inclusion_list") else not listed


def disk_row(ctx, disk, index, resilience, managed_rgs, vm_disks, snapshots):
    row = empty_row("MANAGED_DISK", c.base_columns(ctx, disk, managed_rgs))
    disk_id = c.normalize(disk["id"])
    performance, redundancy = split_sku(disk.get("sku"))
    state = disk.get("state", "")
    unattached_days = ""
    if state.lower() == "unattached":
        unattached_days = c.days_since(disk.get("ownershipUpdated") or disk.get("created"))
    snapshot = snapshots.get(disk_id) or {}
    row.update({
        "kind": "ManagedDisk",
        "sku": disk.get("sku", ""),
        "performance": performance,
        "redundancy": redundancy,
        "capacity_gb": disk.get("sizeGB", ""),
        "disk_tier": disk.get("tier", ""),
        "disk_iops": disk.get("iops", ""),
        "disk_mbps": disk.get("mbps", ""),
        "bursting_enabled": c.yes_no(disk.get("bursting")) if disk.get("bursting") is not None else "NO",
        "disk_state": state,
        "attached_to": str(disk.get("managedBy") or "").split("/")[-1],
        "unattached_days": unattached_days,
        "os_type": disk.get("osType", ""),
        "zones": ";".join(disk.get("zones") or []),
        "snapshot_count": snapshot.get("snapshot_count", 0),
        "snapshot_total_gb": snapshot.get("snapshot_total_gb", 0),
        "snapshot_incremental_count": snapshot.get("snapshot_incremental_count", 0),
        "newest_snapshot": snapshot.get("newest_snapshot", ""),
        "oldest_snapshot_age_days": snapshot.get("oldest_snapshot_age_days", ""),
    })
    disk_items = [item for item in index.for_resource(disk_id)
                  if c.normalize(item["datasource_type"]) == "microsoft.compute/disks"]
    vm_items, excluded = [], False
    vm_info = vm_disks.get(disk_id)
    if vm_info:
        vm_id, vm_name, lun = vm_info
        row["vm_disk_lun"] = lun
        candidates = [item for item in index.for_resource(vm_id) if c.normalize(item["workload_type"]) == "vm"]
        if candidates:
            row["covered_by_vm"] = vm_name
            vm_items = [item for item in candidates if disk_included(item, lun)]
            excluded = any(c.item_is_active(item) for item in candidates) and not any(
                c.item_is_active(item) for item in vm_items)
            row["vm_backup_disk_included"] = "NO" if excluded else "YES"
    protect(row, index, disk_items + vm_items, ctx.args)
    row.update(c.resilience_columns(resilience, disk_id))
    bv_active = any(c.item_is_active(item) for item in disk_items)
    vm_active = any(c.item_is_active(item) for item in vm_items)
    if bv_active:
        row["backup_method"] = "VAULT_BV"
    elif vm_active:
        row["backup_method"] = "VM_BACKUP"
    elif row["snapshot_count"]:
        row["backup_method"] = "SNAPSHOT_ONLY"
    else:
        row["backup_method"] = "NONE"
    features = []
    if row["snapshot_count"]:
        features.append(f"SNAPSHOTS={row['snapshot_count']}")
    if redundancy.upper() == "ZRS":
        features.append("ZRS")
    row["native_recovery_features"] = ";".join(features)
    row["_bv_active"] = bv_active
    row["_vm_active"] = vm_active
    row["_excluded"] = excluded
    return row


def elastic_san_rows(ctx, san, volumes, index, managed_rgs):
    base = c.base_columns(ctx, san, managed_rgs)
    performance, redundancy = split_sku(san.get("sku"))
    rows = []
    for volume in volumes:
        row = empty_row("ELASTIC_SAN_VOLUME", base)
        snapshots = sorted(c.safe_get(item, "systemData", "createdAt", default="") for item in volume["_snapshots"])
        row.update({
            "name": f"{san['name']}/{volume['_group']}/{volume['name']}",
            "id": volume["id"],
            "kind": "ElasticSanVolume",
            "sku": san.get("sku", ""),
            "performance": performance,
            "redundancy": redundancy,
            "capacity_gb": c.safe_get(volume, "properties", "sizeGiB", default=""),
            "snapshot_count": len(volume["_snapshots"]),
            "newest_snapshot": c.iso(snapshots[-1]) if snapshots and snapshots[-1] else "",
        })
        protect(row, index, index.for_resource(volume["id"]), ctx.args)
        row["backup_method"] = ("VAULT_BV" if row["backup_enabled"] == "YES"
                                else "SNAPSHOT_ONLY" if row["snapshot_count"] else "NONE")
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# Findings
# ---------------------------------------------------------------------------


def apply_findings(row, args):
    findings = c.Findings()
    environment = row["environment"]
    service = row["service_type"]
    if row["backup_scope"] == "EXCLUDED_PLATFORM":
        row.update(findings.columns(default="EXCLUDED_PLATFORM"))
        return
    if environment == "UNKNOWN":
        findings.add("REQUIRES_ENVIRONMENT_CLASSIFICATION", "MEDIUM")
    method = row["backup_method"]
    has_backup = method in ("VAULT_BV", "VAULT_RSV", "VM_BACKUP")
    if environment == "PROD":
        if method == "NONE":
            findings.add("PROD_WITHOUT_RECOVERY", "HIGH")
        if service == "AZURE_FILES_SHARE" and not has_backup:
            findings.add("PROD_SHARE_WITHOUT_BACKUP", "HIGH")
        if service != "MANAGED_DISK" and row["redundancy"].upper() == "LRS" and not has_backup:
            findings.add("PROD_LRS_NO_BACKUP", "MEDIUM")
    if row.get("_excluded"):
        findings.add("DISK_EXCLUDED_FROM_VM_BACKUP", "HIGH" if environment == "PROD" else "MEDIUM")
    if service == "MANAGED_DISK" and str(row["disk_state"]).lower() == "unattached":
        findings.add("UNATTACHED_DISK", "LOW")
    oldest = row["oldest_snapshot_age_days"]
    if service == "MANAGED_DISK" and (int(row["snapshot_count"] or 0) >= args.snapshot_sprawl_count
                                      or (oldest != "" and float(oldest) > args.snapshot_max_age_days)):
        findings.add("SNAPSHOT_SPRAWL", "LOW")
    if environment in ("CERT", "DESA") and row["redundancy"].upper() in GEO_REDUNDANCY:
        findings.add("NONPROD_GEO_REDUNDANT", "LOW")
    if row.get("_vm_active") and (row.get("_bv_active") or int(row["snapshot_count"] or 0) > 0):
        findings.add("DUAL_PROTECTION", "LOW")
    if row["backup_enabled"] == "YES" and str(row["policy_resolved"]).startswith("NO"):
        findings.add("ACTIVE_BACKUP_POLICY_NOT_RESOLVED", "MEDIUM")
    if row.get("_unhealthy"):
        findings.add("BACKUP_UNHEALTHY", "HIGH" if environment == "PROD" else "MEDIUM")
    row.update(findings.columns())


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def merge_snapshot_groups(raw_rows):
    """Merge groups of one source split across subscription chunks (snapshots in another subscription)."""
    merged = {}
    for raw in raw_rows:
        current = merged.get(raw["id"])
        if current is None:
            merged[raw["id"]] = dict(raw)
            continue
        for key in ("snapshotCount", "totalGB", "incrementalCount"):
            current[key] = (current.get(key) or 0) + (raw.get(key) or 0)
        current["newest"] = max(filter(None, [current.get("newest"), raw.get("newest")]), default=None)
        current["oldest"] = min(filter(None, [current.get("oldest"), raw.get("oldest")]), default=None)
        for key in ("skus", "resourceGroups"):
            current[key] = sorted(set(current.get(key) or []) | set(raw.get(key) or []))
    return list(merged.values())


def snapshot_rows(raw_rows, disks):
    rows = {}
    for raw in merge_snapshot_groups(raw_rows):
        source = raw["id"]
        disk = disks.get(source)
        source_type = "NONE" if source == "no-source" else (
            "/".join(source.split("/providers/")[-1].split("/")[:2]) if "/providers/" in source else "OTHER")
        rows[source] = {
            "source_resource_id": source,
            "source_type": source_type,
            "source_exists": "YES" if disk else ("NOT_IN_SCOPE_OR_DELETED" if source_type == "microsoft.compute/disks" else ""),
            "source_disk_state": disk.get("state", "") if disk else "",
            "snapshot_count": raw.get("snapshotCount", 0),
            "snapshot_total_gb": raw.get("totalGB") or 0,
            "snapshot_incremental_count": raw.get("incrementalCount", 0),
            "snapshot_full_count": (raw.get("snapshotCount") or 0) - (raw.get("incrementalCount") or 0),
            "newest_snapshot": c.iso(raw.get("newest")),
            "oldest_snapshot": c.iso(raw.get("oldest")),
            "oldest_snapshot_age_days": c.days_since(raw.get("oldest")),
            "skus": raw.get("skus") or [],
            "snapshot_resource_groups": raw.get("resourceGroups") or [],
        }
    return rows


def main():
    parser = c.build_parser(__doc__.splitlines()[0], c.results_path("backup-storage"))
    parser.set_defaults(workers=32, arm_rps=40.0)
    parser.add_argument("--skip-capacity", action="store_true",
                        help="Skip the UsedCapacity metric call per storage account")
    parser.add_argument("--skip-lifecycle", action="store_true",
                        help="Skip the lifecycle management policy call per storage account")
    parser.add_argument("--skip-share-stats", action="store_true",
                        help="Skip the per-share usage call (shareUsageBytes)")
    parser.add_argument("--snapshot-sprawl-count", type=int, default=10,
                        help="Snapshots per disk from which SNAPSHOT_SPRAWL is raised")
    parser.add_argument("--snapshot-max-age-days", type=int, default=90,
                        help="Oldest snapshot age (days) from which SNAPSHOT_SPRAWL is raised")
    args = parser.parse_args()
    c.setup_logging(args.verbose)
    retries = install_retry_counter(args.verbose)

    ctx = c.AzureContext(args)
    ctx.load_subscriptions()
    index = c.load_backup_index(ctx)
    managed_rgs = c.load_managed_resource_groups(ctx)
    resilience = c.load_resilience(ctx, ["AzureStorageAccount", "AzureDisks"])
    data = ctx.queries(QUERIES)

    accounts = interleave_by_subscription(data["storage_accounts"])
    disks = {c.normalize(disk["id"]): disk for disk in data["disks"]}
    function_groups = {row["id"] for row in data["function_app_groups"]}
    vm_disks = vm_disk_map(data["vm_disks"])
    snapshots = snapshot_rows(data["snapshots_by_source"], disks)

    if args.skip_arm:
        details = [{} for _ in accounts]
        san_volumes = [[] for _ in data["elastic_sans"]]
    else:
        LOG.info("Reading service properties of %d storage accounts", len(accounts))
        details = ctx.parallel(lambda account: account_details(ctx, account, args), accounts, label="storage_accounts")
        san_volumes = ctx.parallel(lambda san: elastic_san_volumes(ctx, san), data["elastic_sans"], label="elastic_san")

    rows = []
    for account, detail in zip(accounts, details):
        detail = detail or {}
        if str(account.get("kind") or "").lower() != "filestorage":
            rows.append(blob_row(ctx, account, detail, index, resilience, managed_rgs, function_groups))
        rows.extend(share_rows(ctx, account, detail, index, resilience, managed_rgs, function_groups))
    for disk in data["disks"]:
        rows.append(disk_row(ctx, disk, index, resilience, managed_rgs, vm_disks, snapshots))
    for san, volumes in zip(data["elastic_sans"], san_volumes):
        rows.extend(elastic_san_rows(ctx, san, volumes or [], index, managed_rgs))

    for row in rows:
        apply_findings(row, args)
    for row in rows:
        for key in [key for key in row if key.startswith("_")]:
            del row[key]

    live_ids = {c.normalize(account["id"]) for account in accounts} | set(disks)
    orphans = c.orphan_item_rows(index, live_ids, lambda item: (
        c.normalize(item["workload_type"]) in STORAGE_WORKLOADS
        or c.normalize(item["datasource_type"]) in STORAGE_WORKLOADS))
    policies = c.policy_rows(index, workloads=["AzureFileShare", "blobServices", "adlsBlobServices",
                                               "Microsoft.Compute/disks", "ElasticSan"])

    output = c.Output(args.output_dir, args.delimiter)
    output.csv("storage_inventory.csv", rows, INVENTORY_FIELDS)
    output.csv("storage_snapshots_by_source.csv", sorted(snapshots.values(), key=lambda row: -row["snapshot_count"]),
               SNAPSHOT_FIELDS)
    output.csv("storage_orphan_backup_items.csv", orphans, c.ORPHAN_FIELDS)
    output.csv("storage_backup_policies.csv", policies)
    output.csv("storage_errors.csv", ctx.errors.rows, c.ErrorLog.FIELDS)
    orphan_snapshots = [row for row in snapshots.values() if row["source_exists"] == "NOT_IN_SCOPE_OR_DELETED"]
    output.json("storage_summary.json", c.run_metadata(ctx, {
        "rows": len(rows),
        "arm_retries_by_status": dict(retries.counts),
        "service_type_by_backup_method": c.count_by(rows, "service_type", "backup_method"),
        "environment_by_finding": c.count_by(rows, "environment", "finding"),
        "redundancy_by_environment": c.count_by(rows, "service_type", "redundancy", "environment"),
        "data_role_hint": c.count_by(rows, "service_type", "data_role_hint"),
        "orphan_backup_items": c.count_by(orphans, "orphan_reason", "workload_type"),
        "snapshots_without_live_source": {
            "sources": len(orphan_snapshots),
            "snapshots": sum(row["snapshot_count"] for row in orphan_snapshots),
            "total_gb": sum(row["snapshot_total_gb"] for row in orphan_snapshots),
        },
        "unattached_disks": {
            "count": sum(1 for row in rows if str(row["disk_state"]).lower() == "unattached"),
            "total_gb": sum(int(row["capacity_gb"] or 0) for row in rows if str(row["disk_state"]).lower() == "unattached"),
        },
    }))
    if args.xlsx:
        output.xlsx("storage_inventory.xlsx")
    LOG.info("Done: %d rows, %d errors", len(rows), len(ctx.errors.rows))


if __name__ == "__main__":
    main()
