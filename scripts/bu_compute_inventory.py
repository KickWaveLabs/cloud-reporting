#!/usr/bin/env python3
"""Compute backup inventory: Azure VMs, AKS/ARO clusters and MARS/MABS/DPM.

Outputs (in --output-dir):
  compute_vm_inventory.csv         one row per VM: sizing, classification, Azure Backup, ASR, findings
  compute_k8s_inventory.csv        one row per AKS / ARO cluster: sizing, backup extension, Backup vault
  compute_mars_mabs_dpm.csv        MARS agents, MABS and DPM servers and their protected items, if any
  compute_orphan_backup_items.csv  VM backup items whose VM no longer exists or with no backup in 30 days
  compute_backup_policies.csv      VM and AKS backup policies
  compute_backup_vaults.csv        vaults with redundancy and item counts
  compute_errors.csv               access gaps and API errors
  compute_summary.json             pivots: environment x class, backup coverage, findings
  compute_inventory.xlsx           with --xlsx

Dependencies:
  pip install azure-identity azure-mgmt-resourcegraph requests
  pip install openpyxl   # optional, only for --xlsx
"""

from collections import Counter

import bu_common as bu

DEFAULT_OUTPUT_DIR = "./results/backup-compute"
CONTAINER_TYPES = ("MAB", "DPM", "AzureBackupServer")

VM_FIELDS = (
    ["service_type"] + bu.BASE_FIELDS
    + ["vm_size", "os_type", "power_state", "priority", "license_type", "zones", "vmss_id",
       "image_publisher", "image_offer", "image_sku",
       "os_disk_name", "os_disk_size_gb", "os_disk_sku", "data_disk_count", "data_disks_total_gb",
       "total_disk_gb", "data_disk_skus"]
    + ["protected_disk_luns", "excluded_disk_luns", "has_disk_selection", "workload_backup_items",
       "asr_replicated"]
    + ["backup_method"] + bu.PROTECTION_FIELDS + bu.RESILIENCE_FIELDS + bu.FINDING_FIELDS
)

K8S_FIELDS = (
    ["service_type"] + bu.BASE_FIELDS
    + ["kubernetes_version", "sku_tier", "power_state", "node_pool_count", "node_count", "max_node_count",
       "vm_sizes", "os_disk_total_gb", "node_resource_group", "private_cluster", "network_plugin"]
    + ["backup_extension_installed", "gitops_flux", "extension_types"]
    + ["backup_method"] + bu.PROTECTION_FIELDS + bu.RESILIENCE_FIELDS + bu.FINDING_FIELDS
)

MARS_FIELDS = [
    "record_type", "backup_management_type", "vault_name", "vault_resource_group", "vault_subscription_id",
    "vault_location", "vault_storage_redundancy", "container_type", "friendly_name", "server_name",
    "agent_version", "protected_item_count", "health_status", "protection_state", "last_backup_status",
    "last_backup_time", "last_refreshed_at", "oldest_recovery_point", "recovery_point_count",
    "policy_name", "protection_group", "disk_storage_used_gb", "id",
]

QUERIES = {
    "vms": (
        "resources"
        " | where type =~ 'microsoft.compute/virtualmachines'"
        " | project id, name, resourceGroup, subscriptionId, location, tags, zones,"
        " vmSize = tostring(properties.hardwareProfile.vmSize),"
        " osType = tostring(properties.storageProfile.osDisk.osType),"
        " osDiskId = tolower(tostring(properties.storageProfile.osDisk.managedDisk.id)),"
        " osDiskName = tostring(properties.storageProfile.osDisk.name),"
        " osDiskSizeGb = properties.storageProfile.osDisk.diskSizeGB,"
        " dataDisks = properties.storageProfile.dataDisks,"
        " powerState = tostring(properties.extended.instanceView.powerState.code),"
        " priority = tostring(properties.priority),"
        " licenseType = tostring(properties.licenseType),"
        " vmssId = tostring(properties.virtualMachineScaleSet.id),"
        " image = properties.storageProfile.imageReference"
    ),
    "disks": (
        "resources"
        " | where type =~ 'microsoft.compute/disks' and isnotempty(managedBy)"
        " | project id = tolower(id), sizeGb = toint(properties.diskSizeGB), sku = tostring(sku.name)"
    ),
    "aks": (
        "resources"
        " | where type =~ 'microsoft.containerservice/managedclusters'"
        " | project id, name, resourceGroup, subscriptionId, location, tags,"
        " skuTier = tostring(sku.tier),"
        " version = tostring(coalesce(properties.currentKubernetesVersion, properties.kubernetesVersion)),"
        " powerState = tostring(properties.powerState.code),"
        " nodeResourceGroup = tostring(properties.nodeResourceGroup),"
        " privateCluster = tostring(properties.apiServerAccessProfile.enablePrivateCluster),"
        " networkPlugin = tostring(properties.networkProfile.networkPlugin),"
        " pools = properties.agentPoolProfiles"
    ),
    "aro": (
        "resources"
        " | where type =~ 'microsoft.redhatopenshift/openshiftclusters'"
        " | project id, name, resourceGroup, subscriptionId, location, tags,"
        " version = tostring(properties.clusterProfile.version),"
        " masterSize = tostring(properties.masterProfile.vmSize),"
        " workers = properties.workerProfiles,"
        " apiVisibility = tostring(properties.apiserverProfile.visibility),"
        " clusterResourceGroup = tostring(properties.clusterProfile.resourceGroupId),"
        " provisioningState = tostring(properties.provisioningState)"
    ),
    # Only the extension type is projected: extension settings can hold sensitive values.
    "extensions": (
        "kubernetesconfigurationresources"
        " | where type in~ ('microsoft.kubernetesconfiguration/extensions',"
        " 'microsoft.kubernetesconfiguration/fluxconfigurations')"
        " | project id, type,"
        " extensionType = tolower(tostring(coalesce(properties.ExtensionType, properties.extensionType)))"
    ),
}


# ---------------------------------------------------------------------------
# VMs
# ---------------------------------------------------------------------------


def vm_backup_items(index, vm_id):
    """Split items whose source is this VM into VM backup items and in-guest workload items."""
    items = index.for_resource(vm_id)
    vm_items = [item for item in items if item["backup_management_type"] == "AzureIaasVM"]
    workload_items = [item for item in items if item["backup_management_type"] != "AzureIaasVM"]
    return vm_items, workload_items


def disk_selection(item, data_luns):
    """Return (protected_luns, excluded_luns, has_selection) for a VM backup item."""
    if not item:
        return "", "", "NOT_APPLICABLE"
    listed = [lun for lun in str(item.get("disk_lun_list") or "").split(";") if lun != ""]
    if not listed and item.get("disk_is_inclusion_list") is None:
        return "ALL", "", "NO"
    all_luns = [str(lun) for lun in data_luns]
    if item.get("disk_is_inclusion_list"):
        protected = listed
        excluded = [lun for lun in all_luns if lun not in listed]
    else:
        excluded = listed
        protected = [lun for lun in all_luns if lun not in listed]
    return ";".join(["OS"] + protected), ";".join(excluded), "YES"


def build_vm_row(ctx, vm, disks, managed_rgs, index, resilience):
    row = {"service_type": "AZURE_VM"}
    row.update(bu.base_columns(ctx, vm, managed_rgs))

    image = vm.get("image") or {}
    data_disks = vm.get("dataDisks") or []
    os_disk = disks.get(vm.get("osDiskId") or "", {})
    data_sizes, data_skus, data_luns = [], Counter(), []
    for disk in data_disks:
        disk_id = bu.normalize(bu.safe_get(disk, "managedDisk", "id"))
        info = disks.get(disk_id, {})
        size = bu.first_value(info.get("sizeGb"), disk.get("diskSizeGB"))
        if size is not None:
            data_sizes.append(int(size))
        sku = bu.first_value(info.get("sku"), bu.safe_get(disk, "managedDisk", "storageAccountType"))
        if sku:
            data_skus[sku] += 1
        data_luns.append(disk.get("lun"))
    os_size = bu.first_value(os_disk.get("sizeGb"), vm.get("osDiskSizeGb"))
    row.update({
        "vm_size": vm.get("vmSize", ""),
        "os_type": vm.get("osType", ""),
        "power_state": str(vm.get("powerState") or "").replace("PowerState/", "") or "unknown",
        "priority": vm.get("priority") or "Regular",
        "license_type": vm.get("licenseType", ""),
        "zones": ";".join(vm.get("zones") or []),
        "vmss_id": vm.get("vmssId", ""),
        "image_publisher": image.get("publisher", ""),
        "image_offer": image.get("offer", ""),
        "image_sku": image.get("sku", ""),
        "os_disk_name": vm.get("osDiskName", ""),
        "os_disk_size_gb": os_size if os_size is not None else "",
        "os_disk_sku": os_disk.get("sku", ""),
        "data_disk_count": len(data_disks),
        "data_disks_total_gb": sum(data_sizes),
        "total_disk_gb": sum(data_sizes) + int(os_size or 0),
        "data_disk_skus": ";".join(f"{sku}={count}" for sku, count in sorted(data_skus.items())),
    })

    vm_items, workload_items = vm_backup_items(index, vm["id"])
    active = [item for item in vm_items if bu.item_is_active(item)]
    shown = active[0] if active else (vm_items[0] if vm_items else None)
    protected_luns, excluded_luns, has_selection = disk_selection(shown, data_luns)
    resilience_cols = bu.resilience_columns(resilience, vm["id"])
    asr = "Azure Site Recovery" in resilience_cols.get("resilience_solutions", "")
    row.update({
        "protected_disk_luns": protected_luns,
        "excluded_disk_luns": excluded_luns,
        "has_disk_selection": has_selection,
        "workload_backup_items": len(workload_items),
        "asr_replicated": "YES" if asr else "NO",
    })

    protection = bu.protection_columns(index, vm_items)
    if active:
        method = "VAULT_RSV+ASR" if asr else "VAULT_RSV"
    elif protection["backup_stopped"] == "YES":
        method = "VAULT_RSV_STOPPED"
    elif asr:
        method = "ASR_ONLY"
    else:
        method = "NONE"
    row["backup_method"] = method
    row.update(protection)
    row.update(resilience_cols)
    row.update(vm_findings(ctx, row, active, excluded_luns).columns())
    return row


def vm_findings(ctx, row, active, excluded_luns):
    findings = bu.Findings()
    if row["backup_scope"] == "EXCLUDED_PLATFORM":
        findings.add("EXCLUDED_PLATFORM", "INFO")
        if active:
            findings.add("EXCLUDED_PLATFORM_WITH_BACKUP", "LOW")
        return findings
    if row["backup_scope"] == "REVIEW_REQUIRED":
        findings.add("REVIEW_MANAGED_PLATFORM", "INFO")

    environment = row["environment"]
    if environment == "UNKNOWN":
        findings.add("REQUIRES_ENVIRONMENT_CLASSIFICATION", "MEDIUM")
    if active:
        if not all(bu.item_is_healthy(item, ctx.args.stale_hours) for item in active):
            findings.add("PROD_BACKUP_UNHEALTHY" if environment == "PROD" else "NONPROD_BACKUP_UNHEALTHY",
                         "HIGH" if environment == "PROD" else "LOW")
        if row["policy_resolved"] != "YES":
            findings.add("ACTIVE_BACKUP_POLICY_NOT_RESOLVED", "MEDIUM")
        if environment in ("CERT", "DESA"):
            findings.add("REVIEW_NONPROD_BACKUP", "LOW")
        if row["power_state"] in ("deallocated", "stopped"):
            findings.add("DEALLOCATED_VM_WITH_BACKUP", "LOW")
        if excluded_luns:
            findings.add("BACKUP_WITH_DISK_EXCLUSION", "INFO")
    else:
        if row["retained_restore_point"] == "YES":
            findings.add("BACKUP_STOPPED_WITH_RESTORE_POINT", "LOW")
        if environment == "PROD":
            findings.add("PROD_WITHOUT_BACKUP", "HIGH")
    return findings


# ---------------------------------------------------------------------------
# AKS / ARO
# ---------------------------------------------------------------------------


def cluster_id_from_extension(extension_id):
    lowered = bu.normalize(extension_id)
    return lowered.split("/providers/microsoft.kubernetesconfiguration/")[0]


def index_extensions(rows):
    """Return {cluster_id_lower: {"types": set, "flux_configs": int}}."""
    result = {}
    for row in rows:
        entry = result.setdefault(cluster_id_from_extension(row["id"]), {"types": set(), "flux_configs": 0})
        if row["type"].lower().endswith("/fluxconfigurations"):
            entry["flux_configs"] += 1
        elif row.get("extensionType"):
            entry["types"].add(row["extensionType"])
    return result


def build_aks_row(ctx, cluster, managed_rgs, index, resilience, extensions):
    row = {"service_type": "AKS"}
    row.update(bu.base_columns(ctx, cluster, managed_rgs))
    pools = cluster.get("pools") or []
    sizes = Counter()
    node_count = max_nodes = disk_gb = 0
    for pool in pools:
        count = int(pool.get("count") or 0)
        node_count += count
        max_nodes += int(pool.get("maxCount") or count) if pool.get("enableAutoScaling") else count
        sizes[pool.get("vmSize", "")] += count
        disk_gb += count * int(pool.get("osDiskSizeGB") or 0)
    ext = extensions.get(bu.normalize(cluster["id"]), {"types": set(), "flux_configs": 0})
    backup_extension = "microsoft.dataprotection.kubernetes" in ext["types"]
    flux = "microsoft.flux" in ext["types"] or ext["flux_configs"] > 0
    row.update({
        "kubernetes_version": cluster.get("version", ""),
        "sku_tier": cluster.get("skuTier", ""),
        "power_state": cluster.get("powerState", ""),
        "node_pool_count": len(pools),
        "node_count": node_count,
        "max_node_count": max_nodes,
        "vm_sizes": ";".join(f"{size}x{count}" for size, count in sorted(sizes.items())),
        "os_disk_total_gb": disk_gb,
        "node_resource_group": cluster.get("nodeResourceGroup", ""),
        "private_cluster": bu.yes_no(str(cluster.get("privateCluster")).lower() == "true"
                                     if cluster.get("privateCluster") else None),
        "network_plugin": cluster.get("networkPlugin", ""),
        "backup_extension_installed": "YES" if backup_extension else "NO",
        "gitops_flux": "YES" if flux else "NO",
        "extension_types": ";".join(sorted(ext["types"])),
    })
    items = index.for_resource(cluster["id"])
    protection = bu.protection_columns(index, items)
    active = [item for item in items if bu.item_is_active(item)]
    if active:
        method = "VAULT_BV"
    elif protection["backup_stopped"] == "YES":
        method = "VAULT_BV_STOPPED"
    elif backup_extension:
        method = "BACKUP_EXTENSION_ONLY"
    elif flux:
        method = "REBUILD_GITOPS"
    else:
        method = "NONE"
    row["backup_method"] = method
    row.update(protection)
    row.update(bu.resilience_columns(resilience, cluster["id"]))

    findings = bu.Findings()
    environment = row["environment"]
    if environment == "UNKNOWN":
        findings.add("REQUIRES_ENVIRONMENT_CLASSIFICATION", "MEDIUM")
    if row["power_state"].lower() == "stopped":
        findings.add("CLUSTER_STOPPED", "INFO")
    if active:
        if not all(bu.item_is_healthy(item, ctx.args.stale_hours) for item in active):
            findings.add("PROD_BACKUP_UNHEALTHY" if environment == "PROD" else "NONPROD_BACKUP_UNHEALTHY",
                         "HIGH" if environment == "PROD" else "LOW")
        if environment in ("CERT", "DESA"):
            findings.add("REVIEW_NONPROD_BACKUP", "LOW")
    elif environment == "PROD":
        findings.add("PROD_CLUSTER_WITHOUT_BACKUP", "MEDIUM")
    if backup_extension and not active:
        findings.add("BACKUP_EXTENSION_WITHOUT_INSTANCE", "LOW")
    row.update(findings.columns())
    return row


def build_aro_row(ctx, cluster, managed_rgs, resilience):
    row = {"service_type": "ARO"}
    row.update(bu.base_columns(ctx, cluster, managed_rgs))
    workers = cluster.get("workers") or []
    sizes = Counter()
    node_count = disk_gb = 0
    for pool in workers:
        count = int(pool.get("count") or 0)
        node_count += count
        sizes[pool.get("vmSize", "")] += count
        disk_gb += count * int(pool.get("diskSizeGB") or 0)
    if cluster.get("masterSize"):
        sizes[f"{cluster['masterSize']}(master)"] += 3
        node_count += 3
    row.update({
        "kubernetes_version": f"openshift-{cluster.get('version', '')}",
        "sku_tier": "",
        "power_state": cluster.get("provisioningState", ""),
        "node_pool_count": len(workers),
        "node_count": node_count,
        "max_node_count": node_count,
        "vm_sizes": ";".join(f"{size}x{count}" for size, count in sorted(sizes.items())),
        "os_disk_total_gb": disk_gb,
        "node_resource_group": cluster.get("clusterResourceGroup", "").split("/")[-1],
        "private_cluster": "YES" if cluster.get("apiVisibility") == "Private" else "NO",
        "network_plugin": "",
        "backup_extension_installed": "NOT_SUPPORTED",
        "gitops_flux": "",
        "extension_types": "",
        "backup_method": "NOT_SUPPORTED_NATIVE",
    })
    row.update(bu.protection_columns(bu.BackupIndex(), []))
    row.update(bu.resilience_columns(resilience, cluster["id"]))
    findings = bu.Findings()
    if row["environment"] == "UNKNOWN":
        findings.add("REQUIRES_ENVIRONMENT_CLASSIFICATION", "MEDIUM")
    # Azure Backup does not cover ARO; OADP/Velero inside the cluster is not visible from Azure.
    findings.add("ARO_BACKUP_NOT_VISIBLE", "MEDIUM" if row["environment"] == "PROD" else "INFO")
    row.update(findings.columns())
    return row


# ---------------------------------------------------------------------------
# MARS / MABS / DPM
# ---------------------------------------------------------------------------


def fetch_vault_agents(ctx, vault):
    """Containers and items for MARS (MAB), DPM and MABS (AzureBackupServer) in one vault."""
    base = vault["id"]
    rows = []
    for management_type in CONTAINER_TYPES:
        params = {"$filter": f"backupManagementType eq '{management_type}'"}
        containers = ctx.try_arm("mars_mabs_dpm", base, ctx.arm_list,
                                 f"{base}/backupProtectionContainers", bu.RSV_API, params)
        if not containers:
            continue
        for container in containers:
            rows.append(agent_row("CONTAINER", management_type, vault, container))
        items = ctx.try_arm("mars_mabs_dpm", base, ctx.arm_list,
                            f"{base}/backupProtectedItems", bu.RSV_API, params) or []
        for item in items:
            rows.append(agent_row("ITEM", management_type, vault, item))
    return rows


def agent_row(record_type, management_type, vault, body):
    properties = body.get("properties") or {}
    extended = properties.get("extendedInfo") or {}
    used_bytes = extended.get("diskStorageUsedInBytes")
    return {
        "record_type": record_type,
        "backup_management_type": management_type,
        "vault_name": vault["name"],
        "vault_resource_group": vault["resource_group"],
        "vault_subscription_id": vault["subscription_id"],
        "vault_location": vault["location"],
        "vault_storage_redundancy": vault["storage_redundancy"],
        "container_type": properties.get("containerType") or properties.get("protectedItemType", ""),
        "friendly_name": properties.get("friendlyName", ""),
        "server_name": bu.first_value(properties.get("computerName"), properties.get("backupEngineName"),
                                      ";".join(properties.get("dpmServers") or [])) or "",
        "agent_version": bu.first_value(properties.get("agentVersion"), properties.get("dpmAgentVersion")) or "",
        "protected_item_count": properties.get("protectedItemCount", ""),
        "health_status": bu.first_value(properties.get("healthStatus"), properties.get("containerHealthState"),
                                        properties.get("protectionStatus")) or "",
        "protection_state": properties.get("protectionState", ""),
        "last_backup_status": bu.first_value(properties.get("lastBackupStatus"),
                                             extended.get("lastBackupStatus")) or "",
        "last_backup_time": bu.iso(properties.get("lastBackupTime")),
        "last_refreshed_at": bu.iso(extended.get("lastRefreshedAt")),
        "oldest_recovery_point": bu.iso(extended.get("oldestRecoveryPoint")),
        "recovery_point_count": extended.get("recoveryPointCount", ""),
        "policy_name": bu.first_value(properties.get("policyName"), extended.get("policyName")) or "",
        "protection_group": extended.get("protectionGroupName", ""),
        "disk_storage_used_gb": round(used_bytes / 1024 ** 3, 2) if used_bytes else "",
        "id": body.get("id", ""),
    }


# ---------------------------------------------------------------------------
# Orphans
# ---------------------------------------------------------------------------


def vm_orphans(ctx, index, live_vm_ids):
    rows = bu.orphan_item_rows(index, live_vm_ids, lambda item: item["backup_management_type"] == "AzureIaasVM")
    for row in rows:
        # A VM in a subscription we cannot read looks deleted; report it as an access gap instead.
        source_sub = bu.parse_resource_id(row["source_resource_id"])["subscription_id"]
        if row["orphan_reason"] == "SOURCE_RESOURCE_DELETED" and bu.normalize(source_sub) not in ctx.subscriptions:
            row["orphan_reason"] = "SOURCE_SUBSCRIPTION_NOT_VISIBLE"
    return rows


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    parser = bu.build_parser("Compute backup inventory: Azure VMs, AKS/ARO and MARS/MABS/DPM.", DEFAULT_OUTPUT_DIR)
    parser.add_argument("--skip-agents", action="store_true",
                        help="Skip the MARS/MABS/DPM lookup (one ARM call per vault and agent type)")
    args = parser.parse_args()
    bu.setup_logging(args.verbose)
    ctx = bu.AzureContext(args)
    ctx.load_subscriptions()

    data = ctx.queries(QUERIES)
    managed_rgs = bu.load_managed_resource_groups(ctx)
    index = bu.load_backup_index(ctx)
    resilience = bu.load_resilience(ctx, ["AzureVirtualMachine", "KubernetesServices"])

    disks = {row["id"]: row for row in data["disks"]}
    vm_rows = [build_vm_row(ctx, vm, disks, managed_rgs, index, resilience) for vm in data["vms"]]
    vm_rows.sort(key=lambda row: (row["subscription"], row["resource_group"].lower(), row["name"].lower()))

    extensions = index_extensions(data["extensions"])
    k8s_rows = [build_aks_row(ctx, cluster, managed_rgs, index, resilience, extensions) for cluster in data["aks"]]
    k8s_rows += [build_aro_row(ctx, cluster, managed_rgs, resilience) for cluster in data["aro"]]
    k8s_rows.sort(key=lambda row: (row["service_type"], row["subscription"], row["name"].lower()))

    agent_rows = []
    rsv_vaults = [vault for vault in index.vaults.values() if vault["kind"] == "RECOVERY_SERVICES_VAULT"]
    if args.skip_arm or args.skip_agents:
        ctx.errors.add("mars_mabs_dpm", "all vaults", "skipped", code="Skipped",
                       message="MARS/MABS/DPM lookup skipped by option")
    else:
        for result in ctx.parallel(lambda vault: fetch_vault_agents(ctx, vault), rsv_vaults, label="mars_mabs_dpm"):
            agent_rows.extend(result or [])

    orphan_rows = vm_orphans(ctx, index, {bu.normalize(vm["id"]) for vm in data["vms"]})

    output = bu.Output(args.output_dir, args.delimiter)
    output.csv("compute_vm_inventory.csv", vm_rows, VM_FIELDS)
    output.csv("compute_k8s_inventory.csv", k8s_rows, K8S_FIELDS)
    output.csv("compute_mars_mabs_dpm.csv", agent_rows, MARS_FIELDS)
    output.csv("compute_orphan_backup_items.csv", orphan_rows, bu.ORPHAN_FIELDS)
    output.csv("compute_backup_policies.csv",
               bu.policy_rows(index, ["VM", "Microsoft.ContainerService"]), ["policy_id"] + bu.POLICY_FIELDS)
    output.csv("compute_backup_vaults.csv", bu.vault_rows(index))
    output.csv("compute_errors.csv", ctx.errors.rows, bu.ErrorLog.FIELDS)

    traditional = [row for row in vm_rows if row["infrastructure_class"] == "TRADITIONAL"]
    summary = bu.run_metadata(ctx, {
        "counts": {
            "vms": len(vm_rows),
            "k8s_clusters": len(k8s_rows),
            "mars_mabs_dpm_rows": len(agent_rows),
            "mars_mabs_dpm_found": bool(agent_rows),
            "orphan_backup_items": len(orphan_rows),
            "rsv_vaults_checked": 0 if (args.skip_arm or args.skip_agents) else len(rsv_vaults),
        },
        "vm_environment_by_infrastructure_class": bu.count_by(vm_rows, "environment", "infrastructure_class"),
        "traditional_vm_environment_by_backup_enabled": bu.count_by(traditional, "environment", "backup_enabled"),
        "vm_finding_by_environment": bu.count_by(vm_rows, "finding", "environment"),
        "vm_backup_method": bu.count_by(vm_rows, "backup_method"),
        "k8s_service_by_backup_method": bu.count_by(k8s_rows, "service_type", "backup_method"),
        "k8s_finding_by_environment": bu.count_by(k8s_rows, "finding", "environment"),
        "orphans_by_reason": bu.count_by(orphan_rows, "orphan_reason"),
        "mars_mabs_dpm_by_type": bu.count_by(agent_rows, "backup_management_type", "record_type"),
    })
    output.json("compute_summary.json", summary)
    if args.xlsx:
        output.xlsx("compute_inventory.xlsx")
    bu.LOG.info("Done in %.0fs. Gaps: %s", summary["duration_seconds"], summary["errors_by_status"] or "none")


if __name__ == "__main__":
    main()
