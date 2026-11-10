#!/usr/bin/env python3
"""Database backup inventory: native PaaS backups and Azure Backup coverage.

One row per database or database server, with uniform columns so native backup and
vault backup can be compared in a pivot table:
  AZURE_SQL_DB, SQL_DW, SQL_MI_DB, POSTGRES_FLEX, MYSQL_FLEX, POSTGRES_SINGLE,
  MYSQL_SINGLE, COSMOS_DB, COSMOS_MONGO_VCORE, SQL_SERVER_VM, SAP_HANA_VM, SAP_ASE_VM.

Outputs (in --output-dir):
  database_inventory.csv
  database_vm_workload_items.csv      protected and discovered-but-unprotected DBs inside VMs
  database_orphan_backup_items.csv    vault items whose database or VM no longer exists
  database_backup_policies.csv
  database_backup_vaults.csv
  database_errors.csv
  database_summary.json
  database_inventory.xlsx             with --xlsx

Dependencies:
  pip install -r requirements.txt
"""

import sys
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor

import bu_common as c

SQL_API_VERSIONS = ("2023-08-01", "2021-11-01")
SQL_VM_API_VERSIONS = ("2023-10-01", "2022-02-01")
FALLBACK_CODES = {"InvalidApiVersionParameter", "NoRegisteredProviderFound", "InvalidResourceType"}

DB_WORKLOADS = {"sqldatabase", "saphanadatabase", "saphanadbinstance", "sapasedatabase"}
DB_DATASOURCE_MARKERS = ("microsoft.dbforpostgresql", "microsoft.dbformysql", "microsoft.sql")
SAP_HANA_WORKLOADS = {"saphanadatabase", "saphanadbinstance", "saphanasystem", "saphanadbinstancesystem"}
SAP_ASE_WORKLOADS = {"sapasedatabase", "sapasesystem"}

RESILIENCE_TYPES = [
    "AzureSqlDatabase", "AzureSqlManagedInstance", "AzureCosmosDb",
    "AzureDatabaseForPostgreSqlFlexserver", "AzureDatabaseForMySqlFlexserver",
]

SIZING_FIELDS = [
    "server_name", "database_name", "engine_version", "status", "sku_name", "sku_tier", "sku_capacity",
    "compute_model", "elastic_pool", "max_size_gb", "storage_gb", "ha_mode", "zone_redundant",
    "replica_role", "locations_count", "multi_region_write", "vm_size", "license_type",
]
NATIVE_FIELDS = [
    "native_backup", "native_retention_days", "native_backup_interval", "native_backup_redundancy",
    "native_geo_backup", "earliest_restore_date", "diff_backup_interval_hours", "ltr_configured",
    "ltr_weekly", "ltr_monthly", "ltr_yearly", "ltr_max_days", "sql_auto_backup_enabled",
    "sql_auto_backup_retention_days", "sql_auto_backup_schedule", "workload_items_protected",
    "workload_items_unprotected", "vm_backup_enabled", "vm_backup_vault", "notes",
]
INVENTORY_FIELDS = (["service_type"] + c.BASE_FIELDS + SIZING_FIELDS + NATIVE_FIELDS + ["backup_method"]
                    + c.PROTECTION_FIELDS + c.RESILIENCE_FIELDS + c.FINDING_FIELDS)

WORKLOAD_ITEM_FIELDS = [
    "item_kind", "vm_name", "vm_id", "subscription", "subscription_id", "resource_group", "environment",
    "cod_app", "workload_type", "instance_name", "database_name", "protection_state", "health_status",
    "auto_protected", "vault_name", "vault_storage_redundancy", "policy_name", "schedule_frequency",
    "log_backup_frequency", "retention_max_days", "last_backup_status", "last_backup_time",
    "recovery_point_count", "item_id",
]

QUERIES = {
    "sql_servers": (
        "resources | where type =~ 'microsoft.sql/servers'"
        " | project id, name, tags, version = tostring(properties.version)"
    ),
    "sql_dbs": (
        "resources | where type =~ 'microsoft.sql/servers/databases' and name !~ 'master'"
        " | project id, name, resourceGroup, subscriptionId, location, tags, kind, sku,"
        " status = tostring(properties.status),"
        " maxSizeBytes = tolong(properties.maxSizeBytes),"
        " currentRedundancy = tostring(properties.currentBackupStorageRedundancy),"
        " requestedRedundancy = tostring(properties.requestedBackupStorageRedundancy),"
        " secondaryType = tostring(properties.secondaryType),"
        " zoneRedundant = tostring(properties.zoneRedundant),"
        " elasticPoolId = tostring(properties.elasticPoolId),"
        " earliestRestoreDate = tostring(properties.earliestRestoreDate),"
        " licenseType = tostring(properties.licenseType),"
        " haReplicas = tostring(properties.highAvailabilityReplicaCount)"
    ),
    "sql_mis": (
        "resources | where type =~ 'microsoft.sql/managedinstances'"
        " | project id, name, tags, sku,"
        " vCores = tostring(properties.vCores),"
        " storageSizeInGB = tostring(properties.storageSizeInGB),"
        " currentRedundancy = tostring(properties.currentBackupStorageRedundancy),"
        " requestedRedundancy = tostring(properties.requestedBackupStorageRedundancy),"
        " zoneRedundant = tostring(properties.zoneRedundant),"
        " licenseType = tostring(properties.licenseType)"
    ),
    "sql_mi_dbs": (
        "resources | where type =~ 'microsoft.sql/managedinstances/databases'"
        " | project id, name, resourceGroup, subscriptionId, location, tags,"
        " status = tostring(properties.status),"
        " earliestRestoreDate = tostring(properties.earliestRestorePoint)"
    ),
    "oss_servers": (
        "resources | where type in~ ('microsoft.dbforpostgresql/flexibleservers', 'microsoft.dbformysql/flexibleservers',"
        " 'microsoft.dbforpostgresql/servers', 'microsoft.dbformysql/servers')"
        " | project id, name, type, resourceGroup, subscriptionId, location, tags, sku,"
        " version = tostring(properties.version),"
        " state = tostring(coalesce(properties.state, properties.userVisibleState)),"
        " retentionDays = tostring(coalesce(properties.backup.backupRetentionDays, properties.storageProfile.backupRetentionDays)),"
        " geoBackup = tostring(coalesce(properties.backup.geoRedundantBackup, properties.storageProfile.geoRedundantBackup)),"
        " earliestRestoreDate = tostring(coalesce(properties.backup.earliestRestoreDate, properties.earliestRestoreDate)),"
        " storageGB = tostring(coalesce(properties.storage.storageSizeGB, toint(properties.storageProfile.storageMB) / 1024)),"
        " haMode = tostring(properties.highAvailability.mode),"
        " replicationRole = tostring(properties.replicationRole)"
    ),
    "cosmos": (
        "resources | where type =~ 'microsoft.documentdb/databaseaccounts'"
        " | project id, name, resourceGroup, subscriptionId, location, tags, kind,"
        " backupType = tostring(properties.backupPolicy.type),"
        " intervalMinutes = tostring(properties.backupPolicy.periodicModeProperties.backupIntervalInMinutes),"
        " retentionHours = tostring(properties.backupPolicy.periodicModeProperties.backupRetentionIntervalInHours),"
        " redundancy = tostring(properties.backupPolicy.periodicModeProperties.backupStorageRedundancy),"
        " continuousTier = tostring(properties.backupPolicy.continuousModeProperties.tier),"
        " multiWrite = tostring(properties.enableMultipleWriteLocations),"
        " locationsCount = array_length(properties.locations),"
        " zoneRedundant = tostring(properties.locations[0].isZoneRedundant),"
        " capacityMode = tostring(properties.capacityMode),"
        " capabilities = properties.capabilities,"
        " offerType = tostring(properties.databaseAccountOfferType)"
    ),
    "mongo_clusters": (
        "resources | where type =~ 'microsoft.documentdb/mongoclusters'"
        " | project id, name, resourceGroup, subscriptionId, location, tags,"
        " version = tostring(properties.serverVersion),"
        " tier = tostring(coalesce(properties.compute.tier, properties.nodeGroupSpecs[0].sku)),"
        " storageGB = tostring(coalesce(properties.storage.sizeGb, properties.nodeGroupSpecs[0].diskSizeGB)),"
        " haMode = tostring(coalesce(properties.highAvailability.targetMode, properties.nodeGroupSpecs[0].enableHa)),"
        " earliestRestoreDate = tostring(properties.backup.earliestRestoreTime)"
    ),
    "sql_vms": (
        "resources | where type =~ 'microsoft.sqlvirtualmachine/sqlvirtualmachines'"
        " | project id, name, resourceGroup, subscriptionId, location, tags,"
        " vmId = tostring(properties.virtualMachineResourceId),"
        " licenseType = tostring(properties.sqlServerLicenseType),"
        " imageOffer = tostring(properties.sqlImageOffer),"
        " imageSku = tostring(properties.sqlImageSku),"
        " management = tostring(properties.sqlManagement)"
    ),
    "sap_vis": (
        "resources | where type =~ 'microsoft.workloads/sapvirtualinstances'"
        " | project id, name, resourceGroup, subscriptionId, location, tags,"
        " sapEnvironment = tostring(properties.environment), sapProduct = tostring(properties.sapProduct),"
        " state = tostring(properties.state)"
    ),
    "vms": (
        "resources | where type =~ 'microsoft.compute/virtualmachines'"
        " | project id, name, resourceGroup, subscriptionId, location, tags,"
        " vmSize = tostring(properties.hardwareProfile.vmSize)"
    ),
}


def gb(value):
    try:
        return round(int(value) / 1024 ** 3, 1) if value not in (None, "") else ""
    except (TypeError, ValueError):
        return ""


def to_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def server_id_of(resource_id, marker):
    lowered = c.normalize(resource_id)
    position = lowered.find(marker)
    return lowered[:position] if position > 0 else ""


def merged_tags(*tag_sets):
    """Parent tags first, child tags override them."""
    merged = {}
    for tags in tag_sets:
        if isinstance(tags, dict):
            merged.update(tags)
    return merged


def is_geo(redundancy):
    value = c.normalize(redundancy)
    if not value:
        return ""
    return "YES" if value.startswith("geo") or value == "enabled" else "NO"


def arm_get_versions(ctx, stage, path, versions, params=None):
    """GET with an api-version fallback list. Logs the final failure as a gap."""
    for position, version in enumerate(versions):
        try:
            return ctx.arm_get(path, version, params)
        except c.ArmError as error:
            if error.code in FALLBACK_CODES and position + 1 < len(versions):
                continue
            ctx.errors.add(stage, path, error.url.split("?")[0].replace(c.ARM_ENDPOINT, ""),
                           error.status, error.code, error.message)
            return None
    return None


# ---------------------------------------------------------------------------
# ARM detail calls
# ---------------------------------------------------------------------------


def fetch_sql_retention(ctx, database_id):
    """Short-term (PITR) and long-term retention policies of a SQL DB or MI database."""
    short = arm_get_versions(ctx, "sql_short_term_retention",
                             f"{database_id}/backupShortTermRetentionPolicies/default", SQL_API_VERSIONS) or {}
    long = arm_get_versions(ctx, "sql_long_term_retention",
                            f"{database_id}/backupLongTermRetentionPolicies/default", SQL_API_VERSIONS) or {}
    short_props = short.get("properties") or {}
    long_props = long.get("properties") or {}
    ltr = {}
    for key in ("weeklyRetention", "monthlyRetention", "yearlyRetention"):
        value = long_props.get(key)
        ltr[key] = value if value and value.upper() not in ("PT0S", "P0D") else ""
    ltr_days = [c.iso_duration_days(value) for value in ltr.values() if value]
    return {
        "fetched_short": bool(short),
        "fetched_long": bool(long),
        "native_retention_days": short_props.get("retentionDays", ""),
        "diff_backup_interval_hours": short_props.get("diffBackupIntervalInHours", ""),
        "ltr_weekly": ltr["weeklyRetention"],
        "ltr_monthly": ltr["monthlyRetention"],
        "ltr_yearly": ltr["yearlyRetention"],
        "ltr_max_days": max(ltr_days) if ltr_days else "",
        "ltr_configured": ("YES" if ltr_days else "NO") if long else "",
    }


def fetch_sql_vm_backup(ctx, sql_vm_id):
    """Automated backup settings of a SQL IaaS extension resource. Only non-secret fields are kept."""
    body = arm_get_versions(ctx, "sql_vm_auto_backup", sql_vm_id, SQL_VM_API_VERSIONS, {"$expand": "*"})
    if body is None:
        return None
    settings = c.safe_get(body, "properties", "autoBackupSettings") or {}
    # The settings are only reported when the SQL IaaS extension runs in Full management mode.
    return {
        "sql_auto_backup_enabled": c.yes_no(settings.get("enable")) if settings else "NOT_REPORTED",
        "sql_auto_backup_retention_days": settings.get("retentionPeriod", ""),
        "sql_auto_backup_schedule": ";".join(str(value) for value in (
            settings.get("backupScheduleType"), settings.get("fullBackupFrequency"),
            f"log{settings['logBackupFrequency']}min" if settings.get("logBackupFrequency") else None) if value),
        "vm_id": c.safe_get(body, "properties", "virtualMachineResourceId", default=""),
    }


def fetch_vault_workloads(ctx, vault_id):
    """Workload containers, protected items and discovered (protectable) items of one RSV vault."""
    params = {"$filter": "backupManagementType eq 'AzureWorkload'"}
    result = {}
    for key, collection in (("containers", "backupProtectionContainers"),
                            ("protected", "backupProtectedItems"),
                            ("protectable", "backupProtectableItems")):
        result[key] = ctx.try_arm(f"vault_{key}", vault_id, ctx.arm_list,
                                  f"{vault_id}/{collection}", c.RSV_API, params) or []
    return vault_id, result


# ---------------------------------------------------------------------------
# Row builders
# ---------------------------------------------------------------------------


def empty_row(service_type):
    row = {key: "" for key in INVENTORY_FIELDS}
    row["service_type"] = service_type
    return row


def apply_protection(row, index, items, resilience):
    row.update(c.protection_columns(index, items))
    row.update(c.resilience_columns(resilience, row["id"]))


def sql_compute_model(kind, sku_tier, elastic_pool_id):
    if elastic_pool_id:
        return "ElasticPool"
    if "serverless" in str(kind).lower():
        return "Serverless"
    if str(sku_tier).lower() in ("basic", "standard", "premium"):
        return "DTU"
    return "vCore"


def build_sql_rows(ctx, data, managed_rgs, index, resilience):
    servers = {c.normalize(row["id"]): row for row in data["sql_servers"]}
    rows = []
    for db in data["sql_dbs"]:
        server_id = server_id_of(db["id"], "/databases/")
        server = servers.get(server_id, {})
        sku = db.get("sku") or {}
        is_dw = "datawarehouse" in str(db.get("kind", "")).lower() or str(sku.get("tier", "")).lower() == "datawarehouse"
        row = empty_row("SQL_DW" if is_dw else "AZURE_SQL_DB")
        source = dict(db, tags=merged_tags(server.get("tags"), db.get("tags")))
        row.update(c.base_columns(ctx, source, managed_rgs))
        row.update({
            "server_name": server_id.split("/")[-1],
            "database_name": db["name"],
            "engine_version": server.get("version", ""),
            "status": db.get("status", ""),
            "sku_name": sku.get("name", ""),
            "sku_tier": sku.get("tier", ""),
            "sku_capacity": sku.get("capacity", ""),
            "compute_model": sql_compute_model(db.get("kind"), sku.get("tier"), db.get("elasticPoolId")),
            "elastic_pool": db.get("elasticPoolId", "").split("/")[-1],
            "max_size_gb": gb(db.get("maxSizeBytes")),
            "zone_redundant": c.yes_no(db.get("zoneRedundant") == "true") if db.get("zoneRedundant") else "",
            "replica_role": db.get("secondaryType") or "Primary",
            "license_type": db.get("licenseType", ""),
            "native_backup": "YES",
            "native_backup_redundancy": db.get("currentRedundancy") or db.get("requestedRedundancy", ""),
            "native_geo_backup": is_geo(db.get("currentRedundancy") or db.get("requestedRedundancy")),
            "earliest_restore_date": c.iso(db.get("earliestRestoreDate")),
        })
        if db.get("secondaryType"):
            row["notes"] = "Geo/named replica: automated backups are taken on the primary"
        if is_dw:
            row["notes"] = "Dedicated SQL pool: restore points every 8h, 7-day retention"
        row["_needs_retention"] = not is_dw and not db.get("secondaryType")
        apply_protection(row, index, index.for_resource(db["id"]), resilience)
        rows.append(row)
    return rows


def build_mi_rows(ctx, data, managed_rgs, index, resilience):
    instances = {c.normalize(row["id"]): row for row in data["sql_mis"]}
    rows = []
    for db in data["sql_mi_dbs"]:
        instance_id = server_id_of(db["id"], "/databases/")
        instance = instances.get(instance_id, {})
        sku = instance.get("sku") or {}
        row = empty_row("SQL_MI_DB")
        row.update(c.base_columns(ctx, dict(db, tags=merged_tags(instance.get("tags"), db.get("tags"))), managed_rgs))
        redundancy = instance.get("currentRedundancy") or instance.get("requestedRedundancy", "")
        row.update({
            "server_name": instance_id.split("/")[-1],
            "database_name": db["name"],
            "status": db.get("status", ""),
            "sku_name": sku.get("name", ""),
            "sku_tier": sku.get("tier", ""),
            "sku_capacity": instance.get("vCores", ""),
            "compute_model": "vCore",
            "storage_gb": instance.get("storageSizeInGB", ""),
            "zone_redundant": c.yes_no(instance.get("zoneRedundant") == "true") if instance.get("zoneRedundant") else "",
            "license_type": instance.get("licenseType", ""),
            "native_backup": "YES",
            "native_backup_redundancy": redundancy,
            "native_geo_backup": is_geo(redundancy),
            "earliest_restore_date": c.iso(db.get("earliestRestoreDate")),
            "notes": "Storage and vCores are per instance, shared by its databases",
        })
        row["_needs_retention"] = True
        apply_protection(row, index, index.for_resource(db["id"]), resilience)
        rows.append(row)
    return rows


OSS_TYPES = {
    "microsoft.dbforpostgresql/flexibleservers": "POSTGRES_FLEX",
    "microsoft.dbformysql/flexibleservers": "MYSQL_FLEX",
    "microsoft.dbforpostgresql/servers": "POSTGRES_SINGLE",
    "microsoft.dbformysql/servers": "MYSQL_SINGLE",
}


def build_oss_rows(ctx, data, managed_rgs, index, resilience):
    rows = []
    for server in data["oss_servers"]:
        service = OSS_TYPES[server["type"].lower()]
        sku = server.get("sku") or {}
        role = server.get("replicationRole", "")
        row = empty_row(service)
        row.update(c.base_columns(ctx, server, managed_rgs))
        is_replica = role.lower() in ("replica", "asyncreplica", "georeplica")
        row.update({
            "server_name": server["name"],
            "engine_version": server.get("version", ""),
            "status": server.get("state", ""),
            "sku_name": sku.get("name", ""),
            "sku_tier": sku.get("tier", ""),
            "sku_capacity": sku.get("capacity", ""),
            "compute_model": "vCore",
            "storage_gb": server.get("storageGB", ""),
            "ha_mode": server.get("haMode", ""),
            "zone_redundant": "YES" if server.get("haMode", "").lower() == "zoneredundant" else "NO",
            "replica_role": role,
            "native_backup": "NO" if is_replica else "YES",
            "native_retention_days": server.get("retentionDays", ""),
            "native_backup_interval": "Daily snapshot + transaction logs",
            "native_backup_redundancy": "Geo" if server.get("geoBackup", "").lower() == "enabled" else "Local/Zone",
            "native_geo_backup": is_geo(server.get("geoBackup")),
            "earliest_restore_date": c.iso(server.get("earliestRestoreDate")),
        })
        if service.endswith("_SINGLE"):
            row["notes"] = "Single server is retired; migrate to flexible server"
        if is_replica:
            row["notes"] = "Read replica: backups are not taken on replicas"
        apply_protection(row, index, index.for_resource(server["id"], include_children=True), resilience)
        rows.append(row)
    return rows


def build_cosmos_rows(ctx, data, managed_rgs, index, resilience):
    rows = []
    for account in data["cosmos"]:
        capabilities = {c.normalize(item.get("name")) for item in account.get("capabilities") or [] if isinstance(item, dict)}
        continuous = account.get("backupType", "").lower() == "continuous"
        tier = account.get("continuousTier", "")
        retention_hours = to_float(account.get("retentionHours"))
        if continuous:
            retention_days = 30 if "30" in tier else 7
            interval = f"Continuous ({tier or 'Continuous30Days'})"
        else:
            retention_days = round(retention_hours / 24, 2) if retention_hours is not None else ""
            interval = f"Every {account.get('intervalMinutes')}min" if account.get("intervalMinutes") else ""
        redundancy = account.get("redundancy") or ("Geo/Region" if continuous else "")
        row = empty_row("COSMOS_DB")
        row.update(c.base_columns(ctx, account, managed_rgs))
        row.update({
            "server_name": account["name"],
            "engine_version": account.get("kind", ""),
            "sku_name": account.get("offerType", ""),
            "compute_model": "Serverless" if "enableserverless" in capabilities else (account.get("capacityMode") or "Provisioned"),
            "zone_redundant": c.yes_no(account.get("zoneRedundant") == "true") if account.get("zoneRedundant") else "",
            "locations_count": account.get("locationsCount", ""),
            "multi_region_write": c.yes_no(account.get("multiWrite") == "true") if account.get("multiWrite") else "",
            "native_backup": "YES",
            "native_retention_days": retention_days,
            "native_backup_interval": interval,
            "native_backup_redundancy": redundancy,
            "native_geo_backup": "YES" if continuous else is_geo(account.get("redundancy")),
            "notes": "Continuous backup is billed per GB-month; periodic keeps 2 free copies" if continuous else "",
        })
        apply_protection(row, index, index.for_resource(account["id"]), resilience)
        rows.append(row)

    for cluster in data["mongo_clusters"]:
        row = empty_row("COSMOS_MONGO_VCORE")
        row.update(c.base_columns(ctx, cluster, managed_rgs))
        row.update({
            "server_name": cluster["name"],
            "engine_version": cluster.get("version", ""),
            "sku_name": cluster.get("tier", ""),
            "compute_model": "vCore",
            "storage_gb": cluster.get("storageGB", ""),
            "ha_mode": cluster.get("haMode", ""),
            "native_backup": "YES",
            "native_retention_days": 35,
            "native_backup_interval": "Automatic",
            "earliest_restore_date": c.iso(cluster.get("earliestRestoreDate")),
        })
        apply_protection(row, index, index.for_resource(cluster["id"]), resilience)
        rows.append(row)
    return rows


def vm_engine(workload_type):
    value = c.normalize(workload_type)
    if value in SAP_HANA_WORKLOADS:
        return "SAP_HANA_VM"
    if value in SAP_ASE_WORKLOADS:
        return "SAP_ASE_VM"
    if value.startswith("sql"):
        return "SQL_SERVER_VM"
    return ""


def build_vm_rows(ctx, data, managed_rgs, index, resilience, vault_workloads, sql_vm_details):
    """One row per (VM, engine) for SQL Server, SAP HANA and SAP ASE running in VMs."""
    vms = {c.normalize(vm["id"]): vm for vm in data["vms"]}
    vms_by_rg_name = {(c.normalize(vm["resourceGroup"]), c.normalize(vm["name"])): vm["id"] for vm in data["vms"]}

    container_vm = {}
    for vault_id, payload in vault_workloads.items():
        for container in payload.get("containers", []):
            source = c.safe_get(container, "properties", "sourceResourceId")
            if source:
                container_vm[c.normalize(container.get("name"))] = c.normalize(source)

    def vm_from_container(item_id):
        lowered = c.normalize(item_id)
        if "/protectioncontainers/" not in lowered:
            return ""
        container = lowered.split("/protectioncontainers/")[1].split("/")[0]
        if container in container_vm:
            return container_vm[container]
        parts = container.split(";")
        if len(parts) >= 4:
            return c.normalize(vms_by_rg_name.get((parts[-2], parts[-1]), ""))
        return ""

    groups = defaultdict(lambda: {"protected": [], "protectable": [], "sql_vm": None})
    for item in index.items:
        if c.normalize(item["workload_type"]) in DB_WORKLOADS and item["backup_management_type"] == "AzureWorkload":
            engine = vm_engine(item["workload_type"])
            vm_id = item["source_resource_id"] or vm_from_container(item["item_id"])
            groups[(vm_id, engine)]["protected"].append(item)

    protectable_rows = []
    for vault_id, payload in vault_workloads.items():
        for item in payload.get("protectable", []):
            properties = item.get("properties") or {}
            workload = properties.get("protectableItemType") or properties.get("workloadType", "")
            if c.normalize(workload) not in {"sqldatabase", "saphanadatabase", "sapasedatabase"}:
                continue
            if c.normalize(properties.get("protectionState")) == "protected":
                continue
            vm_id = vm_from_container(item.get("id"))
            engine = vm_engine(workload)
            record = {"vault_id": c.normalize(vault_id), "workload": workload, "properties": properties,
                      "id": item.get("id", ""), "vm_id": vm_id}
            groups[(vm_id, engine)]["protectable"].append(record)
            protectable_rows.append(record)

    for sql_vm in data["sql_vms"]:
        detail = sql_vm_details.get(c.normalize(sql_vm["id"])) or {}
        vm_id = c.normalize(sql_vm.get("vmId") or detail.get("vm_id")) or c.normalize(
            vms_by_rg_name.get((c.normalize(sql_vm["resourceGroup"]), c.normalize(sql_vm["name"])), ""))
        groups[(vm_id, "SQL_SERVER_VM")]["sql_vm"] = (sql_vm, detail)

    for sap in data["sap_vis"]:
        groups[(c.normalize(sap["id"]), "SAP_VIS")]["sap_vis"] = sap

    rows = []
    for (vm_id, engine), group in sorted(groups.items()):
        vm = vms.get(vm_id)
        sql_vm, detail = group.get("sql_vm") or (None, {})
        sap = group.get("sap_vis")
        if sap:
            source = sap
        elif vm:
            source = dict(vm, tags=merged_tags(vm.get("tags"), (sql_vm or {}).get("tags")))
        else:
            parts = c.parse_resource_id(vm_id)
            source = {"id": vm_id, "name": parts["name"], "resourceGroup": parts["resource_group"],
                      "subscriptionId": parts["subscription_id"], "location": "", "tags": (sql_vm or {}).get("tags")}
        row = empty_row(engine or "SQL_SERVER_VM")
        row.update(c.base_columns(ctx, source, managed_rgs))

        vm_items = [item for item in index.for_resource(vm_id) if item["backup_management_type"] == "AzureIaasVM"]
        vm_active = [item for item in vm_items if c.item_is_active(item)]
        workload_items = group["protected"]
        active_workload = [item for item in workload_items if c.item_is_active(item)]
        auto_backup = detail.get("sql_auto_backup_enabled", "") if sql_vm else ""

        row.update({
            "server_name": source.get("name", ""),
            "vm_size": (vm or {}).get("vmSize", ""),
            "license_type": (sql_vm or {}).get("licenseType", ""),
            "sku_name": (sql_vm or {}).get("imageOffer", "") or (sap or {}).get("sapProduct", ""),
            "sku_tier": (sql_vm or {}).get("imageSku", ""),
            "compute_model": (sql_vm or {}).get("management", "") or "IaaS",
            "status": "" if vm or sap else "VM_NOT_FOUND",
            "native_backup": "NOT_APPLICABLE",
            "sql_auto_backup_enabled": auto_backup,
            "sql_auto_backup_retention_days": detail.get("sql_auto_backup_retention_days", ""),
            "sql_auto_backup_schedule": detail.get("sql_auto_backup_schedule", ""),
            "workload_items_protected": len(active_workload),
            "workload_items_unprotected": len(group["protectable"]),
            "vm_backup_enabled": "YES" if vm_active else "NO",
            "vm_backup_vault": ";".join(sorted({item["vault_id"].split("/")[-1] for item in vm_active})),
        })
        if sql_vm and ctx.args.skip_arm:
            row["sql_auto_backup_enabled"] = "UNKNOWN"
        if row["sql_auto_backup_enabled"] == "NOT_REPORTED":
            row["notes"] = "Automated backup settings need the SQL IaaS extension in Full mode; check the VM directly"
        if sap:
            row["notes"] = f"SAP virtual instance ({sap.get('sapEnvironment', '')}); VM coverage is in the SAP_HANA_VM/SAP_ASE_VM rows"
        apply_protection(row, index, workload_items or vm_items, resilience)
        row["_vm_id"] = vm_id
        rows.append(row)
    return rows, protectable_rows


# ---------------------------------------------------------------------------
# Backup method and findings
# ---------------------------------------------------------------------------


def assign_method_and_findings(row, index, args):
    service = row["service_type"]
    active_vault = int(row.get("backup_active_count") or 0) > 0
    if service in ("SQL_SERVER_VM", "SAP_HANA_VM", "SAP_ASE_VM", "SAP_VIS"):
        if int(row.get("workload_items_protected") or 0) > 0:
            method = "VAULT_RSV_WORKLOAD"
        elif row.get("sql_auto_backup_enabled") == "YES":
            method = "SQL_AUTOMATED_BACKUP"
        elif row.get("vm_backup_enabled") == "YES":
            method = "VM_SNAPSHOT_ONLY"
        else:
            method = "NONE"
    elif row.get("native_backup") == "NO" and "replica" in str(row.get("notes", "")).lower():
        method = "NOT_APPLICABLE_REPLICA"
    elif str(row.get("replica_role", "")).lower() not in ("", "primary") and service in ("AZURE_SQL_DB",):
        method = "NOT_APPLICABLE_REPLICA"
    elif active_vault:
        method = "NATIVE+VAULT_BV"
    else:
        method = "NATIVE"
    row["backup_method"] = method

    findings = c.Findings()
    environment = row["environment"]
    is_prod = environment == "PROD"
    is_nonprod = environment in ("CERT", "DESA")

    if row["backup_scope"] == "EXCLUDED_PLATFORM":
        findings.add("EXCLUDED_PLATFORM", "INFO")
    if environment == "UNKNOWN":
        findings.add("REQUIRES_ENVIRONMENT_CLASSIFICATION", "MEDIUM")

    if active_vault and any(not c.item_is_healthy(item, args.stale_hours)
                            for item in index.for_resource(row.get("_vm_id") or row["id"], include_children=True)
                            if c.item_is_active(item) and item["backup_management_type"] != "AzureIaasVM"):
        findings.add("DB_BACKUP_UNHEALTHY", "HIGH" if is_prod else "MEDIUM")

    if method in ("NATIVE", "NATIVE+VAULT_BV") and is_prod:
        retention = to_float(row.get("native_retention_days"))
        if retention is not None and retention < args.min_prod_retention_days:
            findings.add("PROD_RETENTION_BELOW_MIN", "HIGH")
        if row.get("native_geo_backup") == "NO":
            findings.add("PROD_BACKUP_NOT_GEO_REDUNDANT", "MEDIUM")

    if service in ("SQL_SERVER_VM", "SAP_HANA_VM", "SAP_ASE_VM") and is_prod:
        if method == "NONE":
            findings.add("PROD_DB_WORKLOAD_WITHOUT_BACKUP", "HIGH")
        elif method == "VM_SNAPSHOT_ONLY":
            findings.add("PROD_DB_WORKLOAD_WITHOUT_BACKUP", "MEDIUM")
    if int(row.get("workload_items_unprotected") or 0) > 0 and is_prod:
        findings.add("PROD_DB_DISCOVERED_NOT_PROTECTED", "HIGH")

    if is_nonprod and (row.get("ltr_configured") == "YES" or active_vault):
        findings.add("NONPROD_LONG_TERM_RETENTION", "LOW")
    if is_nonprod and row.get("native_geo_backup") == "YES":
        findings.add("NONPROD_GEO_BACKUP_REDUNDANCY", "LOW")
    if method == "NATIVE+VAULT_BV" or (method == "VAULT_RSV_WORKLOAD" and row.get("sql_auto_backup_enabled") == "YES"):
        findings.add("DUAL_PROTECTION", "LOW")

    row.update(findings.columns())


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    parser = c.build_parser(__doc__.splitlines()[0], c.results_path("backup-database"))
    parser.add_argument("--min-prod-retention-days", type=float, default=7,
                        help="PROD native retention below this is a finding")
    args = parser.parse_args()
    c.setup_logging(args.verbose)
    ctx = c.AzureContext(args)
    ctx.load_subscriptions()
    if not ctx.subscriptions:
        c.LOG.error("No subscriptions visible. Sign in with an account that has Reader access.")
        return 1

    with ThreadPoolExecutor(max_workers=4) as pool:
        index_future = pool.submit(c.load_backup_index, ctx)
        managed_future = pool.submit(c.load_managed_resource_groups, ctx)
        resilience_future = pool.submit(c.load_resilience, ctx, RESILIENCE_TYPES)
        data_future = pool.submit(ctx.queries, QUERIES)
        index, managed_rgs = index_future.result(), managed_future.result()
        resilience, data = resilience_future.result(), data_future.result()

    graph_workload_items = sum(1 for item in index.items if item["backup_management_type"] == "AzureWorkload")
    vault_workloads = {}
    sql_vm_details = {}
    arm_workload_items_added = 0
    if not args.skip_arm:
        rsv_vaults = [vault["id"] for vault in index.vaults.values() if vault["kind"] == "RECOVERY_SERVICES_VAULT"]
        for result in ctx.parallel(lambda vault_id: fetch_vault_workloads(ctx, vault_id), rsv_vaults,
                                   label="vault_workloads"):
            if result:
                vault_workloads[c.normalize(result[0])] = result[1]
        # Resource Graph may not index workload items; merge the ones only ARM returned.
        known = {c.normalize(item["item_id"]) for item in index.items}
        for payload in vault_workloads.values():
            for raw in payload["protected"]:
                if c.normalize(raw.get("id")) in known:
                    continue
                index.add_item(c._rsv_item(raw, index))  # same schema as the Resource Graph row
                arm_workload_items_added += 1
        for sql_vm, detail in zip(data["sql_vms"], ctx.parallel(
                lambda row: fetch_sql_vm_backup(ctx, row["id"]), data["sql_vms"], label="sql_vm_auto_backup")):
            if detail:
                sql_vm_details[c.normalize(sql_vm["id"])] = detail

    rows = []
    rows += build_sql_rows(ctx, data, managed_rgs, index, resilience)
    rows += build_mi_rows(ctx, data, managed_rgs, index, resilience)
    rows += build_oss_rows(ctx, data, managed_rgs, index, resilience)
    rows += build_cosmos_rows(ctx, data, managed_rgs, index, resilience)
    vm_rows, protectable = build_vm_rows(ctx, data, managed_rgs, index, resilience, vault_workloads, sql_vm_details)
    rows += vm_rows

    if not args.skip_arm:
        targets = [row for row in rows if row.pop("_needs_retention", False)]
        for row, detail in zip(targets, ctx.parallel(lambda row: fetch_sql_retention(ctx, row["id"]), targets,
                                                     label="sql_retention")):
            if detail:
                row.update({key: value for key, value in detail.items() if key in row})
    for row in rows:
        row.pop("_needs_retention", None)
        assign_method_and_findings(row, index, args)
        row.pop("_vm_id", None)
    rows.sort(key=lambda row: (row["service_type"], row["subscription"], row["resource_group"], row["name"]))

    workload_rows = build_workload_item_rows(ctx, index, protectable, managed_rgs, data)

    existing = {c.normalize(row["id"]) for key in ("sql_dbs", "sql_mi_dbs", "oss_servers", "cosmos", "vms")
                for row in data[key]}
    for item in index.items:
        if any(parent in existing for parent in c.parent_resource_ids(item["source_resource_id"])):
            existing.add(item["source_resource_id"])
    orphans = c.orphan_item_rows(index, existing, is_db_item)

    policies = c.policy_rows(index, workloads=["SQLDataBase", "SAPHana", "SAPAse", "PostgreSQL", "MySQL"])
    used = defaultdict(int)
    for item in index.items:
        used[item["policy_id"]] += 1
    for policy in policies:
        count = policy.get("protected_items_count")
        in_use = used.get(c.normalize(policy["policy_id"]), 0)
        policy["items_in_index"] = in_use
        policy["unused_policy"] = "YES" if not in_use and str(count or "0") in ("0", "") else "NO"

    db_vault_ids = {item["vault_id"] for item in index.items if is_db_item(item)}
    vaults = [row for row in c.vault_rows(index) if c.normalize(row["vault_id"]) in db_vault_ids]

    output = c.Output(args.output_dir, args.delimiter)
    output.csv("database_inventory.csv", rows, INVENTORY_FIELDS)
    output.csv("database_vm_workload_items.csv", workload_rows, WORKLOAD_ITEM_FIELDS)
    output.csv("database_orphan_backup_items.csv", orphans, c.ORPHAN_FIELDS)
    output.csv("database_backup_policies.csv", policies)
    output.csv("database_backup_vaults.csv", vaults)
    output.csv("database_errors.csv", ctx.errors.rows, c.ErrorLog.FIELDS)
    output.json("database_summary.json", c.run_metadata(ctx, {
        "rows": len(rows),
        "service_type_by_backup_method": c.count_by(rows, "service_type", "backup_method"),
        "environment_by_finding": c.count_by(rows, "environment", "finding"),
        "service_type_by_environment": c.count_by(rows, "service_type", "environment"),
        "native_redundancy_by_environment": c.count_by(rows, "environment", "native_backup_redundancy"),
        "workload_items": {
            "resource_graph": graph_workload_items,
            "added_from_arm": arm_workload_items_added,
            "discovered_not_protected": len(protectable),
            "rsv_vaults_checked": len(vault_workloads),
        },
        "orphan_backup_items": c.count_by(orphans, "orphan_reason", "workload_type"),
        "unused_db_policies": sum(1 for policy in policies if policy["unused_policy"] == "YES"),
    }))
    if args.xlsx:
        output.xlsx("database_inventory.xlsx")
    if ctx.errors.rows:
        c.LOG.warning("Partial result: %d API errors or access gaps, see database_errors.csv", len(ctx.errors.rows))
    return 0


def is_db_item(item):
    if c.normalize(item["workload_type"]) in DB_WORKLOADS:
        return True
    return any(marker in c.normalize(item["datasource_type"]) for marker in DB_DATASOURCE_MARKERS)


def build_workload_item_rows(ctx, index, protectable, managed_rgs, data):
    vms = {c.normalize(vm["id"]): vm for vm in data["vms"]}

    def vm_columns(vm_id):
        vm = vms.get(vm_id)
        if vm:
            base = c.base_columns(ctx, vm, managed_rgs)
            return {"vm_name": vm["name"], "vm_id": vm["id"], "subscription": base["subscription"],
                    "subscription_id": base["subscription_id"], "resource_group": base["resource_group"],
                    "environment": base["environment"], "cod_app": base["cod_app"]}
        parts = c.parse_resource_id(vm_id)
        return {"vm_name": parts["name"], "vm_id": vm_id, "subscription": ctx.subscription_name(parts["subscription_id"]),
                "subscription_id": parts["subscription_id"], "resource_group": parts["resource_group"],
                "environment": "UNKNOWN", "cod_app": ""}

    rows = []
    for item in index.items:
        if item["backup_management_type"] != "AzureWorkload" or c.normalize(item["workload_type"]) not in DB_WORKLOADS:
            continue
        vault = index.vaults.get(item["vault_id"], {})
        policy = index.policies.get(item["policy_id"], {})
        row = {"item_kind": "PROTECTED"}
        row.update(vm_columns(item["source_resource_id"]))
        row.update({
            "workload_type": item["workload_type"],
            "instance_name": item.get("parent_name") or item.get("server_name", ""),
            "database_name": item["protected_item_name"],
            "protection_state": item["protection_state"],
            "health_status": item["health_status"],
            "vault_name": vault.get("name") or item["vault_id"].split("/")[-1],
            "vault_storage_redundancy": vault.get("storage_redundancy", ""),
            "policy_name": policy.get("policy_name") or item.get("policy_name_raw", ""),
            "schedule_frequency": policy.get("schedule_frequency", ""),
            "log_backup_frequency": policy.get("log_backup_frequency", ""),
            "retention_max_days": policy.get("retention_max_days", ""),
            "last_backup_status": item["last_backup_status"],
            "last_backup_time": item["last_backup_time"],
            "recovery_point_count": item["recovery_point_count"],
            "item_id": item["item_id"],
        })
        rows.append(row)
    for record in protectable:
        properties = record["properties"]
        vault = index.vaults.get(record["vault_id"], {})
        row = {"item_kind": "DISCOVERED_NOT_PROTECTED"}
        row.update(vm_columns(record["vm_id"]))
        row.update({
            "workload_type": record["workload"],
            "instance_name": properties.get("parentName") or properties.get("serverName", ""),
            "database_name": properties.get("friendlyName", ""),
            "protection_state": properties.get("protectionState", ""),
            "auto_protected": c.yes_no(properties.get("isAutoProtected")) if "isAutoProtected" in properties else "",
            "vault_name": vault.get("name") or record["vault_id"].split("/")[-1],
            "vault_storage_redundancy": vault.get("storage_redundancy", ""),
            "item_id": record["id"],
        })
        rows.append(row)
    rows.sort(key=lambda row: (row["vm_name"].lower(), row["item_kind"], str(row["database_name"]).lower()))
    return rows


if __name__ == "__main__":
    sys.exit(main())
