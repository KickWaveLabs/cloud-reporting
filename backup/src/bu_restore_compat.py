#!/usr/bin/env python3
"""Long-term restore compatibility inputs: OS versions, VM generation, customer-managed disk keys,
database engine versions and AKS versions, for the resources that matter to backup retention.

Outputs:
  restore_vm_versions.csv     one row per VM: OS, version, generation, security type, disk encryption set
  restore_engine_versions.csv PostgreSQL / MySQL flexible servers and AKS clusters with their versions
  versions.json               counts used by bu_insights_workbook.py (--versions)
  restore_summary.json        run metadata and evidence (queries executed)

Dependencies:
  pip install azure-identity azure-mgmt-resourcegraph requests
"""

import json
import os

import bu_common as common

# Community end of support used to flag engines and clusters. Update when versions retire.
EOL_POSTGRES_MAJOR = {"11", "12", "13"}
EOL_MYSQL_PREFIX = ("5.7",)
MIN_SUPPORTED_AKS_MINOR = 32


def main():
    parser = common.build_parser(__doc__.splitlines()[0], common.results_path("backup-restore"))
    parser.add_argument("--compute-inventory", default=common.results_path("backup-compute", "compute_vm_inventory.csv"),
                        help="compute_vm_inventory.csv, to count only VMs with an active backup")
    args = parser.parse_args()
    common.setup_logging(args.verbose)
    ctx = common.AzureContext(args)
    ctx.load_subscriptions()

    results = ctx.queries({
        "vm_versions": (
            "resources | where type =~ 'microsoft.compute/virtualmachines'"
            " | project id = tolower(id), name, resourceGroup, subscriptionId,"
            " osName = tostring(properties.extended.instanceView.osName),"
            " osVersion = tostring(properties.extended.instanceView.osVersion),"
            " generation = tostring(properties.extended.instanceView.hyperVGeneration),"
            " securityType = tostring(properties.securityProfile.securityType),"
            " diskEncryptionSet = tostring(properties.storageProfile.osDisk.managedDisk.diskEncryptionSet.id)"
        ),
        "engine_versions": (
            "resources | where type in~ ('microsoft.dbforpostgresql/flexibleservers', 'microsoft.dbformysql/flexibleservers',"
            " 'microsoft.containerservice/managedclusters')"
            " | project id = tolower(id), name, type, resourceGroup, subscriptionId,"
            " version = iff(type =~ 'microsoft.containerservice/managedclusters',"
            " tostring(properties.currentKubernetesVersion), tostring(properties.version))"
        ),
    })

    protected = set()
    if os.path.exists(args.compute_inventory):
        for row in common_csv(args.compute_inventory):
            if row.get("backup_enabled") == "YES":
                protected.add(row["id"].lower())
    else:
        common.LOG.warning("Compute inventory not found; counts cover every VM, not only VMs with backup")

    vm_rows = []
    for row in results["vm_versions"]:
        vm_rows.append({
            "subscription": ctx.subscription_name(row["subscriptionId"]), "resource_group": row["resourceGroup"],
            "name": row["name"], "os_name": row["osName"], "os_version": row["osVersion"],
            "generation": row["generation"], "security_type": row["securityType"] or "Standard",
            "customer_managed_key": "YES" if row["diskEncryptionSet"] else "NO",
            "backup_enabled": "YES" if row["id"] in protected else ("NO" if protected else "UNKNOWN"),
            "id": row["id"],
        })

    def flag(row):
        kind = row["type"].lower()
        version = row["version"] or ""
        if "dbforpostgresql" in kind:
            return "OUT_OF_SUPPORT" if version.split(".")[0] in EOL_POSTGRES_MAJOR else "SUPPORTED"
        if "dbformysql" in kind:
            return "OUT_OF_SUPPORT" if version.startswith(EOL_MYSQL_PREFIX) else "SUPPORTED"
        try:
            return "OUT_OF_SUPPORT" if int(version.split(".")[1]) < MIN_SUPPORTED_AKS_MINOR else "SUPPORTED"
        except (IndexError, ValueError):
            return "UNKNOWN"

    engine_rows = [{
        "subscription": ctx.subscription_name(row["subscriptionId"]), "resource_group": row["resourceGroup"],
        "name": row["name"], "type": row["type"], "version": row["version"], "support": flag(row), "id": row["id"],
    } for row in results["engine_versions"]]

    scope = [row for row in vm_rows if row["backup_enabled"] == "YES"] if protected else vm_rows
    counts = {
        "vms_with_backup": len(scope),
        "gen1_with_backup": sum(1 for row in scope if row["generation"] == "V1"),
        "cmk_vms_with_backup": sum(1 for row in scope if row["customer_managed_key"] == "YES"),
        "cmk_vms": sum(1 for row in vm_rows if row["customer_managed_key"] == "YES"),
        "eol_db_engines": sum(1 for row in engine_rows if row["support"] == "OUT_OF_SUPPORT"
                              and "containerservice" not in row["type"].lower()),
        "aks_unsupported": sum(1 for row in engine_rows if row["support"] == "OUT_OF_SUPPORT"
                               and "containerservice" in row["type"].lower()),
    }

    output = common.Output(args.output_dir, args.delimiter)
    output.csv("restore_vm_versions.csv", vm_rows)
    output.csv("restore_engine_versions.csv", engine_rows)
    with open(os.path.join(args.output_dir, "versions.json"), "w", encoding="utf-8") as handle:
        json.dump(counts, handle, indent=2)
    output.csv("restore_errors.csv", ctx.errors.rows, common.ErrorLog.FIELDS)
    output.json("restore_summary.json", common.run_metadata(ctx, {"counts": counts}))


def common_csv(path):
    import csv
    with open(path, encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


if __name__ == "__main__":
    main()
