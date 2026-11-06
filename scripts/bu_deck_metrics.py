#!/usr/bin/env python3
"""Export the figures used by the backup slide decks to one JSON file.

Reads the inventory CSVs and the recalculated backup_insights.xlsx (opportunity savings,
risk and long-term restore counts), so the decks and the workbook show the same numbers.
"""

import argparse
import csv
import json
import os
from collections import Counter

from openpyxl import load_workbook


def read_csv(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def sheet_rows(workbook, name, first_row):
    return [row for row in workbook[name].iter_rows(min_row=first_row, values_only=True) if row[0]]


def as_int(value):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return 0


def third_party(workbook):
    """Scenario costs, measured volumes and RPO/RTO rows from the Comparacion_terceros sheet."""
    if "Comparacion_terceros" not in workbook.sheetnames:
        return {}
    sheet = workbook["Comparacion_terceros"]
    values, scenarios, rpo, restore = {}, [], [], []
    section = None
    for row in sheet.iter_rows(values_only=True):
        first = row[0]
        if first == "Escenario":
            section = "scenarios"
            continue
        if first == "Workload":
            section = "rpo"
            continue
        if isinstance(first, str) and first.startswith("Restauración desde"):
            section = "restore"
            continue
        if not first:
            if section in ("scenarios", "restore"):
                section = None
            continue
        if section == "scenarios" and isinstance(row[1], (int, float)):
            scenarios.append({"name": first, "tb": round(row[1] or 0, 1), "monthly": round(row[2] or 0),
                              "delta": round(row[4] or 0), "per_tb": round(row[5] or 0, 2),
                              "breakeven": row[6] if isinstance(row[6], (int, float)) else None})
        elif section == "rpo" and row[1]:
            rpo.append({"workload": first, "rpo": row[1], "rto": row[2], "ltr": row[3], "replaces": row[4], "note": row[5]})
        elif section == "restore" and isinstance(row[1], (int, float)):
            restore.append({"name": first, "tb": row[1], "hours": round(row[2] or 0, 1)})
        elif isinstance(row[2], (int, float)) and first and not section:
            values[first] = row[2]
    return {"values": values, "scenarios": scenarios, "rpo": rpo, "restore": restore}


def meter_group(meter):
    lowered = meter.lower()
    if "data stored" in lowered:
        return "Dato almacenado"
    if "transfer" in lowered or "geo-replication" in lowered:
        return "Transferencia y geo-replicación"
    return "Operaciones y transacciones"


def top_storage_accounts(pattern, storage, limit=6, storage_path=""):
    """Most expensive storage accounts with cost split by stored data, operations and transfer."""
    import glob
    costs = {}
    csv_path = os.path.join(os.path.dirname(os.path.dirname(storage_path)), "backup-cost", "cost_by_storage_account.csv") if storage_path else ""
    for row in read_csv(csv_path) if csv_path else []:
        name = row["resource_id"].lower().split("/storageaccounts/")[-1].split("/")[0]
        costs.setdefault(name, Counter())[meter_group(row["meter"])] += float(row["cost_usd"] or 0)
    for path in (glob.glob(pattern) if pattern and not costs else []):
        with open(path, encoding="utf-8") as handle:
            body = json.load(handle)
        columns = [column["name"] for column in body["properties"]["columns"]]
        for values in body["properties"]["rows"]:
            record = dict(zip(columns, values))
            resource_id = str(record["ResourceId"]).lower()
            name = resource_id.split("/storageaccounts/")[-1].split("/")[0]
            entry = costs.setdefault(name, Counter())
            entry[meter_group(record["Meter"])] += record["CostUSD"]
    info = {row["name"].lower(): row for row in storage}
    top = sorted(costs.items(), key=lambda pair: -sum(pair[1].values()))[:limit]
    return [{"name": name, "total": round(sum(parts.values())),
             "parts": {key: round(value) for key, value in parts.items()},
             "environment": info.get(name, {}).get("environment", ""), "sku": info.get(name, {}).get("sku", ""),
             "backup_method": info.get(name, {}).get("backup_method", "")} for name, parts in top]


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--results", default="./results")
    parser.add_argument("--workbook", required=True, help="Recalculated backup_insights.xlsx")
    parser.add_argument("--account-costs", default="", help="Glob of per-storage-account cost JSON files (optional)")
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    base = args.results
    vm = read_csv(f"{base}/backup-compute/compute_vm_inventory.csv")
    k8s = read_csv(f"{base}/backup-compute/compute_k8s_inventory.csv")
    orphans = read_csv(f"{base}/backup-compute/compute_orphan_backup_items.csv")
    vaults = read_csv(f"{base}/backup-compute/compute_backup_vaults.csv")
    policies = read_csv(f"{base}/backup-compute/compute_backup_policies.csv")
    db = read_csv(f"{base}/backup-database/database_inventory.csv")
    native = read_csv(f"{base}/backup-native/native_protection_inventory.csv")
    plans = read_csv(f"{base}/backup-native/native_app_service_plans.csv")
    storage = read_csv(f"{base}/backup-storage/storage_inventory.csv")
    snapshots = read_csv(f"{base}/backup-storage/storage_snapshots_by_source.csv")
    workbook = load_workbook(args.workbook, data_only=True)

    traditional = [row for row in vm if row["infrastructure_class"] == "TRADITIONAL"]
    active = [row for row in vm if row["backup_enabled"] == "YES"]
    coverage = {}
    for environment in ("PROD", "CERT", "DESA"):
        rows = [row for row in traditional if row["environment"] == environment]
        coverage[environment] = {"with": sum(1 for row in rows if row["backup_enabled"] == "YES"),
                                 "without": sum(1 for row in rows if row["backup_enabled"] != "YES")}

    def retention_bucket(value):
        days = as_int(value)
        if days <= 35:
            return "≤ 35 días"
        if days <= 400:
            return "~1 año"
        return "4 a 12 años"

    opportunities = [{"id": row[0], "front": row[1], "title": row[2], "count": row[3],
                      "monthly": round(row[4] or 0), "annual": round(row[5] or 0), "confidence": row[6],
                      "change": row[7], "risk": row[8]}
                     for row in sheet_rows(workbook, "Oportunidades", 5) if str(row[0]).startswith("O")]
    risks = [{"id": row[0], "front": row[1], "title": row[2], "count": row[3], "severity": row[4], "action": row[5]}
             for row in sheet_rows(workbook, "Riesgos", 4) if str(row[0]).startswith("R")]
    long_term = [{"id": row[0], "title": row[1], "count": row[2], "why": row[3], "mitigation": row[4]}
                 for row in sheet_rows(workbook, "Restauracion_LP", 4) if str(row[0]).startswith("L")]
    calibration = next(workbook["Costo_calibrado"].iter_rows(min_row=2, max_row=2, values_only=True))
    vault_cost_rows = [row for row in workbook["Costo_calibrado"].iter_rows(min_row=3, values_only=True)
                       if isinstance(row[2], (int, float))]
    orphan_vault_storage = sum(row[4] or 0 for row in vault_cost_rows if row[2] and row[3] == row[2])
    assumptions = {row[0]: row[2] for row in workbook["Supuestos"].iter_rows(min_row=4, values_only=True) if row[0]}

    sql = [row for row in db if row["service_type"] == "AZURE_SQL_DB"]
    redis = [row for row in native if row["service_type"] == "REDIS"]
    key_vaults = [row for row in native if row["service_type"] == "KEY_VAULT"]
    pi_orphans = len(orphans) * (assumptions.get("pi_usd") or 0)

    metrics = {
        "compute": {
            "vms_total": len(vm),
            "classes": Counter(row["infrastructure_class"] for row in vm).most_common(),
            "platform_excluded": sum(1 for row in vm if row["backup_scope"] == "EXCLUDED_PLATFORM"),
            "coverage": coverage,
            "prod_traditional": coverage["PROD"]["with"] + coverage["PROD"]["without"],
            "findings": Counter(row["finding"] for row in vm).most_common(),
            "active_backups": len(active),
            "retention": [[bucket, Counter(retention_bucket(row["retention_max_days"]) for row in active)[bucket]]
                          for bucket in ("≤ 35 días", "~1 año", "4 a 12 años")],
            "retention_360": sum(1 for row in active if as_int(row["retention_max_days"]) == 360),
            "policy_v1": sum(1 for row in active if row["policy_type"] == "V1"),
            "archive": sum(1 for row in active if row["archive_enabled"] == "YES"),
            "k8s_total": len(k8s),
            "aks": sum(1 for row in k8s if row["service_type"] == "AKS"),
            "aro": sum(1 for row in k8s if row["service_type"] == "ARO"),
            "k8s_backup": sum(1 for row in k8s if row["backup_enabled"] == "YES"),
            "flux": sum(1 for row in k8s if row["gitops_flux"] == "YES"),
            "aks_prod": sum(1 for row in k8s if row["service_type"] == "AKS" and row["environment"] == "PROD"),
        },
        "orphans": {
            "total": len(orphans),
            "deleted": sum(1 for row in orphans if row["orphan_reason"] == "SOURCE_RESOURCE_DELETED"),
            "stale": sum(1 for row in orphans if row["orphan_reason"] == "ACTIVE_WITHOUT_BACKUP_30D"),
            "vaults": len({row["vault_name"] for row in orphans}),
            "recovery_points": sum(as_int(row["recovery_point_count"]) for row in orphans),
            "age_years": sorted(Counter(min(as_int(as_int(row["days_since_last_recovery_point"]) / 365), 5)
                                        for row in orphans).items()),
            "measured_floor": round(orphan_vault_storage + pi_orphans),
        },
        "vaults": {
            "total": len(vaults),
            "with_items": sum(1 for row in vaults if as_int(row["item_count"]) > 0),
            "grs": sum(1 for row in vaults if row["storage_redundancy"] == "GeoRedundant"),
            "immutable": sum(1 for row in vaults if row["immutability"] in ("Enabled", "Locked")),
            "crr": sum(1 for row in vaults if row["cross_region_restore"] == "Enabled"),
            "items": sum(as_int(row["item_count"]) for row in vaults),
            "policies": len(policies),
            "policies_unused": sum(1 for row in policies if str(row.get("protected_items_count")) in ("0", "")),
        },
        "database": {
            "total": len(db),
            "services": Counter(row["service_type"] for row in db).most_common(),
            "methods": Counter(row["backup_method"] for row in db).most_common(),
            "native": sum(1 for row in db if row["backup_method"] == "NATIVE"),
            "vault": sum(1 for row in db if "VAULT" in row["backup_method"]),
            "sql_total": len(sql),
            "sql_prod": sum(1 for row in sql if row["environment"] == "PROD"),
            "sql_prod_ltr": sum(1 for row in sql if row["environment"] == "PROD" and row["ltr_configured"] == "YES"),
            "sql_nonprod_geo": sum(1 for row in db if row["finding"] == "NONPROD_GEO_BACKUP_REDUNDANCY"),
            "prod_not_geo": sum(1 for row in db if row["finding"] == "PROD_BACKUP_NOT_GEO_REDUNDANT"),
            "cosmos_periodic": sum(1 for row in db if row["service_type"] == "COSMOS_DB"
                                   and "Continuous" not in row["native_backup_interval"]),
            "sql_vm": sum(1 for row in db if row["service_type"] == "SQL_SERVER_VM"),
            "redundancy_by_env": {env: Counter(row["native_backup_redundancy"] or "N/A" for row in sql
                                               if row["environment"] == env).most_common()
                                  for env in ("PROD", "CERT", "DESA")},
        },
        "storage": {
            "available": bool(storage),
            "total": len(storage),
            "services": Counter(row["service_type"] for row in storage).most_common(),
            "methods": Counter(row["backup_method"] for row in storage).most_common(),
            "findings": Counter(row["finding"] for row in storage).most_common(),
            "prod_without_recovery": sum(1 for row in storage if row["finding"] == "PROD_WITHOUT_RECOVERY"),
            "prod_share_without_backup": sum(1 for row in storage if row["finding"] == "PROD_SHARE_WITHOUT_BACKUP"),
            "prod_lrs_no_backup": sum(1 for row in storage if row["finding"] == "PROD_LRS_NO_BACKUP"),
            "unattached": sum(1 for row in storage if row["finding"] == "UNATTACHED_DISK"),
            "snapshot_sprawl": sum(1 for row in storage if row["finding"] == "SNAPSHOT_SPRAWL"),
            "snapshot_gb": round(sum(float(row["snapshot_total_gb"] or 0) for row in storage)),
            "vault_protected": sum(1 for row in storage if row["backup_method"] in ("VAULT_BV", "VAULT_RSV")),
            "vm_backup_disks": sum(1 for row in storage if row["backup_method"] == "VM_BACKUP"),
            "by_env_method": {env: Counter(row["backup_method"] for row in storage if row["environment"] == env).most_common()
                              for env in ("PROD", "CERT", "DESA")},
        },
        "native": {
            "total": len(native),
            "services": Counter(row["service_type"] for row in native).most_common(),
            "iac": sum(1 for row in native if row["iac_managed"] == "YES"),
            "kv_total": len(key_vaults),
            "kv_purge": sum(1 for row in key_vaults if row["purge_protection"] == "YES"),
            "redis_total": len(redis),
            "redis_premium": sum(1 for row in redis if "premium" in (row["sku"] or "").lower()),
            "redis_premium_nonprod": sum(1 for row in redis if row["finding"] == "NONPROD_PREMIUM_SKU"),
            "dns_no_lock": sum(1 for row in native if row["finding"] == "PROD_DNS_ZONE_WITHOUT_LOCK"),
            "netcore_no_lock": sum(1 for row in native if row["finding"] == "PROD_NETWORK_CORE_WITHOUT_LOCK"),
            "locked": sum(1 for row in native if row["has_delete_lock"] == "YES"),
            "plans_total": len(plans),
            "plans_empty": sum(1 for row in plans if row["finding"] == "EMPTY_APP_SERVICE_PLAN"),
            "plans_zr_nonprod": sum(1 for row in plans if row["finding"] == "NONPROD_ZONE_REDUNDANT_PLAN"),
        },
        "cost": {
            "calibrated_total": round((calibration[4] or 0) + (calibration[5] or 0) + (calibration[6] or 0)),
            "calibrated_storage": round(calibration[4] or 0),
            "calibrated_pi": round(calibration[5] or 0),
            "calibrated_vaults": len(vault_cost_rows),
            "calibrated_items": calibration[2],
            "calibrated_orphans": calibration[3],
            "orphan_vaults_measured": round(calibration[9] or 0),
            "pi_rate": assumptions.get("pi_usd"),
            "gb_rate": assumptions.get("grs_gb_usd"),
            "live_item_storage": round(assumptions.get("live_item_storage_usd") or 0, 1),
        },
        "snapshots": {
            "count": sum(as_int(row["snapshot_count"]) for row in snapshots),
            "full": sum(as_int(row["snapshot_full_count"]) for row in snapshots),
            "incremental": sum(as_int(row["snapshot_incremental_count"]) for row in snapshots),
            "gb": round(sum(float(row["snapshot_total_gb"] or 0) for row in snapshots)),
            "orphan_count": sum(as_int(row["snapshot_count"]) for row in snapshots if row["source_exists"] != "YES"),
            "orphan_gb": round(sum(float(row["snapshot_total_gb"] or 0) for row in snapshots if row["source_exists"] != "YES")),
            "old_sources": sum(1 for row in snapshots if row["source_exists"] == "YES"
                               and float(row["oldest_snapshot_age_days"] or 0) > 365),
            "spend": round(assumptions.get("snapshot_spend_usd") or 0),
        },
        "estate": {label: value for front, label, value, *_ in workbook["Resumen"].iter_rows(min_row=5, values_only=True)
                   if front == "Costo"},
        "storage_cost": top_storage_accounts(args.account_costs, storage,
                                             storage_path=f"{base}/backup-storage/storage_inventory.csv"),
        "third_party": third_party(workbook),
        "opportunities": opportunities,
        "risks": risks,
        "long_term": long_term,
    }
    backup_ids = {"O1", "O2", "O3", "O4", "O10", "O12"}
    metrics["savings"] = {
        "backup_monthly": sum(item["monthly"] for item in opportunities if item["id"] in backup_ids),
        "adjacent_monthly": sum(item["monthly"] for item in opportunities if item["id"] not in backup_ids),
        "total_monthly": sum(item["monthly"] for item in opportunities),
    }
    with open(args.output, "w", encoding="utf-8") as handle:
        json.dump(metrics, handle, indent=2, ensure_ascii=False, default=str)
    print(args.output)


if __name__ == "__main__":
    main()
