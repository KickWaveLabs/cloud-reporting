#!/usr/bin/env python3
"""AKS and Azure Container Apps inventory, plus the API Management APIs that front them.

Resource Graph for clusters, node pools, Container Apps environments and apps, and APIM services;
read-only ARM GET for the APIs of each APIM service (name, path and backend host). Tags are only
used to derive the environment and are not written. Identity fields (aadProfile, identity,
servicePrincipalProfile, linuxProfile, windowsProfile) are never projected.

Outputs (in --output-dir):
  aks_clusters.csv       one row per cluster: version, tier, size, add-ons and the configuration
                         properties compared in the workbook's Config_AKS sheet
  aks_node_pools.csv     one row per node pool: mode, VM size, count, autoscaling, zones, OS disk
  aca_environments.csv   Container Apps environments: workload profiles, VNet, zone redundancy
  aca_apps.csv           one row per app: environment, profile, min/max replicas, CPU/memory, scale rules
  apim_services.csv      API Management services: SKU, capacity, region, custom domains (to tell BS from Core)
  apim_apis.csv          APIs of the selected APIM services: id, path, backend host
  inventory_errors.csv
  inventory_summary.json counts by environment, and every query run
"""

from collections import Counter
from urllib.parse import urlparse

import aa_common as aa

bu = aa.bu
APIM_API = "2022-08-01"

QUERIES = {
    "aks": (
        "resources"
        " | where type =~ 'microsoft.containerservice/managedclusters'"
        " | project id, name, resourceGroup, subscriptionId, location, tags,"
        " skuName = tostring(sku.name), skuTier = tostring(sku.tier),"
        " version = tostring(coalesce(properties.currentKubernetesVersion, properties.kubernetesVersion)),"
        " powerState = tostring(properties.powerState.code),"
        " nodeResourceGroup = tostring(properties.nodeResourceGroup),"
        " privateCluster = tostring(properties.apiServerAccessProfile.enablePrivateCluster),"
        " authorizedIpRanges = array_length(properties.apiServerAccessProfile.authorizedIPRanges),"
        " disableLocalAccounts = tostring(properties.disableLocalAccounts),"
        " enableRbac = tostring(properties.enableRBAC),"
        " entraManaged = tostring(properties.aadProfile.managed),"
        " azureRbac = tostring(properties.aadProfile.enableAzureRBAC),"
        " networkPlugin = tostring(properties.networkProfile.networkPlugin),"
        " networkPluginMode = tostring(properties.networkProfile.networkPluginMode),"
        " networkPolicy = tostring(properties.networkProfile.networkPolicy),"
        " networkDataplane = tostring(properties.networkProfile.networkDataplane),"
        " outboundType = tostring(properties.networkProfile.outboundType),"
        " upgradeChannel = tostring(properties.autoUpgradeProfile.upgradeChannel),"
        " nodeOsUpgradeChannel = tostring(properties.autoUpgradeProfile.nodeOSUpgradeChannel),"
        " defender = tostring(properties.securityProfile.defender.securityMonitoring.enabled),"
        " workloadIdentity = tostring(properties.securityProfile.workloadIdentity.enabled),"
        " imageCleaner = tostring(properties.securityProfile.imageCleaner.enabled),"
        " oidcIssuer = tostring(properties.oidcIssuerProfile.enabled),"
        " diskEncryptionSet = isnotempty(properties.diskEncryptionSetID),"
        " azurePolicy = tostring(properties.addonProfiles.azurepolicy.enabled),"
        " keyVaultCsi = tostring(properties.addonProfiles.azureKeyvaultSecretsProvider.enabled),"
        " containerInsights = tostring(properties.addonProfiles.omsagent.enabled),"
        " logAnalyticsWorkspace = tostring(properties.addonProfiles.omsagent.config.logAnalyticsWorkspaceResourceID),"
        " managedPrometheus = tostring(properties.azureMonitorProfile.metrics.enabled),"
        " serviceMesh = tostring(properties.serviceMeshProfile.mode),"
        " istioRevisions = strcat_array(properties.serviceMeshProfile.istio.revisions, ';'),"
        " istioIngress = array_length(properties.serviceMeshProfile.istio.components.ingressGateways),"
        " keda = tostring(properties.workloadAutoScalerProfile.keda.enabled),"
        " vpa = tostring(properties.workloadAutoScalerProfile.verticalPodAutoscaler.enabled),"
        " webAppRouting = tostring(properties.ingressProfile.webAppRouting.enabled),"
        " costAnalysis = tostring(properties.metricsProfile.costAnalysis.enabled),"
        " supportPlan = tostring(properties.supportPlan),"
        " pools = properties.agentPoolProfiles"
    ),
    # Only the extension type: extension settings can hold sensitive values.
    "extensions": (
        "kubernetesconfigurationresources"
        " | where type in~ ('microsoft.kubernetesconfiguration/extensions',"
        " 'microsoft.kubernetesconfiguration/fluxconfigurations')"
        " | project id, type,"
        " extensionType = tolower(tostring(coalesce(properties.ExtensionType, properties.extensionType)))"
    ),
    "aca_environments": (
        "resources"
        " | where type =~ 'microsoft.app/managedenvironments'"
        " | project id, name, resourceGroup, subscriptionId, location, tags,"
        " workloadProfiles = properties.workloadProfiles,"
        " internal = tostring(properties.vnetConfiguration.internal),"
        " vnet = isnotempty(properties.vnetConfiguration.infrastructureSubnetId),"
        " zoneRedundant = tostring(properties.zoneRedundant),"
        " infrastructureResourceGroup = tostring(properties.infrastructureResourceGroup),"
        " peerEncryption = tostring(properties.peerTrafficConfiguration.encryption.enabled),"
        " publicNetworkAccess = tostring(properties.publicNetworkAccess)"
    ),
    # Only sizes, scale settings and rule types: env values and secrets are not projected.
    "aca_apps": (
        "resources"
        " | where type in~ ('microsoft.app/containerapps', 'microsoft.app/jobs')"
        " | project id, name, type, resourceGroup, subscriptionId, location, tags,"
        " environmentId = tolower(tostring(coalesce(properties.environmentId, properties.managedEnvironmentId))),"
        " workloadProfile = tostring(properties.workloadProfileName),"
        " minReplicas = properties.template.scale.minReplicas,"
        " maxReplicas = properties.template.scale.maxReplicas,"
        " cooldownPeriod = properties.template.scale.cooldownPeriod,"
        " ruleTypes = properties.template.scale.rules,"
        " containers = properties.template.containers,"
        " ingressExternal = tostring(properties.configuration.ingress.external),"
        " revisionsMode = tostring(properties.configuration.activeRevisionsMode),"
        " dapr = tostring(properties.configuration.dapr.enabled),"
        " runningStatus = tostring(properties.runningStatus)"
    ),
    "apim": (
        "resources"
        " | where type =~ 'microsoft.apimanagement/service'"
        " | project id, name, resourceGroup, subscriptionId, location, tags,"
        " sku = tostring(sku.name), capacity = toint(sku.capacity),"
        " vnetType = tostring(properties.virtualNetworkType),"
        # Only hostName is read from each entry; certificate fields are ignored.
        " hostnameConfigurations = properties.hostnameConfigurations,"
        " additionalLocations = array_length(properties.additionalLocations)"
    ),
}


def text(value):
    return "" if value is None else str(value)


def base_row(ctx, row):
    environment, evidence = aa.cluster_environment(row, ctx.subscription_name(row["subscriptionId"]))
    return {
        "subscription": ctx.subscription_name(row["subscriptionId"]),
        "subscription_id": row["subscriptionId"],
        "resource_group": row["resourceGroup"],
        "name": row["name"],
        "location": row["location"],
        "environment": environment,
        "environment_evidence": evidence,
        "cod_app": bu.normalize_cod_app(row.get("tags")),
    }


def pool_rows(cluster, pools):
    rows = []
    for pool in pools or []:
        rows.append({
            "cluster_id": cluster["id"].lower(),
            "cluster": cluster["name"],
            "environment": cluster["environment"],
            "pool": pool.get("name", ""),
            "mode": pool.get("mode", ""),
            "vm_size": pool.get("vmSize", ""),
            "count": pool.get("count") or 0,
            "autoscaling": text(pool.get("enableAutoScaling")),
            "min_count": text(pool.get("minCount")),
            "max_count": text(pool.get("maxCount")),
            "zones": ";".join(pool.get("availabilityZones") or []),
            "os_type": pool.get("osType", ""),
            "os_sku": pool.get("osSKU", ""),
            "os_disk_type": pool.get("osDiskType", ""),
            "os_disk_gb": text(pool.get("osDiskSizeGB")),
            "max_pods": text(pool.get("maxPods")),
            "encryption_at_host": text(pool.get("enableEncryptionAtHost")),
            "fips": text(pool.get("enableFIPS")),
            "orchestrator_version": pool.get("currentOrchestratorVersion") or pool.get("orchestratorVersion", ""),
            "power_state": bu.safe_get(pool, "powerState", "code", default=""),
            "spot": "YES" if pool.get("scaleSetPriority") == "Spot" else "NO",
        })
    return rows


def build_clusters(ctx, rows, extensions):
    clusters, pools = [], []
    ext_by_cluster = {}
    for ext in extensions:
        parent = ext["id"].lower().split("/providers/microsoft.kubernetesconfiguration/")[0]
        ext_by_cluster.setdefault(parent, []).append("flux" if "fluxconfigurations" in ext["type"].lower()
                                                     else ext.get("extensionType", ""))
    for row in rows:
        cluster = base_row(ctx, row)
        cluster_pools = pool_rows(dict(cluster, id=row["id"]), row.get("pools"))
        system = [pool for pool in cluster_pools if pool["mode"] == "System"]
        user = [pool for pool in cluster_pools if pool["mode"] != "System"]
        extension_types = sorted(set(ext_by_cluster.get(row["id"].lower(), [])))
        cluster.update({
            "kubernetes_version": row.get("version", ""),
            "sku_tier": row.get("skuTier", ""),
            "power_state": row.get("powerState", ""),
            "node_resource_group": row.get("nodeResourceGroup", ""),
            "node_pools": len(cluster_pools),
            "system_nodes": sum(int(pool["count"]) for pool in system),
            "user_nodes": sum(int(pool["count"]) for pool in user),
            "vm_sizes": ";".join(sorted({pool["vm_size"] for pool in cluster_pools})),
            "user_max_nodes": sum(int(pool["max_count"] or pool["count"]) for pool in user),
            "private_cluster": row.get("privateCluster", ""),
            "authorized_ip_ranges": "YES" if (row.get("authorizedIpRanges") or 0) > 0 else "NO",
            "disable_local_accounts": row.get("disableLocalAccounts", ""),
            "enable_rbac": row.get("enableRbac", ""),
            "entra_integration": row.get("entraManaged", ""),
            "azure_rbac": row.get("azureRbac", ""),
            "network_plugin": row.get("networkPlugin", ""),
            "network_plugin_mode": row.get("networkPluginMode", ""),
            "network_policy": row.get("networkPolicy", ""),
            "network_dataplane": row.get("networkDataplane", ""),
            "outbound_type": row.get("outboundType", ""),
            "upgrade_channel": row.get("upgradeChannel", ""),
            "node_os_upgrade_channel": row.get("nodeOsUpgradeChannel", ""),
            "defender": row.get("defender", ""),
            "workload_identity": row.get("workloadIdentity", ""),
            "oidc_issuer": row.get("oidcIssuer", ""),
            "image_cleaner": row.get("imageCleaner", ""),
            "disk_encryption_set": "YES" if row.get("diskEncryptionSet") else "NO",
            "encryption_at_host": "YES" if cluster_pools and all(p["encryption_at_host"] == "True" for p in cluster_pools)
                                  else "NO",
            "azure_policy": row.get("azurePolicy", ""),
            "key_vault_csi": row.get("keyVaultCsi", ""),
            "container_insights": row.get("containerInsights", ""),
            "log_analytics_workspace": row.get("logAnalyticsWorkspace", ""),
            "managed_prometheus": row.get("managedPrometheus", ""),
            "service_mesh": row.get("serviceMesh", "") or "None",
            "istio_revisions": row.get("istioRevisions", ""),
            "istio_ingress_gateways": row.get("istioIngress") or 0,
            "keda": row.get("keda", ""),
            "vpa": row.get("vpa", ""),
            "web_app_routing": row.get("webAppRouting", ""),
            "cost_analysis": row.get("costAnalysis", ""),
            "support_plan": row.get("supportPlan", ""),
            "zones": ";".join(sorted({zone for pool in cluster_pools for zone in pool["zones"].split(";") if zone})),
            "gitops_flux": "YES" if "flux" in extension_types or "microsoft.flux" in extension_types else "NO",
            "extension_types": ";".join(extension_types),
            "id": row["id"],
        })
        clusters.append(cluster)
        pools.extend(cluster_pools)
    return clusters, pools


def build_aca(ctx, environments, apps):
    env_rows = []
    for row in environments:
        environment = base_row(ctx, row)
        profiles = row.get("workloadProfiles") or []
        environment.update({
            "plan": "WORKLOAD_PROFILES" if profiles else "CONSUMPTION_ONLY",
            "profiles": ";".join(f"{p.get('name')}:{p.get('workloadProfileType')}:{p.get('minimumCount', '')}-"
                                 f"{p.get('maximumCount', '')}" for p in profiles),
            "dedicated_profiles": sum(1 for p in profiles if p.get("workloadProfileType") != "Consumption"),
            "vnet_integrated": "YES" if row.get("vnet") else "NO",
            "internal": row.get("internal", ""),
            "zone_redundant": row.get("zoneRedundant", ""),
            "peer_encryption": row.get("peerEncryption", ""),
            "public_network_access": row.get("publicNetworkAccess", ""),
            "infrastructure_resource_group": row.get("infrastructureResourceGroup", ""),
            "id": row["id"],
        })
        env_rows.append(environment)
    env_names = {row["id"].lower(): row["name"] for row in env_rows}
    app_rows = []
    for row in apps:
        app = base_row(ctx, row)
        containers = row.get("containers") or []
        app.update({
            "kind": "JOB" if row["type"].lower().endswith("/jobs") else "APP",
            "managed_environment": env_names.get(row.get("environmentId", ""), row.get("environmentId", "").split("/")[-1]),
            "workload_profile": row.get("workloadProfile", "") or "Consumption",
            "min_replicas": text(row.get("minReplicas")),
            "max_replicas": text(row.get("maxReplicas")),
            "cooldown_seconds": text(row.get("cooldownPeriod")),
            "scale_rule_types": ";".join(sorted({
                rule.get("custom", {}).get("type", "") if "custom" in rule else next(iter(set(rule) - {"name"}), "")
                for rule in row.get("ruleTypes") or []})),
            "containers": len(containers),
            "cpu": round(sum(float(bu.safe_get(c, "resources", "cpu", default=0) or 0) for c in containers), 3),
            "memory_gib": round(sum(float(str(bu.safe_get(c, "resources", "memory", default="0")).rstrip("Gi") or 0)
                                    for c in containers), 3),
            "ingress_external": row.get("ingressExternal", ""),
            "revisions_mode": row.get("revisionsMode", ""),
            "dapr": row.get("dapr", ""),
            "running_status": row.get("runningStatus", ""),
            "id": row["id"],
        })
        app_rows.append(app)
    return env_rows, app_rows


def backend_host(service_url):
    if not service_url:
        return ""
    parsed = urlparse(service_url)
    return (parsed.hostname or "").lower()


def build_apim(ctx, services, args):
    service_rows = []
    for row in services:
        service = base_row(ctx, row)
        hostnames = sorted({str(item.get("hostName", "")).lower() for item in row.get("hostnameConfigurations") or []
                            if item.get("hostName") and not str(item.get("hostName", "")).endswith(".azure-api.net")})
        service.update({"sku": row.get("sku", ""), "capacity": row.get("capacity", ""),
                        "custom_domains": ";".join(hostnames),
                        "vnet_type": row.get("vnetType", ""), "additional_locations": row.get("additionalLocations") or 0,
                        "id": row["id"]})
        service_rows.append(service)
    names = {item.strip().lower() for item in args.apim_names.split(",") if item.strip()}
    environments = aa.wanted_environments(args)
    selected = [row for row in service_rows
                if (names and any(name in row["name"].lower() for name in names))
                or (not names and (environments is None or row["environment"] in environments))]
    api_rows = []
    if args.skip_arm:
        return service_rows, api_rows

    def fetch(service):
        # Only API definitions: subscription keys and named values are never read.
        apis = ctx.try_arm("apim-apis", service["name"], ctx.arm_list, f"{service['id']}/apis", APIM_API)
        rows = []
        for api in apis or []:
            properties = api.get("properties", {})
            rows.append({
                "apim": service["name"], "apim_environment": service["environment"], "api_id": api.get("name", ""),
                "display_name": properties.get("displayName", ""), "path": properties.get("path", ""),
                "backend_host": backend_host(properties.get("serviceUrl")),
                "api_type": properties.get("type", "") or properties.get("apiType", ""),
                "is_current": text(properties.get("isCurrent")), "apim_id": service["id"],
            })
        return rows

    for result in ctx.parallel(fetch, selected, label="apim-apis", workers=2):
        api_rows.extend(result or [])
    return service_rows, api_rows


def main():
    parser = bu.build_parser(__doc__.splitlines()[0], aa.results_path("aa-inventory"))
    aa.add_common_arguments(parser)
    parser.add_argument("--apim-names", default="",
                        help="Comma-separated name fragments of the APIM services whose APIs are listed "
                             "(for example 'bs,core'). Default: every APIM in --environments")
    parser.add_argument("--include-poc", action="store_true",
                        help="Keep resources in POC subscriptions (left out by default)")
    args = parser.parse_args()
    bu.setup_logging(args.verbose)

    ctx = bu.AzureContext(args)
    ctx.load_subscriptions()
    results = ctx.queries(QUERIES)
    poc = set()
    if not args.include_poc:
        for key, rows in results.items():
            kept = []
            for row in rows:
                name = ctx.subscription_name(row.get("subscriptionId") or row["id"].split("/")[2])
                if aa.is_poc_subscription(name):
                    poc.add(name)
                else:
                    kept.append(row)
            results[key] = kept
    clusters, pools = build_clusters(ctx, results["aks"], results["extensions"])
    environments, apps = build_aca(ctx, results["aca_environments"], results["aca_apps"])
    services, apis = build_apim(ctx, results["apim"], args)

    output = bu.Output(args.output_dir, args.delimiter)
    output.csv("aks_clusters.csv", clusters)
    output.csv("aks_node_pools.csv", pools)
    output.csv("aca_environments.csv", environments)
    output.csv("aca_apps.csv", apps)
    output.csv("apim_services.csv", services)
    output.csv("apim_apis.csv", apis)
    output.csv("inventory_errors.csv", ctx.errors.rows, bu.ErrorLog.FIELDS)
    output.json("inventory_summary.json", bu.run_metadata(ctx, {
        "aks_by_environment": dict(Counter(row["environment"] for row in clusters)),
        "aks_nodes_by_environment": bu.count_by(clusters, "environment", "vm_sizes"),
        "aks_service_mesh": dict(Counter(row["service_mesh"] for row in clusters)),
        "aks_monitoring": bu.count_by(clusters, "container_insights", "managed_prometheus"),
        "aca_apps_by_environment": dict(Counter(row["environment"] for row in apps)),
        "apim_services": [{"name": row["name"], "environment": row["environment"], "sku": row["sku"],
                           "custom_domains": row["custom_domains"]} for row in services],
        "apim_apis": dict(Counter(row["apim"] for row in apis)),
        "excluded_poc_subscriptions": sorted(poc),
    }))
    if args.xlsx:
        output.xlsx("inventory.xlsx")


if __name__ == "__main__":
    main()
