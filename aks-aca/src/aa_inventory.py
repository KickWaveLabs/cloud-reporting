#!/usr/bin/env python3
"""AKS, Azure Container Apps and App Service inventory, plus the API Management APIs that front them.

Resource Graph for clusters, node pools, Container Apps environments and apps, App Service plans and
sites, and APIM services; read-only ARM GET for the APIs of each APIM service (name, path and backend
host). Tags are only used to derive the environment and the app code and are not written. Identity
fields (aadProfile, identity, servicePrincipalProfile, linuxProfile, windowsProfile) and app settings
are never projected.

App code (cod_app): the codApp tag first, then the naming convention (aksveu2coaap01 -> COAA);
cod_app_evidence says which (TAG, NAME, or ENVIRONMENT for a Container App without either).

APIs: apim_layer is the APIM's place in the north-south chain (BS, CORE, UX, OPEN_BANKING, YAPE, OTHER; from its
custom domain). backend_platform says where the backend runs (AKS, APP_SERVICE, CONTAINER_APPS, OPENSHIFT, APIM,
OTHER, NONE). backend_cod_app is the owning application and cod_app_source how it was found (HOST, HOST_NORMALIZED,
K8S_NAMESPACE, BACKEND_RESOURCE, NAMING; see aa_common.resolve_api_cod_app). Only AKS backends count as AKS traffic.

Runtimes (AKS clusters, Container Apps environments and apps, App Service plans and sites, including function apps)
carry created_at (Lima date, from ARM createdTime: Resource Graph has no creation date for AKS or App Service),
measurement_month (the month cost and metrics cover; --measurement-month, by default the last full month) and
existed_at_measurement: YES (created before that month), PARTIAL (created during it), NO (created after it) or
UNKNOWN. The inventory is current while cost and metrics are a past month: filter on this column.

Outputs (in --output-dir):
  aks_clusters.csv       one row per cluster: version, tier, size, add-ons and the configuration
                         properties compared in the workbook's Config_AKS sheet
  aks_node_pools.csv     one row per node pool: mode, VM size, count, autoscaling, zones, OS disk
  aca_environments.csv   Container Apps environments: workload profiles, VNet, zone redundancy
  aca_apps.csv           one row per app: environment, profile, min/max replicas, CPU/memory, scale rules
  apim_services.csv      API Management services: SKU, capacity, region, custom domains and layer
  apim_apis.csv          APIs of the selected APIM services: id, path, layer, backend host, platform and app code
  app_service_plans.csv  App Service plans: SKU, tier, workers, OS, number of sites, runtime (APP_SERVICE
                         or FUNCTIONS: the deck card and cost owner type the plan counts under)
  app_services.csv       web and function apps: type, container or code, state, plan, its SKU and id
  inventory_errors.csv
  inventory_summary.json counts by environment, and every query run
"""

from collections import Counter
from datetime import datetime, timedelta
from urllib.parse import urlparse

import aa_common as aa

azure = aa.azure
APIM_API = "2022-08-01"
RESOURCES_API = "2021-04-01"
# Runtime types whose creation date is read; one ARM list per subscription covers all of them.
RUNTIME_TYPES = ("Microsoft.ContainerService/managedClusters", "Microsoft.App/managedEnvironments",
                 "Microsoft.App/containerApps", "Microsoft.App/jobs", "Microsoft.Web/serverFarms", "Microsoft.Web/sites")
# App Service plan tiers that only host function apps: Consumption, Premium (Elastic) and Flex Consumption.
FUNCTION_TIERS = ("Dynamic", "ElasticPremium", "FlexConsumption")

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
    # App Service plans and sites (web apps and function apps): size and state only, no app settings.
    "app_service_plans": (
        "resources"
        " | where type =~ 'microsoft.web/serverfarms'"
        " | project id, name, resourceGroup, subscriptionId, location, tags, kind,"
        " sku = tostring(sku.name), tier = tostring(sku.tier), workers = toint(sku.capacity),"
        " linux = tostring(properties.reserved), sites = toint(properties.numberOfSites),"
        " zoneRedundant = tostring(properties.zoneRedundant), status = tostring(properties.status)"
    ),
    "app_services": (
        "resources"
        " | where type =~ 'microsoft.web/sites'"
        " | project id, name, resourceGroup, subscriptionId, location, tags, kind,"
        " state = tostring(properties.state), planId = tolower(tostring(properties.serverFarmId))"
    ),
}


def text(value):
    return "" if value is None else str(value)


def base_row(ctx, row):
    environment, evidence = aa.cluster_environment(row, ctx.subscription_name(row["subscriptionId"]))
    cod_app, cod_app_evidence = azure.resolve_cod_app(row.get("tags"), row["name"])
    return {
        "subscription": ctx.subscription_name(row["subscriptionId"]),
        "subscription_id": row["subscriptionId"],
        "resource_group": row["resourceGroup"],
        "name": row["name"],
        "location": row["location"],
        "environment": environment,
        "environment_evidence": evidence,
        "cod_app": cod_app,
        "cod_app_evidence": cod_app_evidence,
    }


def pool_rows(cluster, pools):
    rows = []
    for pool in pools or []:
        rows.append({
            "cluster_id": cluster["id"].lower(),
            "cluster": cluster["name"],
            "cod_app": cluster["cod_app"],
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
            "power_state": azure.safe_get(pool, "powerState", "code", default=""),
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
    env_codes = {row["id"].lower(): row["cod_app"] for row in env_rows}
    app_rows = []
    for row in apps:
        app = base_row(ctx, row)
        if not app["cod_app"] and env_codes.get(row.get("environmentId", "")):
            app.update(cod_app=env_codes[row["environmentId"]], cod_app_evidence="ENVIRONMENT")
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
            "cpu": round(sum(float(azure.safe_get(c, "resources", "cpu", default=0) or 0) for c in containers), 3),
            "memory_gib": round(sum(float(str(azure.safe_get(c, "resources", "memory", default="0")).rstrip("Gi") or 0)
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



def build_app_service(ctx, plans, sites):
    plan_rows = []
    for row in plans:
        plan = base_row(ctx, row)
        plan.update({"sku": row.get("sku", ""), "tier": row.get("tier", ""), "workers": row.get("workers") or 0,
                     "os": "LINUX" if str(row.get("linux", "")).lower() == "true" else "WINDOWS",
                     "kind": row.get("kind", ""), "sites": row.get("sites") or 0,
                     "zone_redundant": row.get("zoneRedundant", ""), "status": row.get("status", ""), "id": row["id"]})
        plan_rows.append(plan)
    by_id = {row["id"].lower(): row for row in plan_rows}
    site_rows = []
    for row in sites:
        site = base_row(ctx, row)
        plan = by_id.get(row.get("planId", ""), {})
        kind = str(row.get("kind", "")).lower()
        site.update({"type": "FUNCTION" if "functionapp" in kind else "WEB",
                     "container": "YES" if "container" in kind else "NO", "kind": row.get("kind", ""),
                     "state": row.get("state", ""), "plan": plan.get("name", row.get("planId", "").split("/")[-1]),
                     "plan_sku": plan.get("sku", ""), "plan_id": row.get("planId", ""), "id": row["id"]})
        site_rows.append(site)
    site_types = {}
    for site in site_rows:
        site_types.setdefault(site["plan_id"], set()).add(site["type"])
    for plan in plan_rows:
        plan["runtime"] = plan_runtime(plan, site_types.get(plan["id"].lower(), set()))
    return plan_rows, site_rows


def plan_runtime(plan, site_types):
    """FUNCTIONS for a Functions-only SKU or a plan whose sites are all function apps; else APP_SERVICE.

    A plan shared by web and function apps counts as APP_SERVICE, so its cost is never split.
    """
    if plan.get("tier", "") in FUNCTION_TIERS or site_types == {"FUNCTION"}:
        return "FUNCTIONS"
    return "APP_SERVICE"


def created_times(ctx, subscription_ids):
    """Resource id (lower case) -> createdTime (ISO, UTC) of every runtime, one ARM list per subscription."""
    type_filter = " or ".join(f"resourceType eq '{kind}'" for kind in RUNTIME_TYPES)

    def fetch(subscription_id):
        items = ctx.try_arm("created-time", subscription_id, ctx.arm_list, f"/subscriptions/{subscription_id}/resources",
                            RESOURCES_API, {"$filter": type_filter, "$expand": "createdTime"})
        return {item["id"].lower(): item.get("createdTime", "") for item in items or []}

    out = {}
    for result in ctx.parallel(fetch, sorted(subscription_ids), label="created-time"):
        out.update(result or {})
    return out


def measurement_window(month):
    """(first day, last day) of a YYYY-MM month, or of the last full month when empty."""
    if not month:
        return aa.last_full_month()
    first = datetime.strptime(month, "%Y-%m").date()
    following = (first.replace(day=28) + timedelta(days=4)).replace(day=1)
    return first, following - timedelta(days=1)


def with_created(rows, created, window):
    """Add created_at, measurement_month and existed_at_measurement right before the id column."""
    first, last = window
    month = first.strftime("%Y-%m")

    def created_day(row):
        value = created.get(row["id"].lower(), "")
        if not value:
            return None
        moment = datetime.fromisoformat(value[:19] + "+00:00")
        return moment.astimezone(aa.LIMA).date()

    def existed(day):
        if day is None:
            return "UNKNOWN"
        if day < first:
            return "YES"
        return "PARTIAL" if day <= last else "NO"

    out = []
    for row in rows:
        day = created_day(row)
        copy = {key: value for key, value in row.items() if key != "id"}
        copy.update(created_at=day.isoformat() if day else "", measurement_month=month, existed_at_measurement=existed(day),
                    id=row["id"])
        out.append(copy)
    return out


def api_code_lookups(clusters, environments, apps, plans, sites):
    """App codes of the inventory, and resource name -> code for App Service sites and Container Apps."""
    known = {row["cod_app"].upper() for row in clusters + environments + apps + plans + sites if row.get("cod_app")}
    resources = {row["name"].lower(): row["cod_app"] for row in sites + apps if row.get("cod_app")}
    return known, resources


def build_apim(ctx, services, args, known_codes, resource_codes):
    service_rows = []
    for row in services:
        service = base_row(ctx, row)
        hostnames = sorted({str(item.get("hostName", "")).lower() for item in row.get("hostnameConfigurations") or []
                            if item.get("hostName") and not str(item.get("hostName", "")).endswith(".azure-api.net")})
        service.update({"sku": row.get("sku", ""), "capacity": row.get("capacity", ""),
                        "custom_domains": ";".join(hostnames),
                        "apim_layer": aa.apim_layer(row["name"], ";".join(hostnames)),
                        "vnet_type": row.get("vnetType", ""), "additional_locations": row.get("additionalLocations") or 0,
                        "id": row["id"]})
        service_rows.append(service)
    names = {item.strip().lower() for item in args.apim_names.split(",") if item.strip()}
    environments = {item.strip().upper() for item in args.apim_environments.split(",") if item.strip()} or None
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
            host = backend_host(properties.get("serviceUrl"))
            code, source = aa.resolve_api_cod_app(host, api.get("name", ""), known_codes, resource_codes)
            rows.append({
                "apim": service["name"], "apim_environment": service["environment"], "apim_layer": service["apim_layer"],
                "api_id": api.get("name", ""),
                "display_name": properties.get("displayName", ""), "path": properties.get("path", ""),
                "backend_host": host, "backend_platform": aa.backend_platform(host),
                "backend_cod_app": code, "cod_app_source": source,
                "api_type": properties.get("type", "") or properties.get("apiType", ""),
                "is_current": text(properties.get("isCurrent")), "apim_id": service["id"],
            })
        return rows

    for result in ctx.parallel(fetch, selected, label="apim-apis", workers=2):
        api_rows.extend(result or [])
    return service_rows, api_rows


def main():
    parser = azure.build_parser(__doc__.splitlines()[0], aa.results_path("aa-inventory"))
    aa.add_common_arguments(parser)
    parser.add_argument("--apim-names", default="",
                        help="Comma-separated name fragments of the APIM services whose APIs are listed "
                             "(for example 'bs,core'). Default: every APIM in --apim-environments")
    parser.add_argument("--apim-environments", default="PROD",
                        help="Environments of the APIM services whose APIs are listed (DESA for a one-off check of "
                             "test activity). Default: %(default)s")
    parser.add_argument("--include-poc", action="store_true",
                        help="Keep resources in POC subscriptions (left out by default)")
    parser.add_argument("--measurement-month", default="",
                        help="Month cost and metrics cover (YYYY-MM), for existed_at_measurement. Default: last full month")
    args = parser.parse_args()
    azure.setup_logging(args.verbose)

    ctx = azure.AzureContext(args)
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
    plans, sites = build_app_service(ctx, results["app_service_plans"], results["app_services"])
    services, apis = build_apim(ctx, results["apim"], args, *api_code_lookups(clusters, environments, apps, plans, sites))
    window = measurement_window(args.measurement_month)
    runtimes = clusters + environments + apps + plans + sites
    created = {} if args.skip_arm else created_times(ctx, {row["subscription_id"] for row in runtimes})
    clusters, environments, apps, plans, sites = (with_created(rows, created, window)
                                                  for rows in (clusters, environments, apps, plans, sites))

    output = azure.Output(args.output_dir, args.delimiter)
    output.csv("aks_clusters.csv", clusters)
    output.csv("aks_node_pools.csv", pools)
    output.csv("aca_environments.csv", environments)
    output.csv("aca_apps.csv", apps)
    output.csv("apim_services.csv", services)
    output.csv("apim_apis.csv", apis)
    output.csv("app_service_plans.csv", plans)
    output.csv("app_services.csv", sites)
    output.csv("inventory_errors.csv", ctx.errors.rows, azure.ErrorLog.FIELDS)
    output.json("inventory_summary.json", azure.run_metadata(ctx, {
        "aks_by_environment": dict(Counter(row["environment"] for row in clusters)),
        "aks_nodes_by_environment": azure.count_by(clusters, "environment", "vm_sizes"),
        "aks_service_mesh": dict(Counter(row["service_mesh"] for row in clusters)),
        "aks_monitoring": azure.count_by(clusters, "container_insights", "managed_prometheus"),
        "aca_apps_by_environment": dict(Counter(row["environment"] for row in apps)),
        "apim_services": [{"name": row["name"], "environment": row["environment"], "sku": row["sku"],
                           "custom_domains": row["custom_domains"]} for row in services],
        "apim_apis": dict(Counter(row["apim"] for row in apis)),
        "apim_apis_with_cod_app": sum(1 for row in apis if row["backend_cod_app"]),
        "apim_apis_by_layer_and_platform": azure.count_by(apis, "apim_layer", "backend_platform"),
        "apim_apis_by_cod_app_source": dict(Counter(row["cod_app_source"] or "NONE" for row in apis)),
        "measurement_month": window[0].strftime("%Y-%m"),
        "existed_at_measurement": {name: dict(Counter(row["existed_at_measurement"] for row in rows))
                                   for name, rows in (("aks_clusters", clusters), ("aca_environments", environments),
                                                      ("aca_apps", apps), ("app_service_plans", plans),
                                                      ("app_services", sites))},
        "app_service_sites_by_environment": azure.count_by(sites, "environment", "type"),
        "cod_app_evidence": dict(Counter(row["cod_app_evidence"] for row in clusters + sites)),
        "excluded_poc_subscriptions": sorted(poc),
    }))
    if args.xlsx:
        output.xlsx("inventory.xlsx")


if __name__ == "__main__":
    main()
