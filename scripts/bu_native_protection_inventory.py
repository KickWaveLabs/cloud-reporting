#!/usr/bin/env python3
"""Native protection and rebuildability inventory for services without Azure Backup.

Covers App Service, Functions, Logic Apps Standard, Key Vault, Managed HSM, API Management,
Container Registry, core network, messaging (Service Bus, Event Hubs, Event Grid) and Redis.
For each resource it reports the native recovery features (automatic backups, soft delete,
geo-replication, zones, persistence), delete locks, IaC evidence and sizing fields for cost
analysis.

Outputs:
  native_protection_inventory.csv
  native_app_service_plans.csv
  native_network_counts.csv
  native_errors.csv
  native_summary.json
  native_protection_inventory.xlsx   (with --xlsx)

Read-only: Resource Graph plus ARM GET for locks, geo-DR aliases and Redis links.
No call that returns secrets is made (app settings, keys, backup configuration).

Dependencies:
  pip install azure-identity azure-mgmt-resourcegraph requests
"""

import argparse
from concurrent.futures import ThreadPoolExecutor

import bu_common as bu
from bu_common import LOG

DEFAULT_OUTPUT_DIR = "./results/backup-native"
LOCKS_API = "2020-05-01"
MESSAGING_API = "2021-11-01"
REDIS_API = "2024-03-01"
REDIS_ENTERPRISE_API = "2024-10-01"

COMMON_PROJECTION = "id, name, type, kind, resourceGroup, subscriptionId, location, tags, sku, zones"

SIZING_FIELDS = [
    "resource_kind", "sku", "tier", "capacity", "instance_count", "plan_id", "plan_name",
    "zone_redundant", "zone_count", "multi_region", "region_count", "state",
]
NATIVE_FIELDS = [
    "native_backup_capability", "native_backup_configured", "recovery_features", "soft_delete",
    "soft_delete_days", "purge_protection", "geo_dr", "has_delete_lock", "lock_scope",
    "iac_managed", "iac_evidence", "rebuild_dependency", "detail",
]
INVENTORY_FIELDS = (["service_type"] + bu.BASE_FIELDS + SIZING_FIELDS + NATIVE_FIELDS + ["backup_method"]
                    + bu.RESILIENCE_FIELDS + bu.FINDING_FIELDS)

RESILIENCE_TYPES = [
    "AzureApiManagement", "AzureApplicationGateway", "AzureContainerRegistry", "AzureEventHubNamespaces",
    "AzureKeyVault", "AzureLoadBalancer", "AzureNatGateways", "AzureRedisCache", "AzureServiceBus",
    "AzureVirtualNetworkGateways", "AzureWebServerFarms",
]

# App Service tiers with automatic hourly backups kept 30 days.
AUTOMATIC_BACKUP_TIERS = ("basic", "standard", "premium", "isolated")
NOT_SUPPORTED_TIERS = ("free", "shared", "dynamic", "elasticpremium", "flexconsumption")

IAC_TAG_KEYS = ("provisionedby", "provisioned_by", "managed", "managedby", "managed_by", "createdby", "iac")
IAC_TAG_VALUES = ("tfc", "terraform", "bicep", "arm", "pulumi", "ansible", "crossplane")

# High-volume network types: counted per subscription, not listed row by row.
COUNTED_NETWORK_TYPES = {
    "microsoft.network/publicipaddresses":
        "isempty(properties.ipConfiguration) and isempty(properties.natGateway)",
    "microsoft.network/networkinterfaces":
        "isempty(properties.virtualMachine) and isempty(properties.privateEndpoint)"
        " and isempty(properties.privateLinkService)",
    "microsoft.network/privateendpoints": "false",
    "microsoft.network/networksecuritygroups":
        "isempty(properties.subnets) and isempty(properties.networkInterfaces)",
    "microsoft.network/routetables": "isempty(properties.subnets)",
}

# Network resource types listed row by row: (resource_kind, is_core, detail expression).
NETWORK_SPECS = {
    "microsoft.network/virtualnetworks": (
        "VIRTUAL_NETWORK", True,
        "pack('address_prefixes', array_length(properties.addressSpace.addressPrefixes),"
        " 'subnets', array_length(properties.subnets),"
        " 'peerings', array_length(properties.virtualNetworkPeerings),"
        " 'ddos_protection', tostring(properties.enableDdosProtection))"),
    "microsoft.network/azurefirewalls": (
        "AZURE_FIREWALL", True,
        "pack('sku_tier', tostring(properties.sku.tier), 'sku_name', tostring(properties.sku.name),"
        " 'threat_intel', tostring(properties.threatIntelMode),"
        " 'policy', tostring(split(properties.firewallPolicy.id, '/')[-1]))"),
    "microsoft.network/firewallpolicies": (
        "FIREWALL_POLICY", True,
        "pack('sku_tier', tostring(properties.sku.tier),"
        " 'child_policies', array_length(properties.childPolicies),"
        " 'firewalls', array_length(properties.firewalls),"
        " 'base_policy', tostring(split(properties.basePolicy.id, '/')[-1]))"),
    "microsoft.network/applicationgateways": (
        "APPLICATION_GATEWAY", True,
        "pack('sku_name', tostring(properties.sku.name), 'sku_tier', tostring(properties.sku.tier),"
        " 'sku_capacity', tostring(properties.sku.capacity),"
        " 'autoscale_min', tostring(properties.autoscaleConfiguration.minCapacity),"
        " 'autoscale_max', tostring(properties.autoscaleConfiguration.maxCapacity),"
        " 'backend_pools', array_length(properties.backendAddressPools),"
        " 'operational_state', tostring(properties.operationalState))"),
    "microsoft.cdn/profiles": (
        "FRONT_DOOR_CDN_PROFILE", True,
        "pack('resource_state', tostring(properties.resourceState))"),
    "microsoft.network/frontdoors": (
        "FRONT_DOOR_CLASSIC", True,
        "pack('enabled_state', tostring(properties.enabledState),"
        " 'resource_state', tostring(properties.resourceState))"),
    "microsoft.network/dnszones": (
        "DNS_ZONE", True,
        "pack('record_sets', tostring(properties.numberOfRecordSets), 'zone_type', tostring(properties.zoneType))"),
    "microsoft.network/privatednszones": (
        "PRIVATE_DNS_ZONE", True,
        "pack('record_sets', tostring(properties.numberOfRecordSets),"
        " 'vnet_links', tostring(properties.numberOfVirtualNetworkLinks),"
        " 'vnet_links_registration', tostring(properties.numberOfVirtualNetworkLinksWithRegistration))"),
    "microsoft.network/virtualnetworkgateways": (
        "VIRTUAL_NETWORK_GATEWAY", True,
        "pack('sku_name', tostring(properties.sku.name), 'sku_tier', tostring(properties.sku.tier),"
        " 'sku_capacity', tostring(properties.sku.capacity), 'gateway_type', tostring(properties.gatewayType),"
        " 'vpn_type', tostring(properties.vpnType), 'active_active', tostring(properties.activeActive))"),
    "microsoft.network/expressroutecircuits": (
        "EXPRESSROUTE_CIRCUIT", True,
        "pack('bandwidth_mbps', tostring(properties.serviceProviderProperties.bandwidthInMbps),"
        " 'peering_location', tostring(properties.serviceProviderProperties.peeringLocation),"
        " 'provider_state', tostring(properties.serviceProviderProvisioningState),"
        " 'circuit_state', tostring(properties.circuitProvisioningState))"),
    "microsoft.network/loadbalancers": (
        "LOAD_BALANCER", False,
        "pack('frontends', array_length(properties.frontendIPConfigurations),"
        " 'frontend_zones', array_length(properties.frontendIPConfigurations[0].zones),"
        " 'backend_pools', array_length(properties.backendAddressPools),"
        " 'backend_members', countof(tostring(properties.backendAddressPools), '/ipConfigurations/')"
        " + countof(tostring(properties.backendAddressPools), '\"ipAddress\"'),"
        " 'rules', array_length(properties.loadBalancingRules))"),
    "microsoft.network/natgateways": (
        "NAT_GATEWAY", False,
        "pack('public_ips', array_length(properties.publicIpAddresses), 'subnets', array_length(properties.subnets))"),
    "microsoft.network/ddosprotectionplans": (
        "DDOS_PROTECTION_PLAN", True,
        "pack('virtual_networks', array_length(properties.virtualNetworks))"),
}


# ---------------------------------------------------------------------------
# Locks
# ---------------------------------------------------------------------------


class LockIndex:
    """Delete-blocking locks (CanNotDelete and ReadOnly) by scope, from one ARM GET per subscription."""

    MARKER = "/providers/microsoft.authorization/locks/"

    def __init__(self):
        self.scopes = {}
        self.loaded_subscriptions = set()
        self.enabled = False

    def load(self, ctx):
        self.enabled = True
        subscriptions = [value["id"] for value in ctx.subscriptions.values()]

        def fetch(subscription_id):
            path = f"/subscriptions/{subscription_id}/providers/Microsoft.Authorization/locks"
            return subscription_id, ctx.try_arm("locks", subscription_id, ctx.arm_list, path, LOCKS_API)

        for result in ctx.parallel(fetch, subscriptions, label="locks"):
            if not result or result[1] is None:
                continue
            subscription_id, locks = result
            self.loaded_subscriptions.add(bu.normalize(subscription_id))
            for lock in locks:
                lock_id = bu.normalize(lock.get("id"))
                if self.MARKER not in lock_id:
                    continue
                scope = lock_id.split(self.MARKER)[0]
                self.scopes[scope] = bu.safe_get(lock, "properties", "level", default="")
        LOG.info("Locks: %d scopes in %d subscriptions", len(self.scopes), len(self.loaded_subscriptions))

    def lookup(self, resource_id, subscription_id, resource_group):
        """Return (has_delete_lock, lock_scope)."""
        if not self.enabled or bu.normalize(subscription_id) not in self.loaded_subscriptions:
            return "UNKNOWN", ""
        resource_id = bu.normalize(resource_id)
        if resource_id in self.scopes:
            return "YES", f"RESOURCE:{self.scopes[resource_id]}"
        for parent in bu.parent_resource_ids(resource_id):
            if parent in self.scopes:
                return "YES", f"PARENT:{self.scopes[parent]}"
        rg_id = bu.resource_group_id(subscription_id, resource_group)
        if rg_id in self.scopes:
            return "YES", f"RESOURCE_GROUP:{self.scopes[rg_id]}"
        sub_id = f"/subscriptions/{bu.normalize(subscription_id)}"
        if sub_id in self.scopes:
            return "YES", f"SUBSCRIPTION:{self.scopes[sub_id]}"
        return "NO", ""


# ---------------------------------------------------------------------------
# Row helpers
# ---------------------------------------------------------------------------


def iac_evidence(tags):
    normalized = bu.tags_lower(tags)
    for key in IAC_TAG_KEYS:
        value = normalized.get(key, "")
        if value and any(token in value.lower() for token in IAC_TAG_VALUES):
            return "YES", f"TAG:{key}={value}"
    return "UNKNOWN", ""


def is_prod(row):
    return row["environment"] == "PROD"


def is_nonprod(row):
    return row["environment"] in ("CERT", "DESA")


def zone_count(zones):
    return len(zones) if isinstance(zones, list) else 0


class RowBuilder:
    """Builds inventory rows with the shared base, lock, IaC and resilience columns."""

    def __init__(self, ctx, managed_rgs, locks, resilience):
        self.ctx = ctx
        self.managed_rgs = managed_rgs
        self.locks = locks
        self.resilience = resilience

    def new(self, service_type, resource):
        row = {key: "" for key in INVENTORY_FIELDS}
        row["service_type"] = service_type
        row.update(bu.base_columns(self.ctx, resource, self.managed_rgs))
        sku = resource.get("sku") if isinstance(resource.get("sku"), dict) else {}
        row["sku"] = sku.get("name", "")
        row["tier"] = sku.get("tier", "")
        row["capacity"] = sku.get("capacity", "")
        row["zone_count"] = zone_count(resource.get("zones"))
        row["zone_redundant"] = "YES" if row["zone_count"] > 1 else "NO"
        row["multi_region"] = "NO"
        row["region_count"] = 1
        row["has_delete_lock"], row["lock_scope"] = self.locks.lookup(
            resource.get("id"), resource.get("subscriptionId"), resource.get("resourceGroup"))
        row["iac_managed"], row["iac_evidence"] = iac_evidence(resource.get("tags"))
        row.update(bu.resilience_columns(self.resilience, resource.get("id")))
        return row

    @staticmethod
    def finish(row, findings):
        # Platform-managed resources (AKS node RGs, Databricks, ARO, Orca) are owned by the platform.
        if row["backup_scope"] == "EXCLUDED_PLATFORM":
            row.update({"finding": "EXCLUDED_PLATFORM", "severity": "INFO", "all_findings": "EXCLUDED_PLATFORM"})
            return row
        if row["environment"] == "UNKNOWN":
            findings.add("REQUIRES_ENVIRONMENT_CLASSIFICATION", "MEDIUM")
        row.update(findings.columns())
        return row


def features(**values):
    return ";".join(f"{key}={value}" for key, value in values.items() if value not in (None, ""))


def detail_text(bag):
    if not isinstance(bag, dict):
        return ""
    return ";".join(f"{key}={value}" for key, value in bag.items() if value not in (None, "", 0, "0"))


# ---------------------------------------------------------------------------
# Resource Graph queries
# ---------------------------------------------------------------------------


def build_queries():
    queries = {
        "sites": (
            "resources | where type =~ 'microsoft.web/sites'"
            f" | project {COMMON_PROJECTION}, state = tostring(properties.state),"
            " serverFarmId = tolower(tostring(properties.serverFarmId)),"
            " linuxFx = tostring(properties.siteConfig.linuxFxVersion),"
            " redundancyMode = tostring(properties.redundancyMode),"
            " publicNetworkAccess = tostring(properties.publicNetworkAccess)"
        ),
        "slots": (
            "resources | where type =~ 'microsoft.web/sites/slots'"
            " | summarize slots = count() by id = tolower(substring(id, 0, indexof(tolower(id), '/slots/')))"
        ),
        "plans": (
            "resources | where type =~ 'microsoft.web/serverfarms'"
            f" | project {COMMON_PROJECTION}, sites = toint(properties.numberOfSites),"
            " workers = toint(properties.numberOfWorkers), zoneRedundant = tostring(properties.zoneRedundant),"
            " elastic = tostring(properties.elasticScaleEnabled),"
            " maxElastic = toint(properties.maximumElasticWorkerCount),"
            " status = tostring(properties.status), computeMode = tostring(properties.computeMode)"
        ),
        "keyvaults": (
            "resources | where type in~ ('microsoft.keyvault/vaults', 'microsoft.keyvault/managedhsms')"
            f" | project {COMMON_PROJECTION}, skuName = tostring(properties.sku.name),"
            " softDelete = tostring(properties.enableSoftDelete),"
            " softDeleteDays = tostring(properties.softDeleteRetentionInDays),"
            " purgeProtection = tostring(properties.enablePurgeProtection),"
            " publicNetworkAccess = tostring(properties.publicNetworkAccess)"
        ),
        "apim": (
            "resources | where type =~ 'microsoft.apimanagement/service'"
            f" | project {COMMON_PROJECTION}, additionalLocations = properties.additionalLocations,"
            " platformVersion = tostring(properties.platformVersion),"
            " vnetType = tostring(properties.virtualNetworkType)"
        ),
        "acr": (
            "resources | where type =~ 'microsoft.containerregistry/registries'"
            f" | project {COMMON_PROJECTION}, zoneRedundancy = tostring(properties.zoneRedundancy),"
            " softDeleteStatus = tostring(properties.policies.softDeletePolicy.status),"
            " softDeleteDays = tostring(properties.policies.softDeletePolicy.retentionDays),"
            " retentionStatus = tostring(properties.policies.retentionPolicy.status),"
            " retentionDays = tostring(properties.policies.retentionPolicy.days),"
            " exportStatus = tostring(properties.policies.exportPolicy.status)"
        ),
        "acr_replications": (
            "resources | where type =~ 'microsoft.containerregistry/registries/replications'"
            " | summarize replicas = count(), regions = make_set(location)"
            " by id = tolower(substring(id, 0, indexof(tolower(id), '/replications/')))"
        ),
        "messaging": (
            "resources | where type in~ ('microsoft.eventhub/namespaces', 'microsoft.servicebus/namespaces')"
            f" | project {COMMON_PROJECTION}, zoneRedundant = tostring(properties.zoneRedundant),"
            " status = tostring(properties.status), autoInflate = tostring(properties.isAutoInflateEnabled),"
            " maxThroughputUnits = toint(properties.maximumThroughputUnits),"
            " partitions = toint(properties.premiumMessagingPartitions),"
            " geoLocations = properties.geoDataReplication.locations"
        ),
        "eventgrid": (
            "resources | where type in~ ('microsoft.eventgrid/topics', 'microsoft.eventgrid/namespaces',"
            " 'microsoft.eventgrid/domains')"
            f" | project {COMMON_PROJECTION}, residency = tostring(properties.dataResidencyBoundary),"
            " publicNetworkAccess = tostring(properties.publicNetworkAccess)"
        ),
        "redis": (
            "resources | where type =~ 'microsoft.cache/redis'"
            f" | project {COMMON_PROJECTION}, skuName = tostring(properties.sku.name),"
            " skuFamily = tostring(properties.sku.family), skuCapacity = toint(properties.sku.capacity),"
            " replicas = toint(properties.replicasPerMaster), shardCount = toint(properties.shardCount),"
            " rdb = tostring(properties.redisConfiguration['rdb-backup-enabled']),"
            " rdbFrequency = tostring(properties.redisConfiguration['rdb-backup-frequency']),"
            " aof = tostring(properties.redisConfiguration['aof-backup-enabled']),"
            " version = tostring(properties.redisVersion), state = tostring(properties.provisioningState)"
        ),
        "redis_enterprise": (
            "resources | where type =~ 'microsoft.cache/redisenterprise'"
            f" | project {COMMON_PROJECTION}, highAvailability = tostring(properties.highAvailability),"
            " state = tostring(properties.resourceState)"
        ),
    }
    for resource_type, (kind, _, detail) in NETWORK_SPECS.items():
        queries[f"net:{kind}"] = (
            f"resources | where type =~ '{resource_type}'"
            f" | project {COMMON_PROJECTION}, detail = {detail}"
        )
    counted = ", ".join(f"'{value}'" for value in COUNTED_NETWORK_TYPES)
    unattached = " or ".join(
        f"(type =~ '{resource_type}' and {condition})" for resource_type, condition in COUNTED_NETWORK_TYPES.items())
    queries["network_counts"] = (
        f"resources | where type in~ ({counted})"
        f" | extend unattached = iff({unattached}, 1, 0)"
        " | extend standardSku = iff(tostring(sku.name) =~ 'Standard', 1, 0)"
        " | summarize total = count(), unattached = sum(unattached), standard_sku = sum(standardSku)"
        " by subscriptionId, type = tolower(type)"
        " | extend id = strcat(subscriptionId, '|', type)"
    )
    return queries


# ---------------------------------------------------------------------------
# Services
# ---------------------------------------------------------------------------


def plan_index(plans):
    return {bu.normalize(plan["id"]): plan for plan in plans}


def app_rows(builder, sites, plans_by_id, slots):
    rows = []
    for site in sites:
        kind = str(site.get("kind") or "").lower()
        if "workflowapp" in kind:
            service_type = "LOGIC_APP_STANDARD"
        elif "functionapp" in kind:
            service_type = "FUNCTION_APP"
        else:
            service_type = "APP_SERVICE"
        row = builder.new(service_type, site)
        plan = plans_by_id.get(site.get("serverFarmId") or "", {})
        plan_sku = plan.get("sku") if isinstance(plan.get("sku"), dict) else {}
        tier = str(plan_sku.get("tier") or "")
        container = "container" in kind or str(site.get("linuxFx") or "").upper().startswith(("DOCKER|", "COMPOSE|"))
        row.update({
            "resource_kind": kind,
            "sku": plan_sku.get("name", ""),
            "tier": tier,
            "capacity": plan_sku.get("capacity", ""),
            "instance_count": plan.get("workers", ""),
            "plan_id": plan.get("id", site.get("serverFarmId", "")),
            "plan_name": plan.get("name", ""),
            "zone_redundant": "YES" if str(plan.get("zoneRedundant")).lower() == "true" else "NO",
            "state": site.get("state", ""),
        })
        tier_lower = tier.lower()
        if not tier:
            capability, configured = "UNKNOWN_PLAN", "UNKNOWN"
        elif tier_lower.startswith(NOT_SUPPORTED_TIERS):
            capability, configured = "NOT_SUPPORTED", "NOT_SUPPORTED"
        elif tier_lower.startswith(AUTOMATIC_BACKUP_TIERS):
            capability, configured = "AUTOMATIC_PLATFORM_BACKUP", "YES"
        else:
            capability, configured = "UNKNOWN_TIER", "UNKNOWN"
        slot_count = slots.get(bu.normalize(site["id"]), 0)
        row.update({
            "native_backup_capability": capability,
            "native_backup_configured": configured,
            "recovery_features": features(
                automatic_backup="hourly_30d" if configured == "YES" else "", slots=slot_count or ""),
            "rebuild_dependency": "IAC+IMAGE_REGISTRY" if container else "IAC+CODE_PIPELINE",
            "detail": features(container=str(container).lower(), linux_fx=site.get("linuxFx", ""),
                               redundancy_mode=site.get("redundancyMode", ""),
                               public_network=site.get("publicNetworkAccess", "")),
            "backup_method": "NATIVE_AUTOMATIC" if configured == "YES" else "IAC_REBUILD_ONLY",
        })
        findings = bu.Findings()
        if service_type == "APP_SERVICE" and is_prod(row) and capability == "NOT_SUPPORTED":
            findings.add("PROD_APP_WITHOUT_PLATFORM_BACKUP", "LOW")
        if str(site.get("state")).lower() == "stopped" and tier_lower.startswith(AUTOMATIC_BACKUP_TIERS):
            findings.add("STOPPED_APP_ON_DEDICATED_PLAN", "LOW")
        rows.append(builder.finish(row, findings))
    return rows


def plan_rows(builder, plans):
    """App Service plans carry the compute cost; empty plans are a direct saving."""
    rows = []
    for plan in plans:
        row = builder.new("APP_SERVICE_PLAN", plan)
        row.update({
            "resource_kind": plan.get("kind", ""),
            "instance_count": plan.get("workers", ""),
            "zone_redundant": "YES" if str(plan.get("zoneRedundant")).lower() == "true" else "NO",
            "state": plan.get("status", ""),
            "detail": features(sites=plan.get("sites"), elastic=plan.get("elastic"),
                               max_elastic_workers=plan.get("maxElastic"), compute_mode=plan.get("computeMode")),
        })
        row["hosted_sites"] = plan.get("sites", "")
        findings = bu.Findings()
        tier = str(row["tier"]).lower()
        if not plan.get("sites") and not tier.startswith(("free", "dynamic", "flexconsumption")):
            findings.add("EMPTY_APP_SERVICE_PLAN", "LOW")
        if is_nonprod(row) and tier.startswith(("premium", "isolated")) and row["zone_redundant"] == "YES":
            findings.add("NONPROD_ZONE_REDUNDANT_PLAN", "LOW")
        rows.append(builder.finish(row, findings))
    return rows


def keyvault_rows(builder, vaults):
    rows = []
    for vault in vaults:
        is_hsm = "managedhsms" in vault["type"].lower()
        row = builder.new("KEY_VAULT", vault)
        soft_delete = str(vault.get("softDelete") or "").lower()
        # Soft delete is mandatory on Key Vault; an absent value means the default of 90 days.
        soft_delete_on = soft_delete != "false"
        days = vault.get("softDeleteDays") or ("90" if soft_delete_on else "")
        purge = str(vault.get("purgeProtection") or "").lower() == "true"
        row.update({
            "resource_kind": "MANAGED_HSM" if is_hsm else "KEY_VAULT",
            "sku": vault.get("skuName") or row["sku"],
            "native_backup_capability": "SOFT_DELETE+MANUAL_BACKUP_API",
            "native_backup_configured": "YES" if soft_delete_on else "NO",
            "soft_delete": "YES" if soft_delete_on else "NO",
            "soft_delete_days": days,
            "purge_protection": "YES" if purge else "NO",
            "recovery_features": features(soft_delete=f"{days}d" if soft_delete_on else "off",
                                          purge_protection="YES" if purge else "NO"),
            "rebuild_dependency": "SECRETS_NOT_REBUILDABLE_FROM_IAC",
            "detail": features(public_network=vault.get("publicNetworkAccess", "")),
            "backup_method": "NATIVE_SOFT_DELETE" if soft_delete_on else "NOT_SUPPORTED",
        })
        if "Geo-Replication" in row["resilience_solutions"]:
            row["multi_region"] = "YES"
            row["region_count"] = 2
        findings = bu.Findings()
        if is_prod(row) and not purge:
            findings.add("PROD_KEYVAULT_WITHOUT_PURGE_PROTECTION", "HIGH")
        if soft_delete_on and days and int(days) < 90:
            findings.add("SOFT_DELETE_RETENTION_BELOW_90", "LOW")
        rows.append(builder.finish(row, findings))
    return rows


def apim_rows(builder, services):
    rows = []
    for service in services:
        row = builder.new("API_MANAGEMENT", service)
        locations = service.get("additionalLocations") or []
        sku = str(row["sku"]).lower()
        regions = 1 + len(locations)
        zones = max([row["zone_count"]] + [zone_count(location.get("zones")) for location in locations])
        consumption = sku == "consumption"
        row.update({
            "resource_kind": "API_MANAGEMENT",
            "instance_count": row["capacity"],
            "zone_count": zones,
            "zone_redundant": "YES" if zones > 1 else "NO",
            "multi_region": "YES" if regions > 1 else "NO",
            "region_count": regions,
            "native_backup_capability": "NOT_SUPPORTED" if consumption else "MANUAL_BACKUP_API",
            "native_backup_configured": "NOT_SUPPORTED" if consumption else "UNKNOWN",
            "recovery_features": features(additional_regions=len(locations) or "", zones=zones or ""),
            "rebuild_dependency": "IAC+APIOPS",
            "detail": features(platform=service.get("platformVersion", ""), vnet=service.get("vnetType", ""),
                               additional_locations=",".join(loc.get("location", "") for loc in locations)),
            "backup_method": ("NATIVE_GEO_REPLICATION" if regions > 1
                              else "IAC_REBUILD_ONLY" if consumption else "MANUAL_BACKUP_API"),
        })
        findings = bu.Findings()
        if is_prod(row) and regions == 1 and zones <= 1:
            findings.add("PROD_APIM_SINGLE_REGION_NO_ZONES", "MEDIUM")
        if is_prod(row) and sku == "developer":
            findings.add("PROD_APIM_DEVELOPER_SKU", "MEDIUM")
        if is_nonprod(row) and sku == "premium":
            findings.add("NONPROD_PREMIUM_SKU", "LOW")
        rows.append(builder.finish(row, findings))
    return rows


def acr_rows(builder, registries, replications):
    rows = []
    for registry in registries:
        row = builder.new("CONTAINER_REGISTRY", registry)
        replication = replications.get(bu.normalize(registry["id"]), {})
        replica_regions = {region.lower() for region in replication.get("regions") or []}
        replica_regions.discard(str(registry.get("location") or "").lower())
        soft_delete = str(registry.get("softDeleteStatus")).lower() == "enabled"
        geo = bool(replica_regions)
        zone_redundant = str(registry.get("zoneRedundancy")).lower() == "enabled"
        row.update({
            "resource_kind": "CONTAINER_REGISTRY",
            "zone_redundant": "YES" if zone_redundant else "NO",
            "multi_region": "YES" if geo else "NO",
            "region_count": 1 + len(replica_regions),
            "native_backup_capability": "SOFT_DELETE+GEO_REPLICATION" if row["sku"] == "Premium" else "SOFT_DELETE",
            "native_backup_configured": "YES" if soft_delete or geo else "NO",
            "soft_delete": "YES" if soft_delete else "NO",
            "soft_delete_days": registry.get("softDeleteDays", "") if soft_delete else "",
            "geo_dr": "YES" if geo else "NO",
            "recovery_features": features(
                soft_delete=f"{registry.get('softDeleteDays')}d" if soft_delete else "off",
                geo_replicas=len(replica_regions) or "",
                untagged_retention=f"{registry.get('retentionDays')}d"
                if str(registry.get("retentionStatus")).lower() == "enabled" else ""),
            "rebuild_dependency": "IAC+IMAGE_PIPELINE",
            "detail": features(export_policy=registry.get("exportStatus", ""),
                               replica_regions=",".join(sorted(replica_regions))),
            "backup_method": ("NATIVE_GEO_REPLICATION" if geo else "NATIVE_SOFT_DELETE" if soft_delete
                              else "IAC_REBUILD_ONLY"),
        })
        findings = bu.Findings()
        if is_prod(row) and not soft_delete and not geo:
            findings.add("PROD_ACR_NO_SOFT_DELETE_NO_GEO", "MEDIUM")
        if is_nonprod(row) and row["sku"] == "Premium" and not geo:
            findings.add("NONPROD_PREMIUM_SKU", "LOW")
        rows.append(builder.finish(row, findings))
    return rows


def network_rows(builder, results):
    rows = []
    for resource_type, (kind, is_core, _) in NETWORK_SPECS.items():
        for resource in results.get(f"net:{kind}", []):
            row = builder.new("NETWORK", resource)
            detail = resource.get("detail") or {}
            row.update({
                "resource_kind": kind,
                "sku": row["sku"] or detail.get("sku_name", ""),
                "tier": row["tier"] or detail.get("sku_tier", ""),
                "capacity": row["capacity"] or detail.get("sku_capacity", ""),
                "native_backup_capability": "NONE",
                "native_backup_configured": "NOT_SUPPORTED",
                "recovery_features": features(delete_lock=row["has_delete_lock"]),
                "rebuild_dependency": "IAC_OR_EXPORT",
                "detail": detail_text(detail),
                "backup_method": "IAC_REBUILD_ONLY",
            })
            if kind == "LOAD_BALANCER":
                zones = int(detail.get("frontend_zones") or 0)
                row["zone_count"] = zones
                row["zone_redundant"] = "YES" if zones > 1 else "NO"
            if kind == "VIRTUAL_NETWORK_GATEWAY":
                row["zone_redundant"] = "YES" if str(row["sku"]).upper().endswith("AZ") else "NO"
            findings = bu.Findings()
            unlocked = row["has_delete_lock"] == "NO"
            if is_prod(row) and unlocked and kind in ("DNS_ZONE", "PRIVATE_DNS_ZONE"):
                findings.add("PROD_DNS_ZONE_WITHOUT_LOCK", "HIGH")
            elif is_prod(row) and unlocked and is_core:
                findings.add("PROD_NETWORK_CORE_WITHOUT_LOCK", "MEDIUM")
            if kind == "LOAD_BALANCER" and str(row["sku"]).lower() == "standard" \
                    and int(detail.get("backend_members") or 0) == 0:
                findings.add("LOAD_BALANCER_WITHOUT_BACKEND", "LOW")
            if kind == "LOAD_BALANCER" and str(row["sku"]).lower() == "basic":
                findings.add("BASIC_LOAD_BALANCER_RETIRED", "MEDIUM")
            if kind == "APPLICATION_GATEWAY" and str(detail.get("operational_state")).lower() == "stopped":
                findings.add("STOPPED_APPLICATION_GATEWAY", "LOW")
            rows.append(builder.finish(row, findings))
    return rows


def fetch_geo_dr(ctx, namespaces):
    """Geo-DR aliases for Event Hubs (Standard+) and Service Bus (Premium) namespaces."""
    candidates = []
    for namespace in namespaces:
        tier = str(bu.safe_get(namespace, "sku", "tier", default="")).lower()
        is_service_bus = "servicebus" in namespace["type"].lower()
        if (is_service_bus and tier == "premium") or (not is_service_bus and tier in ("standard", "premium")):
            candidates.append(namespace["id"])

    def fetch(namespace_id):
        aliases = ctx.try_arm("geo-dr", namespace_id, ctx.arm_list,
                              f"{namespace_id}/disasterRecoveryConfigs", MESSAGING_API)
        if aliases is None:
            return namespace_id, None
        partners = [bu.safe_get(alias, "properties", "partnerNamespace", default="") for alias in aliases]
        roles = [bu.safe_get(alias, "properties", "role", default="") for alias in aliases]
        return namespace_id, {"aliases": len(aliases), "partners": partners, "roles": roles}

    if ctx.args.skip_arm:
        return {}
    return {bu.normalize(key): value for key, value in
            (result for result in ctx.parallel(fetch, candidates, label="geo_dr") if result)}


def messaging_rows(builder, namespaces, geo_dr, eventgrid):
    rows = []
    for namespace in namespaces:
        row = builder.new("MESSAGING", namespace)
        is_service_bus = "servicebus" in namespace["type"].lower()
        locations = namespace.get("geoLocations") or []
        regions = max(1, len(locations))
        dr = geo_dr.get(bu.normalize(namespace["id"]))
        has_alias = bool(dr and dr["aliases"])
        geo = has_alias or regions > 1
        zone_redundant = str(namespace.get("zoneRedundant")).lower() == "true"
        tier = str(row["tier"] or row["sku"]).lower()
        row.update({
            "resource_kind": "SERVICE_BUS" if is_service_bus else "EVENT_HUBS",
            "instance_count": row["capacity"],
            "zone_redundant": "YES" if zone_redundant else "NO",
            "multi_region": "YES" if geo else "NO",
            "region_count": regions + (1 if has_alias else 0),
            "state": namespace.get("status", ""),
            "native_backup_capability": "GEO_DR_METADATA+GEO_REPLICATION" if tier in ("standard", "premium")
            else "NOT_SUPPORTED",
            "native_backup_configured": "YES" if geo else ("UNKNOWN" if dr is None and tier != "basic" else "NO"),
            "geo_dr": "YES" if has_alias else ("UNKNOWN" if dr is None else "NO"),
            "recovery_features": features(geo_dr_alias=has_alias and ",".join(p.split("/")[-1] for p in dr["partners"]),
                                          geo_replication_regions=regions if regions > 1 else ""),
            "rebuild_dependency": "IAC_ENTITIES_MESSAGES_TRANSIENT",
            "detail": features(auto_inflate=namespace.get("autoInflate", ""),
                               max_throughput_units=namespace.get("maxThroughputUnits") or "",
                               premium_partitions=namespace.get("partitions") or ""),
            "backup_method": "NATIVE_GEO_REPLICATION" if geo else "IAC_REBUILD_ONLY",
        })
        findings = bu.Findings()
        if is_prod(row) and tier == "premium" and not geo:
            findings.add("PROD_PREMIUM_WITHOUT_GEO_DR", "MEDIUM")
        if is_nonprod(row) and tier == "premium":
            findings.add("NONPROD_PREMIUM_SKU", "LOW")
        rows.append(builder.finish(row, findings))

    for topic in eventgrid:
        row = builder.new("MESSAGING", topic)
        kind = topic["type"].split("/")[-1].upper()
        row.update({
            "resource_kind": f"EVENT_GRID_{kind.rstrip('S')}",
            "native_backup_capability": "AUTOMATIC_REGIONAL_FAILOVER_METADATA",
            "native_backup_configured": "NO" if str(topic.get("residency")).lower() == "withinregion" else "YES",
            "recovery_features": features(data_residency=topic.get("residency", "")),
            "rebuild_dependency": "IAC_ENTITIES_MESSAGES_TRANSIENT",
            "detail": features(public_network=topic.get("publicNetworkAccess", "")),
            "backup_method": "IAC_REBUILD_ONLY",
        })
        rows.append(builder.finish(row, bu.Findings()))
    return rows


def fetch_redis_links(ctx, caches, enterprises):
    if ctx.args.skip_arm:
        return {}, {}
    premium = [cache["id"] for cache in caches if str(cache.get("skuName")).lower() == "premium"]

    def links(cache_id):
        items = ctx.try_arm("redis-links", cache_id, ctx.arm_list, f"{cache_id}/linkedServers", REDIS_API)
        return cache_id, None if items is None else [
            bu.safe_get(item, "properties", "linkedRedisCacheLocation", default="") for item in items]

    def databases(cluster_id):
        items = ctx.try_arm("redis-enterprise-db", cluster_id, ctx.arm_list,
                            f"{cluster_id}/databases", REDIS_ENTERPRISE_API)
        if items is None:
            return cluster_id, None
        aof = any(bu.safe_get(item, "properties", "persistence", "aofEnabled") for item in items)
        rdb = any(bu.safe_get(item, "properties", "persistence", "rdbEnabled") for item in items)
        geo = sum(len(bu.safe_get(item, "properties", "geoReplication", "linkedDatabases") or []) for item in items)
        return cluster_id, {"databases": len(items), "aof": aof, "rdb": rdb, "geo_links": geo}

    tasks = [("links", value) for value in premium] + [("db", value["id"]) for value in enterprises]
    results = ctx.parallel(lambda task: (task[0],) + (links if task[0] == "links" else databases)(task[1]),
                           tasks, label="redis_details")
    linked, enterprise = {}, {}
    for result in results:
        if not result:
            continue
        target = linked if result[0] == "links" else enterprise
        target[bu.normalize(result[1])] = result[2]
    return linked, enterprise


def redis_rows(builder, caches, enterprises, linked, enterprise_details):
    rows = []
    for cache in caches:
        row = builder.new("REDIS", cache)
        sku = str(cache.get("skuName") or "")
        rdb = str(cache.get("rdb")).lower() == "true"
        aof = str(cache.get("aof")).lower() == "true"
        links = linked.get(bu.normalize(cache["id"]))
        geo = bool(links)
        shards = cache.get("shardCount") or 1
        replicas = cache.get("replicas") or (1 if sku.lower() != "basic" else 0)
        row.update({
            "resource_kind": "AZURE_CACHE_FOR_REDIS",
            "sku": f"{sku} {cache.get('skuFamily', '')}{cache.get('skuCapacity', '')}".strip(),
            "tier": sku,
            "capacity": cache.get("skuCapacity", ""),
            "instance_count": shards * (1 + replicas),
            "multi_region": "YES" if geo else "NO",
            "region_count": 1 + len(links or []),
            "state": cache.get("state", ""),
            "native_backup_capability": "PERSISTENCE_RDB_AOF+GEO_REPLICATION" if sku.lower() == "premium"
            else "NOT_SUPPORTED",
            "native_backup_configured": ("YES" if rdb or aof or geo
                                         else "NOT_SUPPORTED" if sku.lower() != "premium" else "NO"),
            "geo_dr": "YES" if geo else ("UNKNOWN" if links is None and sku.lower() == "premium" else "NO"),
            "recovery_features": features(rdb=cache.get("rdbFrequency") or "on" if rdb else "",
                                          aof="on" if aof else "", geo_links=len(links or []) or ""),
            "rebuild_dependency": "IAC_CACHE_REHYDRATED_FROM_SOURCE",
            "detail": features(version=cache.get("version", ""), shards=cache.get("shardCount") or "",
                               replicas_per_primary=cache.get("replicas") or ""),
            "backup_method": ("NATIVE_GEO_REPLICATION" if geo else "NATIVE_AUTOMATIC" if rdb or aof
                              else "IAC_REBUILD_ONLY"),
        })
        findings = bu.Findings()
        if is_nonprod(row) and sku.lower() == "premium":
            findings.add("NONPROD_PREMIUM_SKU", "LOW")
        if is_prod(row) and sku.lower() == "basic":
            findings.add("PROD_REDIS_BASIC_NO_REPLICA", "MEDIUM")
        rows.append(builder.finish(row, findings))

    for cluster in enterprises:
        row = builder.new("REDIS", cluster)
        details = enterprise_details.get(bu.normalize(cluster["id"]))
        persistence = bool(details and (details["aof"] or details["rdb"]))
        geo = bool(details and details["geo_links"])
        row.update({
            "resource_kind": "AZURE_MANAGED_REDIS",
            "state": cluster.get("state", ""),
            "multi_region": "YES" if geo else "NO",
            "native_backup_capability": "PERSISTENCE_RDB_AOF+ACTIVE_GEO_REPLICATION",
            "native_backup_configured": "UNKNOWN" if details is None else ("YES" if persistence or geo else "NO"),
            "geo_dr": "UNKNOWN" if details is None else ("YES" if geo else "NO"),
            "recovery_features": features(
                high_availability=cluster.get("highAvailability", ""),
                aof="on" if details and details["aof"] else "", rdb="on" if details and details["rdb"] else ""),
            "rebuild_dependency": "IAC_CACHE_REHYDRATED_FROM_SOURCE",
            "detail": features(databases=details["databases"] if details else ""),
            "backup_method": ("NATIVE_GEO_REPLICATION" if geo else "NATIVE_AUTOMATIC" if persistence
                              else "IAC_REBUILD_ONLY"),
        })
        findings = bu.Findings()
        if is_prod(row) and str(cluster.get("highAvailability")).lower() == "disabled":
            findings.add("PROD_REDIS_HIGH_AVAILABILITY_DISABLED", "MEDIUM")
        rows.append(builder.finish(row, findings))
    return rows


def network_count_rows(ctx, counts):
    rows = []
    for row in counts:
        rows.append({
            "subscription": ctx.subscription_name(row["subscriptionId"]),
            "subscription_id": row["subscriptionId"],
            "resource_type": row["type"],
            "total": row["total"],
            "unattached": row["unattached"],
            "standard_sku": row["standard_sku"],
        })
    rows.sort(key=lambda value: (value["subscription"].lower(), value["resource_type"]))
    return rows


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    parser = bu.build_parser(__doc__.splitlines()[0], DEFAULT_OUTPUT_DIR)
    parser.add_argument("--include-locks", action=argparse.BooleanOptionalAction, default=True,
                        help="Read delete locks with one ARM GET per subscription (default: on)")
    args = parser.parse_args()
    bu.setup_logging(args.verbose)
    ctx = bu.AzureContext(args)
    ctx.load_subscriptions()

    locks = LockIndex()
    with ThreadPoolExecutor(max_workers=3) as pool:
        locks_job = pool.submit(locks.load, ctx) if args.include_locks and not args.skip_arm else None
        managed_job = pool.submit(bu.load_managed_resource_groups, ctx)
        resilience_job = pool.submit(bu.load_resilience, ctx, RESILIENCE_TYPES)
        results = ctx.queries(build_queries())
        managed_rgs = managed_job.result()
        resilience = resilience_job.result()
        if locks_job:
            locks_job.result()

    geo_dr = fetch_geo_dr(ctx, results["messaging"])
    linked, enterprise_details = fetch_redis_links(ctx, results["redis"], results["redis_enterprise"])

    builder = RowBuilder(ctx, managed_rgs, locks, resilience)
    plans_by_id = plan_index(results["plans"])
    slots = {row["id"]: row["slots"] for row in results["slots"]}
    replications = {row["id"]: row for row in results["acr_replications"]}

    inventory = []
    inventory += app_rows(builder, results["sites"], plans_by_id, slots)
    inventory += keyvault_rows(builder, results["keyvaults"])
    inventory += apim_rows(builder, results["apim"])
    inventory += acr_rows(builder, results["acr"], replications)
    inventory += network_rows(builder, results)
    inventory += messaging_rows(builder, results["messaging"], geo_dr, results["eventgrid"])
    inventory += redis_rows(builder, results["redis"], results["redis_enterprise"], linked, enterprise_details)
    plans = plan_rows(builder, results["plans"])
    counts = network_count_rows(ctx, results["network_counts"])

    output = bu.Output(args.output_dir, args.delimiter)
    output.csv("native_protection_inventory.csv", inventory, INVENTORY_FIELDS)
    output.csv("native_app_service_plans.csv", plans,
               ["service_type"] + bu.BASE_FIELDS + SIZING_FIELDS + ["hosted_sites", "has_delete_lock",
                                                                    "lock_scope", "iac_managed", "iac_evidence",
                                                                    "detail"] + bu.RESILIENCE_FIELDS
               + bu.FINDING_FIELDS)
    output.csv("native_network_counts.csv", counts)
    output.csv("native_errors.csv", ctx.errors.rows, bu.ErrorLog.FIELDS)

    counted_totals = {}
    for row in counts:
        totals = counted_totals.setdefault(row["resource_type"], {"total": 0, "unattached": 0})
        totals["total"] += row["total"]
        totals["unattached"] += row["unattached"]

    output.json("native_summary.json", bu.run_metadata(ctx, {
        "rows": len(inventory),
        "locks_loaded_subscriptions": len(locks.loaded_subscriptions),
        "service_type_by_backup_method": bu.count_by(inventory, "service_type", "backup_method"),
        "service_type_by_finding": bu.count_by(inventory, "service_type", "finding"),
        "environment_by_finding": bu.count_by(inventory, "environment", "finding"),
        "delete_lock_by_service_type": bu.count_by(inventory, "service_type", "has_delete_lock"),
        "app_service_plans_by_finding": bu.count_by(plans, "environment", "finding"),
        "network_counted_totals": counted_totals,
        "limitations": [
            "App Service custom backup configuration is not read: the API returns storage SAS URLs.",
            "API Management backups are taken on demand through the backup API and cannot be listed.",
            "Event Grid system topics, public IPs, NICs, private endpoints, NSGs and route tables "
            "are counted per subscription, not listed.",
        ],
    }))
    if args.xlsx:
        output.xlsx("native_protection_inventory.xlsx")


if __name__ == "__main__":
    main()
