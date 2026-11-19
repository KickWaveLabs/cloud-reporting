"""Shared helpers for the AKS vs ACA scripts (aa_*.py).

The Azure client (Resource Graph paging, read-only ARM GET with retries, rate limiting, error log and
evidence of every query) lives in aa_azure.py. This module adds the aks-aca results folder, the
office-hours calendar and the Azure Monitor metrics call.
"""

import csv
import os
import re
from datetime import date, datetime, timedelta, timezone

import aa_azure as azure

SRC_DIR = os.path.dirname(os.path.abspath(__file__))
TOPIC_DIR = os.path.dirname(SRC_DIR)

RESULTS_DIR = os.path.join(TOPIC_DIR, "results")

# Lima has no daylight saving time: UTC-5 all year.
LIMA = timezone(timedelta(hours=-5), "America/Lima")
OFFICE_START_HOUR = 8
OFFICE_END_HOUR = 20  # exclusive: 8:00 to 19:59
METRICS_RETENTION_DAYS = 93
METRICS_API = "2023-10-01"
INTERVALS = {"PT5M": timedelta(minutes=5), "PT1H": timedelta(hours=1)}
MIN_SPLIT_STEPS = 12

AKS_TYPE = "microsoft.containerservice/managedclusters"
ACA_TYPES = ("microsoft.app/containerapps", "microsoft.app/managedenvironments", "microsoft.app/jobs")
SYSTEM_NAMESPACES = ("kube-system", "kube-public", "kube-node-lease", "gatekeeper-system", "calico-system",
                     "tigera-operator", "aks-command", "app-routing-system", "dataprotection-microsoft")


# ---------------------------------------------------------------------------
# What counts as an application: one definition for every script
# ---------------------------------------------------------------------------
# Namespaces of the platform and of agents that run in every cluster (security, Terraform, GitOps).
PLATFORM_NAMESPACES = set(SYSTEM_NAMESPACES) | {"orcasecurity", "tfc-system", "flux-system", "argocd", "cert-manager",
                                                "keda", "monitoring", "default", "ibm-finops-agent", "reflector",
                                                "synthetic-monitoring", "operators", "prometheus", "thanos"}
PLATFORM_NAMESPACE_PREFIXES = ("aks-istio", "istio-", "calico", "tigera", "gatekeeper")
# Workloads that are tools, not microservices, even in an application namespace (name prefix).
TOOL_WORKLOAD_PREFIXES = ("pod-restart-cronjob", "agic-ingress", "otel-", "opentelemetry", "shgw", "chaos",
                          "arc-", "actions-runner", "github-runner", "tfc-agent", "sonarqube", "demo-sonarqube",
                          "connectivity-toolbox", "grafana", "alloy", "elasticsearch", "kibana", "logstash", "memcached",
                          "envoy-rw", "apollo-router", "orca-", "thanos-", "irontap")
# Applications whose clusters only run infrastructure tools: excluded from every application figure.
INFRASTRUCTURE_APPS = {
    "COAA": "GitHub Actions runners", "VSDP": "GitHub Actions runners", "HTFC": "Terraform Cloud agents",
    "INCT": "SonarQube", "O11Y": "Observability (Grafana, Alloy)", "ELKP": "Observability (Elasticsearch, Kibana)",
    "APIG": "APIM self-hosted gateway", "GFED": "Shared GraphQL gateway (Apollo Router)",
}
# Cluster categories. Only APPLICATION clusters enter the analysis.
APPLICATION, INFRASTRUCTURE, NO_APPLICATIONS, NO_DATA = "APPLICATION", "INFRASTRUCTURE", "NO_APPLICATIONS", "NO_DATA"


def is_platform_namespace(namespace):
    name = str(namespace or "").lower()
    return name in PLATFORM_NAMESPACES or name.startswith(PLATFORM_NAMESPACE_PREFIXES) or "ingress" in name


# Tools that belong to the application itself (its own API gateway, ingress and telemetry): a cluster that only
# runs these is still an application cluster.
APP_GATEWAY_PREFIXES = ("shgw", "agic-ingress", "otel-", "opentelemetry")


def is_tool_workload(name):
    return str(name or "").lower().startswith(TOOL_WORKLOAD_PREFIXES)


def is_app_gateway(name):
    return str(name or "").lower().startswith(APP_GATEWAY_PREFIXES)


def is_microservice(namespace, workload):
    """A deployment that belongs to an application: outside platform namespaces and not a known tool."""
    return not is_platform_namespace(namespace) and not is_tool_workload(workload)


# AKS ingress hosts carry the app code: azingressscrmeu2.lima.bcp.com.pe -> SCRM. A region right after
# "azingress" (azingresseu2yape...) is not a code.
INGRESS_COD_APP = re.compile(r"^azingress(?![a-z]{2}\d)([a-z0-9]{4})")


def backend_cod_app(host):
    """App code of the AKS ingress behind an API; empty for other backends (App Service, external)."""
    match = INGRESS_COD_APP.match(str(host or "").lower())
    return match.group(1).upper() if match else ""


# APIM layers of the north-south chain: a user call enters UX, UX calls BS, BS calls Core, Core calls the mainframe.
# Open Banking and Yape run their own APIM. The layer comes from the gateway's custom domain (apibs.lima.bcp.com.pe).
APIM_LAYER_DOMAINS = (("apibs", "BS"), ("apicr", "CORE"), ("apiux", "UX"), ("apisux", "UX"))
# Layers that measure an application's traffic, in order of preference: one layer per application, never added.
# BS is the business transaction. UX (the channel), Open Banking and Yape (external entry points) stand in only for
# applications without BS traffic, since their calls go on to BS and would count twice. Core is internal (BS ->
# Core -> mainframe): last, only for applications that have traffic nowhere else.
TRAFFIC_LAYERS = ("BS", "UX", "OPEN_BANKING", "YAPE", "CORE")
APP_SERVICE_SUFFIX, CONTAINER_APPS_SUFFIX = ".azurewebsites.net", ".azurecontainerapps.io"


def apim_layer(name, custom_domains):
    """BS, CORE, UX, OPEN_BANKING, YAPE or OTHER for an APIM service."""
    domains = [domain for domain in str(custom_domains or "").lower().split(";") if domain]
    for prefix, layer in APIM_LAYER_DOMAINS:
        if any(domain.split(".")[0].startswith(prefix) for domain in domains):
            return layer
    text = str(name or "").lower() + ";" + ";".join(domains)
    if "opbk" in text or "apisbcp" in text or "openbanking" in text:
        return "OPEN_BANKING"
    return "YAPE" if "yape" in text else "OTHER"


def backend_platform(host):
    """Where the backend of an API runs, from its host."""
    host = str(host or "").lower()
    if not host:
        return "NONE"
    if host.startswith("azingress") or host.endswith(".svc.cluster.local"):
        return "AKS"
    if host.endswith(APP_SERVICE_SUFFIX):
        return "APP_SERVICE"
    if host.endswith(CONTAINER_APPS_SUFFIX):
        return "CONTAINER_APPS"
    if ".apps.ocp" in host:
        return "OPENSHIFT"
    # A gateway calling another gateway (UX -> apibs.lima.bcp.com.pe): the traffic is counted on that APIM.
    if host.endswith(".azure-api.net") or host.split(".")[0].startswith(tuple(prefix for prefix, _ in APIM_LAYER_DOMAINS)):
        return "APIM"
    return "OTHER"


NAMING_COD_APP = re.compile(r"^(?:bs-|cr-|ux-|business-)?([a-z0-9]{4})-")


def resolve_api_cod_app(host, api_id, known_codes, resource_codes):
    """(code, cod_app_source) of an API, from its backend host first and its name last.

    HOST             the ingress regex of backend_cod_app (azingressscrmeu2... -> SCRM)
    HOST_NORMALIZED  an ingress host whose code starts like a region or follows one (azingresscb2ceu201 -> CB2C,
                     azingresseu2yape -> YAPE); accepted only for an app code of the inventory
    K8S_NAMESPACE    <service>.<namespace>.svc.cluster.local, the namespace being an app code of the inventory. A
                     namespace that is a domain (mbbk, nhbk run in BADI clusters) stays empty here; aa_k8s_snapshot.py
                     maps it to its cluster through Container Insights
    BACKEND_RESOURCE the App Service or Container App the host names, with that resource's own code
    NAMING           bs-srec-..., business-ntlc-... in the API id or the host's first label; an app code of the
                     inventory, used only when the host gives nothing and never for an AKS backend (bs-auth-... is
                     BSEG's API, channel-mfcc-...mbbk.svc.cluster.local runs in a BADI cluster)
    known_codes: app codes of the inventory (upper case). resource_codes: resource name (lower case) -> code.
    """
    host = str(host or "").lower()
    steps = (("HOST", lambda: backend_cod_app(host)),
             ("HOST_NORMALIZED", lambda: normalized_ingress_code(host, known_codes)),
             ("K8S_NAMESPACE", lambda: namespace_code(host, known_codes)),
             ("BACKEND_RESOURCE", lambda: backend_resource_code(host, resource_codes)),
             ("NAMING", lambda: "" if backend_platform(host) == "AKS"
              else naming_code((api_id, host.split(".")[0]), known_codes)))
    for source, step in steps:
        code = step()
        if code:
            return code, source
    return "", ""


def normalized_ingress_code(host, known_codes):
    if not host.startswith("azingress"):
        return ""
    rest = host[len("azingress"):]
    candidates = [rest[:4]] + ([rest[3:7]] if re.match(r"[a-z]{2}\d", rest) else [])
    return next((value.upper() for value in candidates if value.upper() in known_codes), "")


def namespace_code(host, known_codes):
    match = re.match(r"^[a-z0-9-]+\.([a-z0-9-]+)\.svc\.cluster\.local$", host)
    return match.group(1).upper() if match and match.group(1).upper() in known_codes else ""


def backend_resource_code(host, resource_codes):
    name = host.split(".")[0]
    if host.endswith(APP_SERVICE_SUFFIX):
        return resource_codes.get(name) or azure.cod_app_from_name(name)
    return resource_codes.get(name, "") if host.endswith(CONTAINER_APPS_SUFFIX) else ""


def naming_code(texts, known_codes):
    for text in texts:
        match = NAMING_COD_APP.match(str(text or "").lower())
        if match and match.group(1).upper() in known_codes:
            return match.group(1).upper()
    return ""


def cluster_ids_by_name(clusters):
    """Cluster name (lowercase) -> list of resource ids (lowercase), from aks_clusters.csv rows.

    Two clusters can share a name in different subscriptions, so a name maps to a list.
    """
    ids = {}
    for row in clusters:
        ids.setdefault(row["name"].lower(), []).append(row["id"].lower())
    return ids


def cluster_id_of(row, ids_by_name):
    """Resource id of the row's cluster: its own cluster_id column, else the only inventory cluster with that name.

    Empty when the name is unknown or shared by several clusters.
    """
    own = str(row.get("cluster_id") or "").lower()
    if own:
        return own
    ids = ids_by_name.get(str(row.get("cluster") or "").lower(), [])
    return ids[0] if len(ids) == 1 else ""


def results_path(*parts):
    return os.path.join(RESULTS_DIR, *parts)


def read_csv(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def is_office(moment):
    """Monday to Friday, 8:00 to 20:00 Lima time. Public holidays are not excluded."""
    local = moment.astimezone(LIMA)
    return local.weekday() < 5 and OFFICE_START_HOUR <= local.hour < OFFICE_END_HOUR


def lima_month(moment):
    return moment.astimezone(LIMA).strftime("%Y-%m")


def last_full_month():
    """(first day, last day) of the last full calendar month in Lima, as Cost Management's "last month"."""
    last_day = datetime.now(LIMA).date().replace(day=1) - timedelta(days=1)
    return last_day.replace(day=1), last_day


def metrics_window(start=None, end=None):
    """Default window: the last full calendar month (Lima). With only --end, from the first day of that
    month; with only --start, up to yesterday. Never older than the 93-day metrics retention.

    Returns (start, end) as UTC datetimes; end is exclusive (midnight Lima after the last day).
    """
    today = datetime.now(LIMA).date()
    month_first, month_last = last_full_month()
    if not start and not end:
        first_day, last_day = month_first, month_last
    else:
        last_day = date.fromisoformat(end) if end else today - timedelta(days=1)
        first_day = date.fromisoformat(start) if start else last_day.replace(day=1)
    oldest = today - timedelta(days=METRICS_RETENTION_DAYS - 1)
    if first_day < oldest:
        azure.LOG.warning("Start %s is older than the metrics retention; using %s", first_day, oldest)
        first_day = oldest
    begin = datetime(first_day.year, first_day.month, first_day.day, tzinfo=LIMA).astimezone(timezone.utc)
    finish = datetime(last_day.year, last_day.month, last_day.day, tzinfo=LIMA) + timedelta(days=1)
    return begin, finish.astimezone(timezone.utc)


def cod_app_by_cluster(clusters):
    """Cluster name (lowercase) -> app code, from aks_clusters.csv rows."""
    return {row["name"].lower(): row.get("cod_app", "") for row in clusters}


def cod_app_by_api(apis):
    """(APIM, API id) -> app code of the AKS ingress behind the API, from apim_apis.csv rows."""
    return {(row["apim"], row["api_id"]): row.get("backend_cod_app", "") for row in apis}


def with_cod_app(rows, after, code_of):
    """Copy rows adding a cod_app column right after column `after`; code_of(row) returns the code."""
    return insert_column(rows, after, "cod_app", code_of)


def insert_column(rows, after, name, value_of):
    """Copy rows adding column `name` right after column `after` (at the end if absent); value_of(row) gives it."""
    result = []
    for row in rows:
        copy = {}
        for key, value in row.items():
            copy[key] = value
            if key == after:
                copy[name] = value_of(row)
        copy.setdefault(name, value_of(row))
        result.append(copy)
    return result


def cluster_cod_app(lookup):
    """code_of for with_cod_app on rows keyed by `cluster`: inventory code, else the naming convention."""
    return lambda row: lookup.get(str(row.get("cluster", "")).lower()) or azure.cod_app_from_name(row.get("cluster"))


def chunks(start, end, days):
    """Split [start, end) into consecutive chunks of `days` days."""
    current = start
    while current < end:
        following = min(end, current + timedelta(days=days))
        yield current, following
        current = following


def iso_z(moment):
    return moment.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def get_metrics(ctx, resource_id, names, start, end, interval, aggregation, metric_filter="", top=None):
    """Azure Monitor metrics for one resource (GET, same data as `az monitor metrics list`).

    Returns the `value` list: one entry per metric, each with `timeseries` split by the filter dimensions.
    """
    params = {
        "metricnames": ",".join(names),
        "timespan": f"{iso_z(start)}/{iso_z(end)}",
        "interval": interval,
        "aggregation": aggregation,
    }
    if metric_filter:
        params["$filter"] = metric_filter
    if top:
        params["top"] = str(top)
    body = ctx.arm_get(f"{resource_id}/providers/Microsoft.Insights/metrics", METRICS_API, params)
    return body.get("value", [])


def get_metrics_split(ctx, resource_id, names, start, end, interval, aggregation, metric_filter="", top=None):
    """get_metrics for [start, end), as one `value` list per piece, in time order.

    Azure Monitor rejects responses over 8 MB (large clusters, APIM with many APIs), so the window is
    halved on whole intervals until each piece fits or is MIN_SPLIT_STEPS intervals long.
    """
    try:
        return [get_metrics(ctx, resource_id, names, start, end, interval, aggregation, metric_filter, top)]
    except azure.ArmError as error:
        step = INTERVALS[interval]
        if "exceeded maximum limit" not in str(error.message) or end - start < 2 * MIN_SPLIT_STEPS * step:
            raise
        middle = start + ((end - start) / 2 // step) * step
        return (get_metrics_split(ctx, resource_id, names, start, middle, interval, aggregation, metric_filter, top)
                + get_metrics_split(ctx, resource_id, names, middle, end, interval, aggregation, metric_filter, top))


def series_points(series, aggregation):
    """Yield (utc datetime, value) for the points of one time series that have a value."""
    key = aggregation.lower()
    for point in series.get("data", []):
        value = point.get(key)
        if value is None:
            continue
        yield azure.parse_datetime(point["timeStamp"]), value


def series_dimensions(series):
    return {str(item["name"]["value"]).lower(): item.get("value", "") for item in series.get("metadatavalues", [])}


def cluster_environment(row, subscription_name=""):
    """Environment from the tag, then the name suffix (P/C/D + digits), then the subscription name."""
    environment, evidence = azure.normalize_environment(row.get("tags"), "")
    if environment != "UNKNOWN":
        return environment, evidence
    from_name = azure.environment_from_name(row.get("name"))
    if from_name:
        return from_name, "NAME_SUFFIX"
    return azure.normalize_environment({}, subscription_name)


def is_poc_subscription(name):
    """POC subscriptions ("DTI - POC - ...", "POC - APP - ...") are not part of the estate this report describes."""
    return re.search(r"\bpoc\b", name or "", re.IGNORECASE) is not None


def add_common_arguments(parser):
    parser.add_argument("--environments", default="PROD",
                        help="Comma-separated environments to include (PROD, CERT, DESA, UNKNOWN). 'ALL' for every one")
    parser.add_argument("--inventory", default=results_path("aa-inventory"),
                        help="Folder with the aa_inventory.py output")
    return parser


def wanted_environments(args):
    value = (args.environments or "").strip().upper()
    return None if value in ("", "ALL") else {item.strip() for item in value.split(",") if item.strip()}
