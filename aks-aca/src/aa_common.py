"""Shared helpers for the AKS vs ACA scripts (aa_*.py).

The Azure client (Resource Graph paging, read-only ARM GET with retries, rate limiting, error log and
evidence of every query) lives in aa_azure.py. This module adds the aks-aca results folder, the
office-hours calendar and the Azure Monitor metrics call.
"""

import csv
import gzip
import hashlib
import json
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
                          "envoy-rw", "apollo-router", "orca-", "thanos-", "irontap", "nginx")
# Workloads that are not services of the application wherever the word appears in the name: mocks and dummies left
# in PROD, and metric exporters.
TOOL_WORKLOAD_PARTS = ("mock", "dummy", "-exporter", "dynatrace")
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
    name = str(name or "").lower()
    return name.startswith(TOOL_WORKLOAD_PREFIXES) or any(part in name for part in TOOL_WORKLOAD_PARTS)


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
# Entry points outside APIM: an application's own Application Gateway (AGIC routes straight to the pods, so neither
# APIM nor the cluster load balancer sees that traffic) and, for event-driven applications with no API on APIM, the
# messages of their Event Hubs / Service Bus namespace (aa_metrics.py).
GATEWAY_LAYERS = ("APP_GATEWAY",)
QUEUE_LAYERS = ("EVENT_HUB", "SERVICE_BUS")
APPLICATION_TRAFFIC_ORDER = TRAFFIC_LAYERS + GATEWAY_LAYERS + QUEUE_LAYERS
# BS first; an application without BS traffic takes the busiest of the other layers (never their sum). "First with
# traffic" would keep a residue: UX served through a self-hosted gateway reports almost nothing to Azure Monitor.
PRIMARY_LAYER = TRAFFIC_LAYERS[0]
FALLBACK_LAYERS = APPLICATION_TRAFFIC_ORDER[1:]


def event_driven(traffic_layer):
    """True when an application's traffic is measured on its Event Hubs / Service Bus namespace (no API on APIM).
    Those figures are messages entering the namespace, not HTTP calls: the 100 TPS rule cannot be tested on them and
    Container Apps would not bill them as requests. Such an application stays on AKS (owner's rule, 2026-11-24)."""
    return (traffic_layer or "").strip().upper() in QUEUE_LAYERS
APP_SERVICE_SUFFIX, CONTAINER_APPS_SUFFIX = ".azurewebsites.net", ".azurecontainerapps.io"


def apim_layer(name, custom_domains):
    """BS, CORE, UX, OPEN_BANKING, YAPE, SANDBOX or OTHER for an APIM service."""
    domains = [domain for domain in str(custom_domains or "").lower().split(";") if domain]
    for prefix, layer in APIM_LAYER_DOMAINS:
        if any(domain.split(".")[0].startswith(prefix) for domain in domains):
            return layer
    if domains and all(domain.startswith("sandbox.") for domain in domains):
        return "SANDBOX"  # a sandbox for third parties: test traffic, never measured
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


class QueryCache:
    """Answers of read-only queries over a finished time window, kept on disk (gzip JSON) so a rerun or a resumed run
    does not ask Azure again. The key is a hash of everything that defines the query (resource, query text, window,
    parameters), so a changed query is asked again. A window that ends less than CACHE_SETTLE ago is never cached:
    late data may still arrive. Failed queries raise before anything is written, so they are retried next run.
    Delete the _cache folder, or pass --no-cache, to ask everything again."""

    def __init__(self, folder, enabled=True):
        self.folder, self.enabled = folder, enabled
        self.hits = self.misses = 0

    def get(self, parts, window_end, run):
        if not self.enabled or window_end > datetime.now(timezone.utc) - CACHE_SETTLE:
            return run()
        key = hashlib.sha256(json.dumps(parts, sort_keys=True, default=str).encode()).hexdigest()[:32]
        path = os.path.join(self.folder, key[:2], key + ".json.gz")
        if os.path.exists(path):
            with gzip.open(path, "rt", encoding="utf-8") as handle:
                self.hits += 1
                return json.load(handle)
        value = run()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        partial = f"{path}.{os.getpid()}.tmp"
        with gzip.open(partial, "wt", encoding="utf-8") as handle:
            json.dump(value, handle)
        os.replace(partial, path)
        self.misses += 1
        return value

    def summary(self):
        return {"enabled": self.enabled, "hits": self.hits, "queried": self.misses}


CACHE_SETTLE = timedelta(hours=6)


def add_cache_argument(parser):
    parser.add_argument("--no-cache", action="store_true",
                        help="Ask Azure again for every query, ignoring and not writing <output-dir>/_cache")
    return parser


def query_cache(args):
    return QueryCache(os.path.join(args.output_dir, "_cache"), enabled=not args.no_cache)


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
    url = f"{resource_id}/providers/Microsoft.Insights/metrics"
    fetch = lambda: ctx.arm_get(url, METRICS_API, params).get("value", [])
    cache = getattr(ctx, "query_cache", None)  # set by scripts that cache (aa_metrics.py)
    return cache.get(("metrics", resource_id.lower(), METRICS_API, params), end, fetch) if cache else fetch()


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


def keep_version(path):
    """Copy an output about to be overwritten to versions/ next to it, named by its last change
    (aks_aca_corporate_v2_2026-11-26_1300.pptx), so earlier decks and workbooks stay at hand."""
    if not os.path.exists(path):
        return None
    import shutil
    folder = os.path.join(os.path.dirname(os.path.abspath(path)), "versions")
    os.makedirs(folder, exist_ok=True)
    stem, extension = os.path.splitext(os.path.basename(path))
    stamp = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d_%H%M")
    target = os.path.join(folder, f"{stem}_{stamp}{extension}")
    if not os.path.exists(target):
        shutil.copy2(path, target)
    return target
