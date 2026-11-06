"""Shared helpers for the backup inventory scripts (bu_*_inventory.py).

Covers: Azure Resource Graph paging with concurrency, read-only ARM GET calls,
tag normalization (environment, codApp, operation, vendor), managed platform
classification, the Azure Backup protection index (Recovery Services vaults and
Backup vaults), policy summaries, resilience data from Business Continuity Center,
findings and CSV/JSON/XLSX output.

Dependencies:
  pip install azure-identity azure-mgmt-resourcegraph requests
  pip install openpyxl   # optional, only for --xlsx
"""

import argparse
import csv
import json
import logging
import os
import re
import sys
import threading
import time
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone

import requests
from azure.core.exceptions import HttpResponseError
from azure.identity import DefaultAzureCredential
from azure.mgmt.resourcegraph import ResourceGraphClient
from azure.mgmt.resourcegraph.models import QueryRequest, QueryRequestOptions, ResultFormat

ARM_ENDPOINT = "https://management.azure.com"
ARM_SCOPE = "https://management.azure.com/.default"

LOG = logging.getLogger("bu")

# ---------------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------------


def utc_now():
    return datetime.now(timezone.utc)


def normalize(value):
    return str(value or "").strip().rstrip("/").lower()


def safe_get(dictionary, *keys, default=None):
    current = dictionary
    for key in keys:
        if not isinstance(current, dict):
            return default
        current = current.get(key)
        if current is None:
            return default
    return current


def first_value(*values):
    for value in values:
        if value not in (None, "", [], {}):
            return value
    return None


def yes_no(value):
    if value is None or value == "":
        return "UNKNOWN"
    return "YES" if value else "NO"


def parse_datetime(value):
    """Parse ARM timestamps. Returns None for empty or .NET minimum values."""
    if not value or str(value).startswith("0001-01-01"):
        return None
    text = str(value).replace("Z", "+00:00")
    # Trim fractional seconds beyond microseconds, which fromisoformat rejects on older Pythons.
    text = re.sub(r"(\.\d{6})\d+", r"\1", text)
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError:
        return None
    return parsed if parsed.tzinfo else parsed.replace(tzinfo=timezone.utc)


def iso(value):
    parsed = value if isinstance(value, datetime) else parse_datetime(value)
    return parsed.strftime("%Y-%m-%dT%H:%M:%SZ") if parsed else ""


def age_hours(value):
    parsed = value if isinstance(value, datetime) else parse_datetime(value)
    if not parsed:
        return None
    return round((utc_now() - parsed).total_seconds() / 3600, 1)


_TIMESPAN = re.compile(r"^(?:(\d+)\.)?(\d{1,2}):(\d{2}):(\d{2})(?:\.\d+)?$")


def timespan_days(value):
    """Convert a .NET TimeSpan string ("30.00:00:00", "12:00:00") to days.

    Azure Backup uses TimeSpan.MinValue ("-10675199...") for "not configured".
    """
    if not value or str(value).startswith("-"):
        return None
    match = _TIMESPAN.match(str(value))
    if not match:
        return None
    days, hours, minutes, seconds = (int(part or 0) for part in match.groups())
    return round(days + hours / 24 + minutes / 1440 + seconds / 86400, 4)


_ISO_DURATION = re.compile(
    r"^P(?:(\d+)Y)?(?:(\d+)M)?(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$"
)


def iso_duration_days(value):
    """Convert an ISO 8601 duration ("P30D", "P12W", "PT4H") to approximate days."""
    if not value:
        return None
    match = _ISO_DURATION.match(str(value).upper())
    if not match:
        return None
    years, months, weeks, days, hours, minutes, seconds = (int(part or 0) for part in match.groups())
    total = years * 365 + months * 30 + weeks * 7 + days + hours / 24 + minutes / 1440 + seconds / 86400
    return round(total, 4)


DURATION_TYPE_DAYS = {"days": 1, "weeks": 7, "months": 30, "years": 365}


def retention_count_days(duration):
    """Convert {"count": 12, "durationType": "Weeks"} to days."""
    if not isinstance(duration, dict):
        return None
    count = duration.get("count")
    factor = DURATION_TYPE_DAYS.get(str(duration.get("durationType") or "").lower())
    if count is None or factor is None:
        return None
    return int(count) * factor


def parse_resource_id(resource_id):
    """Return subscription, resource group and name parts of an ARM id."""
    parts = str(resource_id or "").strip("/").split("/")
    lowered = [part.lower() for part in parts]
    result = {"subscription_id": "", "resource_group": "", "name": parts[-1] if parts else ""}
    if "subscriptions" in lowered:
        index = lowered.index("subscriptions")
        if index + 1 < len(parts):
            result["subscription_id"] = parts[index + 1]
    if "resourcegroups" in lowered:
        index = lowered.index("resourcegroups")
        if index + 1 < len(parts):
            result["resource_group"] = parts[index + 1]
    return result


def parent_resource_ids(resource_id):
    """Yield the ids of parent resources, nearest first.

    /.../providers/Microsoft.Sql/servers/s1/databases/d1 -> /.../providers/Microsoft.Sql/servers/s1
    """
    parts = normalize(resource_id).strip("/").split("/")
    if "providers" not in parts:
        return
    provider_index = len(parts) - 1 - parts[::-1].index("providers")
    # After "providers/<namespace>" the id alternates type/name pairs.
    base = provider_index + 2
    pairs = (len(parts) - base) // 2
    for keep in range(pairs - 1, 0, -1):
        yield "/" + "/".join(parts[: base + keep * 2])


# ---------------------------------------------------------------------------
# Error collection
# ---------------------------------------------------------------------------


class ErrorLog:
    """Thread-safe list of access gaps and API errors, written to *_errors.csv."""

    FIELDS = ["timestamp", "stage", "scope", "operation", "status", "code", "message"]

    def __init__(self):
        self._lock = threading.Lock()
        self.rows = []

    def add(self, stage, scope, operation, status="", code="", message=""):
        row = {
            "timestamp": iso(utc_now()),
            "stage": stage,
            "scope": scope,
            "operation": operation,
            "status": status,
            "code": code,
            "message": str(message or "")[:400].replace("\n", " "),
        }
        with self._lock:
            self.rows.append(row)
        LOG.debug("error %s %s %s %s", stage, status, code, scope)

    def count_by_status(self):
        counts = defaultdict(int)
        for row in self.rows:
            counts[str(row["status"] or "n/a")] += 1
        return dict(counts)


# ---------------------------------------------------------------------------
# Evidence: what was queried, so every figure can be traced to its source
# ---------------------------------------------------------------------------

_GUID = re.compile(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}")


def arm_pattern(url):
    """Generic form of an ARM URL: ids and names replaced by placeholders.

    /subscriptions/<id>/resourceGroups/rg1/providers/Microsoft.Sql/servers/s1/databases/d1
    -> /subscriptions/{sub}/resourceGroups/{rg}/providers/Microsoft.Sql/servers/{name}/databases/{name}
    """
    path = url.replace(ARM_ENDPOINT, "").split("?")[0]
    parts = path.strip("/").split("/")
    out, index = [], 0
    while index < len(parts):
        part = parts[index]
        lowered = part.lower()
        out.append(part)
        if lowered == "subscriptions" and index + 1 < len(parts):
            out.append("{sub}")
            index += 2
            continue
        if lowered == "resourcegroups" and index + 1 < len(parts):
            out.append("{rg}")
            index += 2
            continue
        if lowered == "providers" and index + 1 < len(parts):
            out.append(parts[index + 1])
            index += 2
            # After the namespace the path alternates type/name.
            while index < len(parts):
                out.append(parts[index])
                if index + 1 < len(parts):
                    name = parts[index + 1]
                    out.append(name if name.lower() in ("default", "current") else "{name}")
                index += 2
            break
        index += 1
    return "/" + "/".join(out)


class Evidence:
    """Thread-safe record of Resource Graph queries, ARM call patterns and request bodies.

    Written into each *_summary.json under "evidence" so a reviewer can rerun or check every source.
    """

    def __init__(self):
        self._lock = threading.Lock()
        self.queries = []
        self.calls = defaultdict(lambda: {"calls": 0, "errors": 0, "statuses": Counter()})
        self.bodies = {}

    def add_query(self, label, query, rows, subscriptions, seconds):
        with self._lock:
            self.queries.append({
                "source": "Azure Resource Graph", "label": label, "query": " ".join(query.split()),
                "rows": rows, "subscriptions": subscriptions, "seconds": round(seconds, 1),
                "timestamp": iso(utc_now()),
            })

    def add_call(self, method, url, api_version, status):
        key = (method, arm_pattern(url), api_version or "")
        with self._lock:
            entry = self.calls[key]
            entry["calls"] += 1
            entry["statuses"][str(status)] += 1
            if status and int(status) >= 400:
                entry["errors"] += 1

    def add_body(self, label, url, api_version, body):
        with self._lock:
            self.bodies.setdefault(label, {"endpoint": arm_pattern(url), "api_version": api_version, "body": body})

    def as_dict(self):
        with self._lock:
            return {
                "resource_graph_queries": list(self.queries),
                "arm_calls": [{"method": method, "pattern": pattern, "api_version": api_version, "calls": value["calls"],
                               "errors": value["errors"], "statuses": dict(value["statuses"])}
                              for (method, pattern, api_version), value in sorted(self.calls.items())],
                "request_bodies": [dict(label=label, **value) for label, value in self.bodies.items()],
            }


# ---------------------------------------------------------------------------
# Throttling
# ---------------------------------------------------------------------------


class RateLimiter:
    """Spaces out calls so concurrent workers stay under a requests-per-second budget."""

    def __init__(self, per_second):
        self._interval = 1.0 / per_second if per_second > 0 else 0
        self._lock = threading.Lock()
        self._next = 0.0

    def wait(self):
        if not self._interval:
            return
        with self._lock:
            now = time.monotonic()
            start = max(now, self._next)
            self._next = start + self._interval
        delay = start - time.monotonic()
        if delay > 0:
            time.sleep(delay)


# ---------------------------------------------------------------------------
# Azure clients
# ---------------------------------------------------------------------------


class ArmError(Exception):
    def __init__(self, status, code, message, url):
        super().__init__(f"{status} {code}: {message}")
        self.status = status
        self.code = code
        self.message = message
        self.url = url


class AzureContext:
    """Credential, Resource Graph client, ARM GET client, subscriptions and error log."""

    def __init__(self, args):
        self.args = args
        self.errors = ErrorLog()
        self.credential = DefaultAzureCredential(exclude_interactive_browser_credential=True)
        self.graph = ResourceGraphClient(self.credential)
        # Resource Graph allows 15 queries per 5 seconds per user.
        self.graph_limiter = RateLimiter(per_second=2.8)
        self.arm_limiter = RateLimiter(per_second=args.arm_rps)
        self._token = None
        self._token_lock = threading.Lock()
        self._local = threading.local()
        self.subscriptions = {}
        self.subscription_filter = []
        self.started = utc_now()
        self.evidence = Evidence()

    # -- subscriptions -----------------------------------------------------

    def load_subscriptions(self):
        rows = self.query(
            "resourcecontainers"
            " | where type =~ 'microsoft.resources/subscriptions'"
            " | project id, subscriptionId, name, state = tostring(properties.state)",
            partition=False,
            label="subscriptions",
        )
        wanted = {normalize(value) for value in (self.args.subscriptions or "").split(",") if value.strip()}
        for row in rows:
            sub_id = row["subscriptionId"]
            if wanted and normalize(sub_id) not in wanted and normalize(row["name"]) not in wanted:
                continue
            self.subscriptions[normalize(sub_id)] = {
                "id": sub_id,
                "name": row["name"],
                "state": row.get("state", ""),
            }
        if wanted:
            self.subscription_filter = [value["id"] for value in self.subscriptions.values()]
            missing = wanted - set(self.subscriptions) - {normalize(v["name"]) for v in self.subscriptions.values()}
            for item in sorted(missing):
                self.errors.add("subscriptions", item, "filter", code="NotVisible",
                                message="Subscription not visible to the signed-in identity")
        LOG.info("Subscriptions in scope: %d", len(self.subscriptions))
        return self.subscriptions

    def subscription_name(self, subscription_id):
        return safe_get(self.subscriptions, normalize(subscription_id), "name", default="")

    # -- Resource Graph ----------------------------------------------------

    def _graph_page(self, query, subscriptions, skip_token, label):
        options = QueryRequestOptions(result_format=ResultFormat.OBJECT_ARRAY, top=1000, skip_token=skip_token)
        request = QueryRequest(subscriptions=subscriptions or None, query=query, options=options)
        for attempt in range(1, 7):
            self.graph_limiter.wait()
            try:
                return self.graph.resources(request)
            except HttpResponseError as error:
                status = getattr(error, "status_code", None)
                if status in (429, 500, 502, 503, 504) and attempt < 6:
                    delay = min(60, 5 * attempt)
                    LOG.warning("Resource Graph %s on %s, retry in %ss", status, label, delay)
                    time.sleep(delay)
                    continue
                raise
        raise RuntimeError("unreachable")

    def _graph_all(self, query, subscriptions, label):
        rows = []
        skip_token = None
        while True:
            response = self._graph_page(query, subscriptions, skip_token, label)
            rows.extend(response.data or [])
            skip_token = response.skip_token
            if not skip_token:
                break
        if getattr(response, "result_truncated", None) == "true":
            self.errors.add("resource-graph", label, "query", code="ResultTruncated",
                            message="Resource Graph truncated the result; project the id column")
        return rows

    def query(self, query, partition=True, label="query"):
        """Run a Resource Graph query across the subscriptions in scope.

        partition=True splits the subscriptions in chunks queried concurrently. Use
        partition=False for summarize queries, whose groups must not be split.
        Row queries should project an `id` column so paging works.
        """
        subscriptions = self.subscription_filter or [value["id"] for value in self.subscriptions.values()]
        started = time.monotonic()
        try:
            if not partition or len(subscriptions) <= 50:
                rows = self._graph_all(query, subscriptions, label)
            else:
                chunk_count = max(1, min(self.args.graph_parallel, len(subscriptions) // 25))
                size = -(-len(subscriptions) // chunk_count)
                chunks = [subscriptions[i:i + size] for i in range(0, len(subscriptions), size)]
                rows = []
                with ThreadPoolExecutor(max_workers=len(chunks)) as pool:
                    for result in pool.map(lambda chunk: self._graph_all(query, chunk, label), chunks):
                        rows.extend(result)
        except HttpResponseError as error:
            code = getattr(getattr(error, "error", None), "code", "") or ""
            self.errors.add("resource-graph", label, "query", getattr(error, "status_code", ""), code, error.message)
            LOG.error("Resource Graph query %s failed: %s", label, error.message)
            return []
        LOG.info("Resource Graph %-32s %6d rows in %.1fs", label, len(rows), time.monotonic() - started)
        self.evidence.add_query(label, query, len(rows), len(subscriptions), time.monotonic() - started)
        return rows

    def queries(self, named_queries):
        """Run several Resource Graph queries concurrently.

        named_queries: {label: query} or {label: (query, partition)}. Returns {label: rows}.
        """
        results = {}
        with ThreadPoolExecutor(max_workers=max(1, self.args.graph_parallel)) as pool:
            futures = {}
            for label, spec in named_queries.items():
                query, partition = spec if isinstance(spec, tuple) else (spec, True)
                futures[pool.submit(self.query, query, partition, label)] = label
            for future in as_completed(futures):
                results[futures[future]] = future.result()
        return results

    # -- ARM ---------------------------------------------------------------

    def _bearer(self):
        with self._token_lock:
            if not self._token or self._token.expires_on - 300 < time.time():
                self._token = self.credential.get_token(ARM_SCOPE)
            return self._token.token

    def _session(self):
        session = getattr(self._local, "session", None)
        if session is None:
            session = requests.Session()
            adapter = requests.adapters.HTTPAdapter(pool_connections=4, pool_maxsize=4)
            session.mount("https://", adapter)
            self._local.session = session
        return session

    def arm_get(self, path, api_version, params=None):
        """GET an ARM resource. `path` is an ARM id/path or a full nextLink URL."""
        url = path if path.startswith("https://") else ARM_ENDPOINT + path
        query = dict(params or {})
        if api_version and "api-version=" not in url:
            query["api-version"] = api_version
        for attempt in range(1, 7):
            self.arm_limiter.wait()
            headers = {"Authorization": f"Bearer {self._bearer()}", "Accept": "application/json"}
            try:
                response = self._session().get(url, params=query, headers=headers, timeout=60)
            except requests.RequestException as error:
                if attempt < 6:
                    time.sleep(2 * attempt)
                    continue
                raise ArmError("", "ConnectionError", str(error), url) from error
            if response.status_code in (429, 500, 502, 503, 504) and attempt < 6:
                retry_after = response.headers.get("Retry-After")
                delay = int(retry_after) if retry_after and retry_after.isdigit() else min(60, 3 * attempt)
                LOG.debug("ARM %s, retry in %ss: %s", response.status_code, delay, url)
                time.sleep(delay)
                continue
            self.evidence.add_call("GET", url, query.get("api-version", ""), response.status_code)
            if response.status_code >= 400:
                try:
                    body = response.json().get("error", {})
                except ValueError:
                    body = {}
                raise ArmError(response.status_code, body.get("code", ""), body.get("message", response.reason), url)
            return response.json() if response.content else {}
        raise ArmError("", "RetryExhausted", "Retries exhausted", url)

    def arm_list(self, path, api_version, params=None, max_pages=200):
        """GET a collection and follow nextLink."""
        items = []
        body = self.arm_get(path, api_version, params)
        pages = 1
        while True:
            items.extend(body.get("value", []))
            next_link = body.get("nextLink")
            if not next_link or pages >= max_pages:
                break
            body = self.arm_get(next_link, None)
            pages += 1
        return items

    def try_arm(self, stage, scope, call, *args, **kwargs):
        """Run an ARM call, logging failures as gaps. Returns None on error."""
        try:
            return call(*args, **kwargs)
        except ArmError as error:
            self.errors.add(stage, scope, error.url.split("?")[0].replace(ARM_ENDPOINT, ""),
                            error.status, error.code, error.message)
            return None

    # -- concurrency -------------------------------------------------------

    def parallel(self, function, items, label="tasks", workers=None):
        """Apply function to items with a thread pool. Returns results in input order.

        Exceptions are logged to the error log and give None for that item.
        """
        items = list(items)
        if not items:
            return []
        workers = workers or self.args.workers
        results = [None] * len(items)
        done = 0
        started = time.monotonic()
        with ThreadPoolExecutor(max_workers=workers) as pool:
            futures = {pool.submit(function, item): index for index, item in enumerate(items)}
            for future in as_completed(futures):
                index = futures[future]
                try:
                    results[index] = future.result()
                except Exception as error:  # noqa: BLE001 - one item must not stop the run
                    self.errors.add(label, str(items[index])[:200], "parallel", code=type(error).__name__,
                                    message=str(error))
                done += 1
                if done % 200 == 0 or done == len(items):
                    LOG.info("%s: %d/%d (%.0fs)", label, done, len(items), time.monotonic() - started)
        return results


# ---------------------------------------------------------------------------
# Tags and classification
# ---------------------------------------------------------------------------

ENVIRONMENT_TAG_KEYS = ("environment", "env", "ambiente", "entorno", "environment ")
ENVIRONMENT_VALUES = {
    "prod": "PROD", "prd": "PROD", "produccion": "PROD", "producción": "PROD", "production": "PROD",
    "cert": "CERT", "certificacion": "CERT", "certificación": "CERT",
    "desa": "DESA", "desarrollo": "DESA", "dev": "DESA", "development": "DESA",
}
SUBSCRIPTION_ENVIRONMENT_TOKENS = {
    "PROD": "PROD", "PRODUCCION": "PROD", "PRD": "PROD",
    "CERT": "CERT", "CERTIFICACION": "CERT",
    "DESA": "DESA", "DESARROLLO": "DESA", "DEV": "DESA",
}
COD_APP_TAG_KEYS = ("codapp", "cod_app", "cod-app", "codigoapp", "codigo_app", "app_code", "appcode")


def tags_lower(tags):
    """Lowercase and trim tag keys. Values keep their case, trimmed."""
    if not isinstance(tags, dict):
        return {}
    return {str(key).strip().lower(): str(value if value is not None else "").strip() for key, value in tags.items()}


def get_tag(tags, *keys):
    normalized = tags_lower(tags)
    for key in keys:
        value = normalized.get(str(key).strip().lower())
        if value:
            return value
    return ""


def normalize_environment(tags, subscription_name=""):
    """Return (environment, evidence). Tag first, then a token in the subscription name."""
    raw = get_tag(tags, *ENVIRONMENT_TAG_KEYS)
    if raw:
        value = ENVIRONMENT_VALUES.get(raw.strip().lower())
        if value:
            return value, "TAG_ENVIRONMENT"
        return "UNKNOWN", f"TAG_ENVIRONMENT_NO_RECONOCIDO:{raw}"
    tokens = set(re.split(r"[^A-Z0-9]+", str(subscription_name or "").upper()))
    for token, value in SUBSCRIPTION_ENVIRONMENT_TOKENS.items():
        if token in tokens:
            return value, f"SUBSCRIPTION_NAME:{token}"
    return "UNKNOWN", "TAG_ENVIRONMENT_AUSENTE"


def normalize_cod_app(tags):
    return get_tag(tags, *COD_APP_TAG_KEYS).upper()


def normalize_operation(tags):
    """Normalize the operation tag variants. Returns (normalized, raw)."""
    raw = get_tag(tags, "operation", "operacion", "operación")
    compact = raw.lower().replace("_", "-").replace(" ", "-")
    while "--" in compact:
        compact = compact.replace("--", "-")
    if compact in {"ibm", "ibm-cloud", "ibmcloud"}:
        return "IBM", raw
    if compact.startswith("kyndryl"):
        return "KYNDRYL", raw
    if not raw:
        return "UNKNOWN", ""
    return raw.upper(), raw


def normalize_vendor(tags):
    raw = get_tag(tags, "vendor")
    if raw.lower() == "databricks":
        return "DATABRICKS"
    if get_tag(tags, "orca") == "true" or get_tag(tags, "orca-role"):
        return "ORCA"
    return raw.upper() if raw else "UNKNOWN"


MANAGED_BY_PLATFORMS = (
    ("microsoft.databricks/workspaces", "DATABRICKS_MANAGED"),
    ("microsoft.redhatopenshift/openshiftclusters", "ARO_OPENSHIFT"),
    ("microsoft.containerservice/managedclusters", "AKS_MANAGED"),
)
EXCLUDED_CLASSES = {"ORCA_SECURITY", "DATABRICKS_MANAGED", "ARO_OPENSHIFT", "AKS_MANAGED"}


def classify_infrastructure(tags, name, resource_group, rg_managed_by=""):
    """Classify platform-managed infrastructure so it does not distort backup coverage.

    Returns (infrastructure_class, backup_scope, evidence).
    backup_scope: EXCLUDED_PLATFORM | REVIEW_REQUIRED | BACKUP_REQUIRED
    """
    normalized = tags_lower(tags)
    name_lower = str(name or "").lower()
    rg_lower = str(resource_group or "").lower()
    managed_by = normalize(rg_managed_by)

    orca = []
    if normalized.get("orca") == "true":
        orca.append("TAG:Orca=True")
    if normalized.get("orca-role"):
        orca.append(f"TAG:orca-role={normalized['orca-role']}")
    if "orca-security" in rg_lower:
        orca.append("RESOURCE_GROUP:orca-security")
    if name_lower.startswith("orca."):
        orca.append("NAME:orca.*")
    if orca:
        return "ORCA_SECURITY", "EXCLUDED_PLATFORM", ";".join(orca)

    for provider_type, platform in MANAGED_BY_PLATFORMS:
        if f"/providers/{provider_type}/" in managed_by:
            return platform, "EXCLUDED_PLATFORM", f"RG_MANAGED_BY:{provider_type}"

    if normalized.get("vendor") == "databricks":
        return "DATABRICKS_MANAGED", "EXCLUDED_PLATFORM", "TAG:Vendor=Databricks"

    aro = []
    if rg_lower.startswith("aro-"):
        aro.append("RESOURCE_GROUP:aro-*")
    if any(key.startswith(("kubernetes.io_cluster.", "kubernetes.io/cluster/")) for key in normalized):
        aro.append("TAG:kubernetes.io_cluster.*")
    if aro and (rg_lower.startswith("aro-") or "-master-" in name_lower or "-worker-" in name_lower):
        return "ARO_OPENSHIFT", "EXCLUDED_PLATFORM", ";".join(aro)
    if rg_lower.startswith("mc_") and normalized.get("aks-managed-cluster-name"):
        return "AKS_MANAGED", "EXCLUDED_PLATFORM", "TAG:aks-managed-cluster-name"

    if managed_by:
        provider = "/".join(managed_by.split("/providers/")[-1].split("/")[:2])
        return "MANAGED_RG_OTHER", "REVIEW_REQUIRED", f"RG_MANAGED_BY:{provider}"

    return "TRADITIONAL", "BACKUP_REQUIRED", "DEFAULT_TRADITIONAL"


def load_managed_resource_groups(ctx):
    """Return {resource_group_id_lower: managedBy_lower} for resource groups owned by a platform."""
    rows = ctx.query(
        "resourcecontainers"
        " | where type =~ 'microsoft.resources/subscriptions/resourcegroups' and isnotempty(managedBy)"
        " | project id = tolower(id), managedBy = tolower(managedBy)",
        label="managed_resource_groups",
    )
    return {row["id"]: row["managedBy"] for row in rows}


def resource_group_id(subscription_id, resource_group):
    return f"/subscriptions/{subscription_id}/resourcegroups/{resource_group}".lower()


BASE_FIELDS = [
    "subscription", "subscription_id", "resource_group", "name", "id", "location",
    "environment", "environment_evidence", "cod_app", "operation_raw", "operation_normalized",
    "vendor_normalized", "infrastructure_class", "backup_scope", "infrastructure_evidence",
]


def base_columns(ctx, row, managed_rgs=None):
    """Common identification and classification columns for an ARG resource row.

    Expects row keys: id, name, resourceGroup, subscriptionId, location, tags.
    Raw tags are not written to the output; only the derived fields are.
    """
    tags = row.get("tags") or {}
    subscription_name = ctx.subscription_name(row.get("subscriptionId"))
    environment, evidence = normalize_environment(tags, subscription_name)
    operation, operation_raw = normalize_operation(tags)
    managed_by = (managed_rgs or {}).get(resource_group_id(row.get("subscriptionId"), row.get("resourceGroup")), "")
    infra_class, scope, infra_evidence = classify_infrastructure(
        tags, row.get("name"), row.get("resourceGroup"), managed_by)
    return {
        "subscription": subscription_name,
        "subscription_id": row.get("subscriptionId", ""),
        "resource_group": row.get("resourceGroup", ""),
        "name": row.get("name", ""),
        "id": row.get("id", ""),
        "location": row.get("location", ""),
        "environment": environment,
        "environment_evidence": evidence,
        "cod_app": normalize_cod_app(tags),
        "operation_raw": operation_raw,
        "operation_normalized": operation,
        "vendor_normalized": normalize_vendor(tags),
        "infrastructure_class": infra_class,
        "backup_scope": scope,
        "infrastructure_evidence": infra_evidence,
    }


# ---------------------------------------------------------------------------
# Azure Backup: vaults, policies and protected items
# ---------------------------------------------------------------------------

RSV_ITEM_TYPE = "microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems"
BV_INSTANCE_TYPE = "microsoft.dataprotection/backupvaults/backupinstances"
RSV_API = "2024-04-01"
BV_API = "2024-04-01"

ACTIVE_STATES = {"protected", "irpending", "protectionerror", "protectionconfigured", "configuringprotection"}
STOPPED_STATES = {"protectionstopped", "protectionpaused", "backupssuspended", "backupschedulessuspended",
                  "retentionschedulessuspended", "softdeleted"}
FAILED_BACKUP_STATUSES = {"failed", "completedwithwarnings"}


def _schedule_summary(schedule):
    if not isinstance(schedule, dict):
        return ""
    frequency = schedule.get("scheduleRunFrequency") or ""
    if schedule.get("scheduleFrequencyInMins"):
        return f"Every{schedule['scheduleFrequencyInMins']}min"
    hourly = schedule.get("hourlySchedule") or {}
    if str(frequency).lower() == "hourly" and hourly.get("interval"):
        return f"Hourly/{hourly['interval']}h"
    if str(frequency).lower() == "weekly":
        days = schedule.get("scheduleRunDays") or safe_get(schedule, "weeklySchedule", "scheduleRunDays") or []
        return "Weekly" + (f"({','.join(day[:3] for day in days)})" if days else "")
    return str(frequency)


def _retention_counts(retention):
    """Return daily/weekly/monthly/yearly counts and max days of a LongTerm or Simple retention policy."""
    result = {"daily": None, "weekly": None, "monthly": None, "yearly": None, "max_days": None}
    if not isinstance(retention, dict):
        return result
    if retention.get("retentionDuration"):
        days = retention_count_days(retention["retentionDuration"])
        result["daily"] = days
        result["max_days"] = days
        return result
    maximum = None
    for key, schedule in (("daily", "dailySchedule"), ("weekly", "weeklySchedule"),
                          ("monthly", "monthlySchedule"), ("yearly", "yearlySchedule")):
        duration = safe_get(retention, schedule, "retentionDuration")
        if duration:
            result[key] = duration.get("count")
            days = retention_count_days(duration)
            if days is not None:
                maximum = max(maximum or 0, days)
    result["max_days"] = maximum
    return result


POLICY_FIELDS = [
    "policy_name", "policy_type", "policy_workload", "schedule_frequency", "instant_restore_days",
    "daily_retention", "weekly_retention", "monthly_retention", "yearly_retention", "log_backup_frequency",
    "retention_max_days", "archive_tiering", "datastores", "retention_summary",
]


def summarize_rsv_policy(policy_id, name, properties):
    properties = properties or {}
    summary = {key: "" for key in POLICY_FIELDS}
    summary.update({
        "policy_id": policy_id,
        "policy_name": name,
        "policy_source": "RECOVERY_SERVICES_VAULT",
        "policy_backup_management_type": properties.get("backupManagementType", ""),
        "policy_workload": properties.get("workLoadType", ""),
        "policy_type": properties.get("policyType") or (
            "V1" if properties.get("backupManagementType") == "AzureIaasVM" else ""),
        "protected_items_count": properties.get("protectedItemsCount", ""),
    })
    sub_policies = properties.get("subProtectionPolicy") or []
    if sub_policies:
        parts = []
        maximum = None
        for sub in sub_policies:
            sub_type = sub.get("policyType", "")
            counts = _retention_counts(sub.get("retentionPolicy"))
            schedule = _schedule_summary(sub.get("schedulePolicy"))
            if sub_type.lower() == "log":
                summary["log_backup_frequency"] = schedule
            elif sub_type.lower() in ("full", "snapshotfull"):
                summary["schedule_frequency"] = schedule
                for key in ("daily", "weekly", "monthly", "yearly"):
                    summary[f"{key}_retention"] = counts[key] if counts[key] is not None else ""
            if counts["max_days"] is not None:
                maximum = max(maximum or 0, counts["max_days"])
            parts.append(f"{sub_type}:{schedule}:{counts['max_days'] or ''}d")
        summary["retention_summary"] = ";".join(parts)
        summary["retention_max_days"] = maximum if maximum is not None else ""
    else:
        counts = _retention_counts(properties.get("retentionPolicy"))
        vault_counts = _retention_counts(safe_get(properties, "vaultRetentionPolicy", "vaultRetention"))
        if vault_counts["max_days"] is not None:
            counts = vault_counts
            summary["datastores"] = "Snapshot+Vault"
        summary["schedule_frequency"] = _schedule_summary(properties.get("schedulePolicy"))
        summary["instant_restore_days"] = first_value(
            properties.get("instantRpRetentionRangeInDays"),
            safe_get(properties, "vaultRetentionPolicy", "snapshotRetentionInDays")) or ""
        for key in ("daily", "weekly", "monthly", "yearly"):
            summary[f"{key}_retention"] = counts[key] if counts[key] is not None else ""
        summary["retention_max_days"] = counts["max_days"] if counts["max_days"] is not None else ""
    tiering = properties.get("tieringPolicy") or {}
    archive = safe_get(tiering, "ArchivedRP", "tieringMode")
    summary["archive_tiering"] = archive if archive and archive != "DoNotTier" else ""
    return summary


def summarize_bv_policy(policy_id, name, properties):
    properties = properties or {}
    summary = {key: "" for key in POLICY_FIELDS}
    summary.update({
        "policy_id": policy_id,
        "policy_name": name,
        "policy_source": "BACKUP_VAULT",
        "policy_backup_management_type": "DataProtection",
        "policy_workload": ";".join(properties.get("datasourceTypes") or []),
        "policy_type": "BackupVault",
        "protected_items_count": "",
    })
    schedules, stores, retention_parts = [], set(), []
    maximum = None
    for rule in properties.get("policyRules") or []:
        object_type = rule.get("objectType")
        if object_type == "AzureBackupRule":
            intervals = safe_get(rule, "trigger", "schedule", "repeatingTimeIntervals") or []
            schedules.extend(interval.split("/")[-1] for interval in intervals)
            store = safe_get(rule, "dataStore", "dataStoreType")
            if store:
                stores.add(store)
        elif object_type == "AzureRetentionRule":
            for lifecycle in rule.get("lifecycles") or []:
                duration = safe_get(lifecycle, "deleteAfter", "duration")
                store = safe_get(lifecycle, "sourceDataStore", "dataStoreType") or ""
                if store:
                    stores.add(store)
                days = iso_duration_days(duration)
                if days is not None:
                    maximum = max(maximum or 0, days)
                retention_parts.append(f"{store}:{rule.get('name', '')}={duration}")
    summary["schedule_frequency"] = ";".join(sorted(set(schedules)))
    summary["datastores"] = ";".join(sorted(stores))
    summary["retention_summary"] = ";".join(retention_parts)
    summary["retention_max_days"] = maximum if maximum is not None else ""
    return summary


PROTECTION_FIELDS = [
    "backup_enabled", "backup_active_count", "backup_stopped", "retained_restore_point",
    "protection_source", "vault_name", "vault_resource_group", "vault_subscription_id",
    "vault_storage_redundancy", "vault_cross_region_restore", "vault_soft_delete", "vault_immutability",
    "protected_item_name", "workload_type", "protection_state", "health_status", "last_backup_status",
    "last_backup_time", "last_backup_age_hours", "last_recovery_point", "oldest_recovery_point",
    "recovery_point_count", "configured_max_retention_days", "configured_rp_frequency_hours",
    "archive_enabled", "policy_resolved",
] + POLICY_FIELDS


class BackupIndex:
    """Azure Backup items from Recovery Services vaults and Backup vaults, keyed by source resource."""

    def __init__(self):
        self.vaults = {}
        self.policies = {}
        self.items = []
        self.by_resource = defaultdict(list)
        self.by_parent = defaultdict(list)

    def add_item(self, item):
        self.items.append(item)
        source = item["source_resource_id"]
        if source:
            self.by_resource[source].append(item)
            for parent in parent_resource_ids(source):
                self.by_parent[parent].append(item)

    def for_resource(self, resource_id, include_children=False):
        key = normalize(resource_id)
        found = list(self.by_resource.get(key, []))
        if include_children:
            found.extend(self.by_parent.get(key, []))
        return found


def load_backup_index(ctx):
    """Load vaults, policies, RSV protected items and Backup vault instances from Resource Graph.

    Policies referenced by an item but missing from Resource Graph are fetched with ARM GET.
    """
    index = BackupIndex()
    results = ctx.queries({
        "vaults": (
            "resources"
            " | where type in~ ('microsoft.recoveryservices/vaults', 'microsoft.dataprotection/backupvaults')"
            " | project id, name, type, resourceGroup, subscriptionId, location, sku = tostring(sku.name),"
            " rsvRedundancy = tostring(properties.redundancySettings.standardTierStorageRedundancy),"
            " rsvCrr = tostring(properties.redundancySettings.crossRegionRestore),"
            " rsvSoftDelete = tostring(properties.securitySettings.softDeleteSettings.softDeleteState),"
            " rsvSoftDeleteDays = tostring(properties.securitySettings.softDeleteSettings.softDeleteRetentionPeriodInDays),"
            " rsvImmutability = tostring(properties.securitySettings.immutabilitySettings.state),"
            " bvStorage = properties.storageSettings,"
            " bvCrr = tostring(properties.featureSettings.crossRegionRestoreSettings.state),"
            " bvSoftDelete = tostring(properties.securitySettings.softDeleteSettings.state),"
            " bvSoftDeleteDays = tostring(properties.securitySettings.softDeleteSettings.retentionDurationInDays),"
            " bvImmutability = tostring(properties.securitySettings.immutabilitySettings.state)"
        ),
        "policies": (
            "recoveryservicesresources"
            " | where type in~ ('microsoft.recoveryservices/vaults/backuppolicies',"
            " 'microsoft.dataprotection/backupvaults/backuppolicies')"
            " | project id, name, type, properties"
        ),
        "rsv_items": f"recoveryservicesresources | where type =~ '{RSV_ITEM_TYPE}' | project id, name, properties",
        "bv_instances": f"recoveryservicesresources | where type =~ '{BV_INSTANCE_TYPE}' | project id, name, properties",
    })

    for row in results["vaults"]:
        is_rsv = row["type"].lower() == "microsoft.recoveryservices/vaults"
        storage = row.get("bvStorage") or []
        bv_redundancy = ";".join(sorted({entry.get("type", "") for entry in storage if isinstance(entry, dict)}))
        index.vaults[normalize(row["id"])] = {
            "id": row["id"],
            "name": row["name"],
            "kind": "RECOVERY_SERVICES_VAULT" if is_rsv else "BACKUP_VAULT",
            "resource_group": row["resourceGroup"],
            "subscription_id": row["subscriptionId"],
            "location": row["location"],
            "sku": row.get("sku", ""),
            "storage_redundancy": row.get("rsvRedundancy") if is_rsv else bv_redundancy,
            "cross_region_restore": row.get("rsvCrr") if is_rsv else row.get("bvCrr"),
            "soft_delete": row.get("rsvSoftDelete") if is_rsv else row.get("bvSoftDelete"),
            "soft_delete_days": row.get("rsvSoftDeleteDays") if is_rsv else row.get("bvSoftDeleteDays"),
            "immutability": row.get("rsvImmutability") if is_rsv else row.get("bvImmutability"),
        }

    for row in results["policies"]:
        policy_id = normalize(row["id"])
        if row["type"].lower().startswith("microsoft.recoveryservices"):
            index.policies[policy_id] = summarize_rsv_policy(row["id"], row["name"], row.get("properties"))
        else:
            index.policies[policy_id] = summarize_bv_policy(row["id"], row["name"], row.get("properties"))

    for row in results["rsv_items"]:
        index.add_item(_rsv_item(row, index))
    for row in results["bv_instances"]:
        index.add_item(_bv_item(row, index))

    _resolve_missing_policies(ctx, index)
    LOG.info("Backup index: %d vaults, %d policies, %d items",
             len(index.vaults), len(index.policies), len(index.items))
    return index


def _vault_id_from(item_id, marker):
    lowered = normalize(item_id)
    position = lowered.find(marker)
    if position < 0:
        return ""
    end = lowered.find("/", position + len(marker))
    return lowered[:end] if end > 0 else lowered


def _rsv_item(row, index):
    properties = row.get("properties") or {}
    extended = properties.get("extendedInfo") or {}
    exclusion = safe_get(properties, "extendedProperties", "diskExclusionProperties") or {}
    vault_id = normalize(properties.get("vaultId")) or _vault_id_from(row["id"], "/vaults/")
    state = properties.get("protectionState") or properties.get("currentProtectionState") or ""
    return {
        "item_id": row["id"],
        "protection_source": "RECOVERY_SERVICES_VAULT",
        "source_resource_id": normalize(first_value(
            properties.get("sourceResourceId"), safe_get(properties, "dataSourceInfo", "resourceID"))),
        "protected_item_name": properties.get("friendlyName", row.get("name", "")),
        "workload_type": properties.get("workloadType", ""),
        "backup_management_type": properties.get("backupManagementType", ""),
        "datasource_type": safe_get(properties, "dataSourceInfo", "datasourceType", default=""),
        "server_name": properties.get("serverName", ""),
        "parent_name": properties.get("parentName", ""),
        "protection_state": state,
        "health_status": properties.get("healthStatus") or properties.get("protectionStatus") or "",
        "last_backup_status": properties.get("lastBackupStatus", ""),
        "last_backup_time": iso(properties.get("lastBackupTime")),
        "last_recovery_point": iso(properties.get("lastRecoveryPoint")),
        "oldest_recovery_point": iso(extended.get("oldestRecoveryPoint")),
        "recovery_point_count": extended.get("recoveryPointCount", ""),
        "configured_max_retention_days": timespan_days(properties.get("configuredMaximumRetention")) or "",
        "configured_rp_frequency_hours": _hours(timespan_days(properties.get("configuredRPGenerationFrequency"))),
        "archive_enabled": yes_no(properties.get("isArchiveEnabled")),
        "policy_id": normalize(first_value(properties.get("policyId"), safe_get(properties, "policyInfo", "policyId"))),
        "policy_name_raw": first_value(properties.get("policyName"), safe_get(properties, "policyInfo", "name")) or "",
        "vault_id": vault_id,
        "disk_lun_list": ";".join(str(lun) for lun in exclusion.get("diskLunList") or []),
        "disk_is_inclusion_list": exclusion.get("isInclusionList"),
        "protected_items_extra": "",
    }


def _bv_item(row, index):
    properties = row.get("properties") or {}
    extended = properties.get("backupInstanceExtendedProperties") or {}
    points = extended.get("recoveryPointInfo") or {}
    policy_info = safe_get(extended, "policyExtendedInfo") or {}
    vault_id = normalize(extended.get("vaultId")) or _vault_id_from(row["id"], "/backupvaults/")
    params = safe_get(properties, "policyInfo", "policyParameters", "backupDatasourceParametersList") or []
    containers = []
    for entry in params:
        containers.extend(entry.get("containersList") or [])
    status = safe_get(properties, "protectionStatus", "status", default="")
    return {
        "item_id": row["id"],
        "protection_source": "BACKUP_VAULT",
        "source_resource_id": normalize(safe_get(properties, "dataSourceInfo", "resourceID")),
        "protected_item_name": properties.get("friendlyName", row.get("name", "")),
        "workload_type": safe_get(properties, "dataSourceInfo", "datasourceType", default=""),
        "backup_management_type": "DataProtection",
        "datasource_type": safe_get(properties, "dataSourceInfo", "datasourceType", default=""),
        "server_name": "",
        "parent_name": "",
        "protection_state": properties.get("currentProtectionState", ""),
        "health_status": status,
        "last_backup_status": "",
        "last_backup_time": iso(points.get("latestRecoveryPointInPrimaryRegion")),
        "last_recovery_point": iso(points.get("latestRecoveryPointInPrimaryRegion")),
        "oldest_recovery_point": iso(points.get("oldestRecoveryPointInPrimaryRegion")),
        "recovery_point_count": "",
        "configured_max_retention_days": timespan_days(policy_info.get("configuredMaximumRetention")) or "",
        "configured_rp_frequency_hours": _hours(timespan_days(policy_info.get("configuredRPGenerationFrequency"))),
        "archive_enabled": "",
        "policy_id": normalize(safe_get(properties, "policyInfo", "policyId")),
        "policy_name_raw": "",
        "vault_id": vault_id,
        "disk_lun_list": "",
        "disk_is_inclusion_list": None,
        "protected_items_extra": ";".join(containers),
    }


def _hours(days):
    return round(days * 24, 2) if days is not None else ""


def _resolve_missing_policies(ctx, index):
    missing = sorted({item["policy_id"] for item in index.items
                      if item["policy_id"] and item["policy_id"] not in index.policies})
    if not missing or ctx.args.skip_arm:
        return

    def fetch(policy_id):
        is_rsv = "/microsoft.recoveryservices/" in policy_id
        body = ctx.try_arm("policies", policy_id, ctx.arm_get, policy_id, RSV_API if is_rsv else BV_API)
        if not body:
            return None
        summarize = summarize_rsv_policy if is_rsv else summarize_bv_policy
        return policy_id, summarize(body.get("id", policy_id), body.get("name", ""), body.get("properties"))

    for result in ctx.parallel(fetch, missing, label="missing_policies"):
        if result:
            index.policies[result[0]] = result[1]


def item_is_active(item):
    return normalize(item.get("protection_state")) in ACTIVE_STATES


def item_is_stopped(item):
    return normalize(item.get("protection_state")) in STOPPED_STATES


def item_has_points(item):
    return bool(item.get("last_recovery_point") or item.get("oldest_recovery_point")
                or str(item.get("recovery_point_count") or "0") not in ("", "0"))


def item_is_healthy(item, stale_hours=48):
    """Active item whose last backup did not fail and whose newest point is recent enough."""
    if not item_is_active(item):
        return False
    if normalize(item.get("last_backup_status")) in FAILED_BACKUP_STATUSES:
        return False
    if normalize(item.get("health_status")) in {"actionrequired", "protectionerror", "unhealthy"}:
        return False
    frequency = item.get("configured_rp_frequency_hours")
    threshold = max(stale_hours, 2 * float(frequency)) if frequency not in ("", None) else stale_hours
    age = age_hours(item.get("last_recovery_point") or item.get("last_backup_time"))
    return age is not None and age <= threshold


def protection_columns(index, items):
    """Flatten the backup items of one resource into PROTECTION_FIELDS.

    With several items the active one with the newest recovery point is shown, and the
    counters reflect all of them.
    """
    columns = {key: "" for key in PROTECTION_FIELDS}
    active = [item for item in items if item_is_active(item)]
    stopped = [item for item in items if item_is_stopped(item)]
    columns["backup_enabled"] = "YES" if active else "NO"
    columns["backup_active_count"] = len(active)
    columns["backup_stopped"] = "YES" if stopped and not active else "NO"
    columns["retained_restore_point"] = "YES" if any(item_has_points(item) for item in stopped) else "NO"
    if not items:
        columns["policy_resolved"] = "NOT_APPLICABLE"
        return columns
    pool = active or stopped or items
    item = max(pool, key=lambda value: value.get("last_recovery_point") or "")
    vault = index.vaults.get(item["vault_id"], {})
    columns.update({
        "protection_source": ";".join(sorted({value["protection_source"] for value in items})),
        "vault_name": vault.get("name") or item["vault_id"].split("/")[-1],
        "vault_resource_group": vault.get("resource_group", ""),
        "vault_subscription_id": vault.get("subscription_id", ""),
        "vault_storage_redundancy": vault.get("storage_redundancy", ""),
        "vault_cross_region_restore": vault.get("cross_region_restore", ""),
        "vault_soft_delete": vault.get("soft_delete", ""),
        "vault_immutability": vault.get("immutability", ""),
        "protected_item_name": item["protected_item_name"],
        "workload_type": item["workload_type"],
        "protection_state": item["protection_state"],
        "health_status": item["health_status"],
        "last_backup_status": item["last_backup_status"],
        "last_backup_time": item["last_backup_time"],
        "last_backup_age_hours": age_hours(item["last_recovery_point"] or item["last_backup_time"]) or "",
        "last_recovery_point": item["last_recovery_point"],
        "oldest_recovery_point": item["oldest_recovery_point"],
        "recovery_point_count": item["recovery_point_count"],
        "configured_max_retention_days": item["configured_max_retention_days"],
        "configured_rp_frequency_hours": item["configured_rp_frequency_hours"],
        "archive_enabled": item["archive_enabled"],
    })
    policy = index.policies.get(item["policy_id"])
    if policy:
        columns["policy_resolved"] = "YES"
        for key in POLICY_FIELDS:
            columns[key] = policy.get(key, "")
    else:
        columns["policy_resolved"] = "NO" if item["policy_id"] else "NO_POLICY_ID"
        columns["policy_name"] = item.get("policy_name_raw", "")
    return columns


ORPHAN_FIELDS = [
    "orphan_reason", "protection_source", "vault_name", "vault_resource_group", "vault_subscription_id",
    "vault_storage_redundancy", "protected_item_name", "workload_type", "source_resource_id",
    "protection_state", "last_backup_status", "last_recovery_point", "days_since_last_recovery_point",
    "oldest_recovery_point", "recovery_point_count", "configured_max_retention_days", "policy_name",
    "item_id",
]


def orphan_item_rows(index, existing_resource_ids, item_filter):
    """Backup items whose source resource no longer exists, or that keep points without new backups.

    These keep consuming vault storage and protected-instance fees: a direct cost opportunity.
    existing_resource_ids: set of lowercase ids of live resources for the workloads checked.
    item_filter: callable(item) -> bool selecting the workloads this script owns.
    """
    rows = []
    for item in index.items:
        if not item_filter(item):
            continue
        source = item["source_resource_id"]
        if source and source not in existing_resource_ids:
            reason = "SOURCE_RESOURCE_DELETED"
        elif item_is_active(item) and (age_hours(item["last_recovery_point"] or item["last_backup_time"]) or 0) > 24 * 30:
            reason = "ACTIVE_WITHOUT_BACKUP_30D"
        else:
            continue
        vault = index.vaults.get(item["vault_id"], {})
        policy = index.policies.get(item["policy_id"], {})
        rows.append({
            "orphan_reason": reason,
            "protection_source": item["protection_source"],
            "vault_name": vault.get("name") or item["vault_id"].split("/")[-1],
            "vault_resource_group": vault.get("resource_group", ""),
            "vault_subscription_id": vault.get("subscription_id", ""),
            "vault_storage_redundancy": vault.get("storage_redundancy", ""),
            "protected_item_name": item["protected_item_name"],
            "workload_type": item["workload_type"],
            "source_resource_id": source,
            "protection_state": item["protection_state"],
            "last_backup_status": item["last_backup_status"],
            "last_recovery_point": item["last_recovery_point"],
            "days_since_last_recovery_point": days_since(item["last_recovery_point"] or item["last_backup_time"]),
            "oldest_recovery_point": item["oldest_recovery_point"],
            "recovery_point_count": item["recovery_point_count"],
            "configured_max_retention_days": item["configured_max_retention_days"],
            "policy_name": policy.get("policy_name") or item.get("policy_name_raw", ""),
            "item_id": item["item_id"],
        })
    return rows


def vault_rows(index):
    """Rows for *_backup_vaults.csv: one per vault with item counts, for cost analysis."""
    counts = defaultdict(lambda: defaultdict(int))
    for item in index.items:
        counts[item["vault_id"]]["items"] += 1
        counts[item["vault_id"]]["active" if item_is_active(item) else "inactive"] += 1
        counts[item["vault_id"]][f"wl:{item['workload_type']}"] += 1
    rows = []
    for vault_id, vault in sorted(index.vaults.items(), key=lambda pair: pair[1]["name"].lower()):
        stats = counts.get(vault_id, {})
        workloads = ";".join(f"{key[3:]}={value}" for key, value in sorted(stats.items()) if key.startswith("wl:"))
        rows.append({
            "vault_kind": vault["kind"], "vault_name": vault["name"], "vault_id": vault["id"],
            "subscription_id": vault["subscription_id"], "resource_group": vault["resource_group"],
            "location": vault["location"], "sku": vault["sku"],
            "storage_redundancy": vault["storage_redundancy"],
            "cross_region_restore": vault["cross_region_restore"], "soft_delete": vault["soft_delete"],
            "soft_delete_days": vault["soft_delete_days"], "immutability": vault["immutability"],
            "item_count": stats.get("items", 0), "active_item_count": stats.get("active", 0),
            "inactive_item_count": stats.get("inactive", 0), "workloads": workloads,
        })
    return rows


def policy_rows(index, workloads=None):
    """Rows for *_backup_policies.csv, optionally filtered by workload/datasource substrings."""
    rows = []
    for policy in index.policies.values():
        workload = str(policy.get("policy_workload", ""))
        if workloads and not any(token.lower() in workload.lower() for token in workloads):
            continue
        row = dict(policy)
        row["vault_name"] = _vault_id_from(policy["policy_id"], "/vaults/").split("/")[-1] or \
            _vault_id_from(policy["policy_id"], "/backupvaults/").split("/")[-1]
        rows.append(row)
    rows.sort(key=lambda value: (value.get("policy_workload", ""), value.get("vault_name", ""), value.get("policy_name", "")))
    return rows


# ---------------------------------------------------------------------------
# Resilience (Business Continuity Center)
# ---------------------------------------------------------------------------

RESILIENCE_FIELDS = ["resilience_solutions", "resilience_primary_region", "resilience_secondary_region",
                     "resilience_recovery_regions"]


def load_resilience(ctx, datasource_types):
    """Return {resource_id_lower: columns} from Business Continuity Center unified items.

    These rows list zone redundancy, geo-replication, ASR and Azure Backup per resource,
    only for resources that have at least one of them.
    """
    types = ", ".join(f"'{value}'" for value in datasource_types)
    rows = ctx.query(
        "recoveryservicesresources"
        " | where type =~ 'microsoft.azurebusinesscontinuity/unifiedprotecteditems'"
        f" | where tostring(properties.dataSourceInformation.dataSourceType) in~ ({types})"
        " | project id, resourceId = tolower(tostring(properties.linkedResourceInformation.linkedResourceId)),"
        " solutions = properties.distinctSolutionNames,"
        " primary = tostring(properties.protectionStatusInPrimaryRegion),"
        " secondary = tostring(properties.protectionStatusInSecondaryRegion),"
        " recoveryRegions = properties.protectedItems[0].recoveryRegions",
        label="resilience",
    )
    result = {}
    for row in rows:
        if not row.get("resourceId"):
            continue
        result[row["resourceId"]] = {
            "resilience_solutions": ";".join(row.get("solutions") or []),
            "resilience_primary_region": row.get("primary", ""),
            "resilience_secondary_region": row.get("secondary", ""),
            "resilience_recovery_regions": ";".join(row.get("recoveryRegions") or []),
        }
    return result


def resilience_columns(resilience, resource_id):
    return resilience.get(normalize(resource_id)) or {key: "" for key in RESILIENCE_FIELDS}


# ---------------------------------------------------------------------------
# Findings
# ---------------------------------------------------------------------------

SEVERITY_ORDER = {"HIGH": 0, "MEDIUM": 1, "LOW": 2, "INFO": 3}


class Findings:
    """Collects (code, severity) pairs for one row and exposes the primary one."""

    def __init__(self):
        self.items = []

    def add(self, code, severity):
        self.items.append((code, severity))

    def columns(self, default="COMPLIANT_OR_NO_ACTION"):
        if not self.items:
            return {"finding": default, "severity": "INFO", "all_findings": default}
        ordered = sorted(self.items, key=lambda pair: SEVERITY_ORDER.get(pair[1], 9))
        return {
            "finding": ordered[0][0],
            "severity": ordered[0][1],
            "all_findings": ";".join(code for code, _ in ordered),
        }


FINDING_FIELDS = ["finding", "severity", "all_findings"]


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------


class Output:
    """Writes CSV files (UTF-8 with BOM so Excel keeps accents), a JSON summary and an optional XLSX."""

    def __init__(self, directory, delimiter=","):
        self.directory = directory
        self.delimiter = delimiter
        self.files = []
        os.makedirs(directory, exist_ok=True)

    def csv(self, name, rows, fields=None):
        fields = list(fields or [])
        for row in rows:
            for key in row:
                if key not in fields:
                    fields.append(key)
        path = os.path.join(self.directory, name)
        with open(path, "w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.DictWriter(handle, fieldnames=fields, delimiter=self.delimiter, extrasaction="ignore")
            writer.writeheader()
            for row in rows:
                writer.writerow({key: _cell(row.get(key)) for key in fields})
        self.files.append((name, fields, rows))
        LOG.info("Wrote %-48s %6d rows", name, len(rows))
        return path

    def json(self, name, payload):
        path = os.path.join(self.directory, name)
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, ensure_ascii=False, default=str)
        LOG.info("Wrote %s", name)
        return path

    def xlsx(self, name):
        try:
            from openpyxl import Workbook
        except ImportError:
            LOG.warning("openpyxl is not installed; skipping %s", name)
            return None
        workbook = Workbook()
        workbook.remove(workbook.active)
        for file_name, fields, rows in self.files:
            sheet = workbook.create_sheet(_sheet_name(file_name))
            sheet.append(fields)
            for row in rows:
                sheet.append([_cell(row.get(key)) for key in fields])
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
        path = os.path.join(self.directory, name)
        workbook.save(path)
        LOG.info("Wrote %s", name)
        return path


def _cell(value):
    if value is None:
        return ""
    if isinstance(value, (list, tuple, set)):
        return ";".join(str(item) for item in value)
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    return value


def _sheet_name(file_name):
    name = os.path.splitext(file_name)[0]
    return name[-31:]


def count_by(rows, *keys):
    """Nested counts for the JSON summary: count_by(rows, "environment", "finding")."""
    if not keys:
        return len(rows)
    groups = defaultdict(list)
    for row in rows:
        groups[str(row.get(keys[0]) or "")].append(row)
    return {group: count_by(items, *keys[1:]) for group, items in sorted(groups.items())}


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser(description, default_output_dir):
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("--output-dir", default=default_output_dir, help="Output directory")
    parser.add_argument("--subscriptions", default="",
                        help="Comma-separated subscription ids or names. Default: every visible subscription")
    parser.add_argument("--workers", type=int, default=16, help="Concurrent ARM calls")
    parser.add_argument("--graph-parallel", type=int, default=4, help="Concurrent Resource Graph queries")
    parser.add_argument("--arm-rps", type=float, default=20.0, help="Max ARM requests per second")
    parser.add_argument("--skip-arm", action="store_true",
                        help="Resource Graph only; skip per-resource ARM detail calls")
    parser.add_argument("--stale-hours", type=int, default=48,
                        help="A backup older than this (or 2x its frequency) counts as unhealthy")
    parser.add_argument("--xlsx", action="store_true", help="Also write one workbook with every CSV as a sheet")
    parser.add_argument("--delimiter", default=",", help="CSV delimiter (use ';' for Excel in Spanish locale)")
    parser.add_argument("--verbose", action="store_true")
    return parser


def setup_logging(verbose=False):
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)-7s %(message)s",
        datefmt="%H:%M:%S",
        stream=sys.stderr,
    )
    for noisy in ("azure", "urllib3", "msal"):
        logging.getLogger(noisy).setLevel(logging.WARNING)


def run_metadata(ctx, extra=None):
    """Metadata block for *_summary.json: scope, timings and whether the result is partial."""
    errors = ctx.errors.count_by_status()
    payload = {
        "generated_at": iso(utc_now()),
        "duration_seconds": round((utc_now() - ctx.started).total_seconds(), 1),
        "subscriptions_in_scope": len(ctx.subscriptions),
        "subscription_filter": ctx.args.subscriptions or "ALL_VISIBLE",
        "skip_arm": ctx.args.skip_arm,
        "errors_by_status": errors,
        "partial": bool(ctx.errors.rows),
    }
    payload.update(extra or {})
    payload["evidence"] = ctx.evidence.as_dict()
    return payload


def timestamp_dir(base):
    return os.path.join(base, utc_now().strftime("%Y%m%d-%H%M%S"))


def days_since(value):
    hours = age_hours(value)
    return round(hours / 24, 1) if hours is not None else ""


def within_days(value, days):
    parsed = parse_datetime(value)
    return bool(parsed and parsed >= utc_now() - timedelta(days=days))
