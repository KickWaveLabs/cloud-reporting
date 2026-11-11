"""Azure client for the aks-aca scripts: read-only and self-contained, so this folder runs on its own.

Covers: Azure Resource Graph paging with concurrency, read-only ARM GET calls with retries and rate
limiting, the Cost Management query API, the error log of access gaps, evidence of every query,
environment and codApp tags, and CSV/JSON/XLSX output.

Started as a copy of the generic part of backup/src/bu_common.py; the two are maintained apart.
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
from datetime import datetime, timezone

import requests
from azure.core.exceptions import HttpResponseError
from azure.identity import DefaultAzureCredential
from azure.mgmt.resourcegraph import ResourceGraphClient
from azure.mgmt.resourcegraph.models import QueryRequest, QueryRequestOptions, ResultFormat

ARM_ENDPOINT = "https://management.azure.com"
ARM_SCOPE = "https://management.azure.com/.default"
COST_API = "2023-11-01"

LOG = logging.getLogger("aa")

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


def parse_datetime(value):
    """Parse ARM timestamps. Returns None for empty or .NET minimum values.

    Azure Backup reports lastBackupTime 2001-01-01T00:00:00 for items never backed up.
    """
    if not value or str(value).startswith(("0001-01-01", "2001-01-01T00:00:00")):
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
                    payload = response.json()
                except ValueError:
                    payload = {}
                # ARM nests the error under "error"; Azure Monitor may return it flat or as a plain string.
                if isinstance(payload, dict):
                    body = payload.get("error", payload)
                else:
                    body = {"message": str(payload)[:500]}
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


NAME_SUFFIX_ENVIRONMENTS = {"p": "PROD", "c": "CERT", "d": "DESA"}


def environment_from_name(name):
    """Environment from the naming convention: P, C or D plus digits at the end (msqleu2appxp01 -> PROD).

    Fallback for resources without an environment tag, such as servers that no longer exist.
    """
    match = re.search(r"([pcd])\d+$", str(name or "").lower())
    return NAME_SUFFIX_ENVIRONMENTS[match.group(1)] if match else ""


# ---------------------------------------------------------------------------
# Cost Management
# ---------------------------------------------------------------------------


def cost_query(ctx, subscription_id, body):
    """POST a Cost Management query and follow nextLink. Returns (columns, rows)."""
    url = f"{ARM_ENDPOINT}/subscriptions/{subscription_id}/providers/Microsoft.CostManagement/query"
    columns, rows = [], []
    params = {"api-version": COST_API}
    grouping = "+".join(group["name"] for group in body["dataset"].get("grouping", []))
    ctx.evidence.add_body(f"Cost Management query by {grouping}", url, COST_API, body)
    while url:
        for attempt in range(1, 8):
            headers = {"Authorization": f"Bearer {ctx._bearer()}", "Content-Type": "application/json"}
            try:
                response = ctx._session().post(url, params=params, json=body, headers=headers, timeout=180)
            except requests.RequestException as error:
                if attempt == 7:
                    raise ArmError("", "ConnectionError", str(error), url) from error
                time.sleep(10 * attempt)
                continue
            ctx.evidence.add_call("POST", url, COST_API, response.status_code)
            if response.status_code == 429 or response.status_code >= 500:
                retry = next((value for key, value in response.headers.items()
                              if key.lower().endswith("retry-after") and value.isdigit()), None)
                time.sleep(int(retry) if retry else 10 * attempt)
                continue
            break
        if response.status_code >= 400:
            try:
                payload = response.json() if response.content else {}
            except ValueError:
                payload = {}
            error = payload.get("error", payload) if isinstance(payload, dict) else {"message": str(payload)[:500]}
            raise ArmError(response.status_code, error.get("code", ""), error.get("message", ""), url)
        properties = response.json().get("properties", {})
        columns = [column["name"] for column in properties.get("columns", [])]
        rows.extend(properties.get("rows", []))
        url = properties.get("nextLink")
        params = None
    return columns, rows


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

