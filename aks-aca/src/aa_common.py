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
    result = []
    for row in rows:
        copy = {}
        for key, value in row.items():
            copy[key] = value
            if key == after:
                copy["cod_app"] = code_of(row)
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
