# AKS vs Azure Container Apps

Read-only comparison of AKS and Azure Container Apps (ACA) for applications with fewer than 15
microservices: monthly cost and break-even point, application-level parity, service mesh, and
configuration drift of the current AKS clusters. Output: an Excel workbook with editable assumptions
and one slide deck built with Copilot.

## Layout

| Path | Content | In git |
|---|---|---|
| `src/` | Python scripts (`aa_*.py`) and `requirements.txt` | Yes |
| `docs/` | Cost analysis views, application-level comparison, deck outline and Copilot prompts | Yes |
| `scripts/` | Queries to check figures by hand: `container-insights.kql` (cluster > Logs), `resource-graph.kql` (Resource Graph Explorer), `grafana.promql` (Grafana > Explore), `deployment-prefixes.pq` (Excel Power Query on the deployment export) | Yes |
| `results/aa-<source>/` | CSV and JSON written by each script | No |
| `reports/<date>-aks-aca/` | `aks_aca_comparison.xlsx` and a README with that run's scope | No |

The folder is self-contained: it does not need `backup/` or any other topic. `aa_azure.py` is the Azure
client (Resource Graph paging, ARM GET with retries, Cost Management query, error log, query evidence);
`aa_common.py` adds the metrics call, the office-hours calendar and the paths of this topic.

Guides:

- [docs/cost-screenshots.md](docs/cost-screenshots.md): Cost analysis views that show the AKS cost.
- [docs/app-level-comparison.md](docs/app-level-comparison.md): what changes for an application on
  each platform (build, endpoints, data, producers and consumers, configuration, release, GitOps),
  then the service mesh.
- [docs/deck-and-prompts.md](docs/deck-and-prompts.md): slide order, source of each slide and the
  Copilot prompts.
- [docs/data-guide.md](docs/data-guide.md): what each CSV holds, how to rebuild every deck figure with
  pivot tables, and metrics not collected yet.
- [docs/adt-019-runtime.md](docs/adt-019-runtime.md): the bank's runtime decision tree (ADT_019: AKS,
  Container Apps, App Service, Functions or VM), summarised.
- [docs/fat-001-monolith-vs-distributed.md](docs/fat-001-monolith-vs-distributed.md): the bank's
  monolith vs distributed decision tree (FAT_001), summarised.
- [docs/microservice-sizing-criteria.md](docs/microservice-sizing-criteria.md): the bank's seven
  criteria to size a microservice, summarised.

## Sources

| Need | Source | Script | Data type |
|---|---|---|---|
| Clusters, node pools, add-ons, ACA apps, App Service plans and sites, APIM APIs | Resource Graph; ARM GET `{apim}/apis` | `aa_inventory.py` | Measured |
| Node CPU and memory, requests per API | Azure Monitor platform metrics (GET) | `aa_metrics.py` | Measured |
| Workloads, replicas, CPU/memory per replica | Container Insights (Log Analytics, GET) | `aa_container_insights.py` | Measured |
| Actual cost of clusters, ACA environments and PROD App Service / Functions plans | Cost Management query API, `AmortizedCost` | `aa_cost.py` | Measured |
| List prices | Azure Retail Prices API (public, no sign-in) | `aa_prices.py` | Vendor data |
| Unschedulable pods, OOMKilled containers, APIM backend latency | Azure Monitor platform metrics; Container Insights (GET) | `aa_signals.py` | Measured |

What is never read: tags (only used to derive the environment), identities, APIM subscription keys and
named values, Container Apps secrets and environment variables, container logs, pod labels.

## ADT_019 rules used in the model

- ACA in Consumption (serverless) mode; dedicated workload profiles need architecture approval.
- 1 replica Monday to Friday 8:00–20:00, scale to zero otherwise; never warm 24/7.
- Max 4 vCPU / 8 GiB per replica; no batch or ETL; consumers must tolerate a cold start (up to ~30 s).
- AKS when total load is above 100 tps, use is constant more than 10 h a day with more than 15
  microservices, or a multi-application platform is needed.

The `Break_even` sheet tests these thresholds against the cost model. Today every AKS cluster serves a
single application, so the model compares one application against a whole cluster.

## Run order

### Setup (once per machine)

Python 3.10+, the Azure CLI and `uv` (or plain `pip`). Commands run from the repository root.

```bash
uv venv .local/venv                       # or: python -m venv .local/venv
source .local/venv/bin/activate           # PowerShell: .local\venv\Scripts\Activate.ps1
uv pip install -r aks-aca/src/requirements.txt
az login                                  # bank tenant; the scripts reuse this session
az account show --query tenantId -o tsv   # check it is the bank tenant before running
```

`.local/` is gitignored. The scripts sign in through `DefaultAzureCredential`, which reads the `az`
session; if `AZURE_CONFIG_DIR` is set in the shell, `az login` and the scripts must share it.

### Run

Run in office hours (Lima). Every script writes to `aks-aca/results/aa-<source>/` and overwrites the
previous run: copy `results/` first to keep it. Without dates, every script reads the last full calendar
month (Lima), the same as Cost Management's "last month"; the scripts need no date arguments.

```bash
python aks-aca/src/aa_inventory.py
python aks-aca/src/aa_metrics.py
python aks-aca/src/aa_container_insights.py
python aks-aca/src/aa_cost.py
python aks-aca/src/aa_prices.py
python aks-aca/src/aa_k8s_snapshot.py --snapshot-date <export-day>   # needs the exports below; the date is required
python aks-aca/src/aa_cluster_snapshot.py   # criteria v2 inputs, after the snapshot
python aks-aca/src/aa_signals.py   # optional; after aa_metrics.py
python aks-aca/src/aa_microservice_history.py   # optional; at least once a month, keeps a growing history
python aks-aca/src/aa_workbook.py
python aks-aca/src/aa_deck.py        # after opening the workbook in Excel and saving it
python aks-aca/src/aa_deck_exec.py
python aks-aca/src/aa_criteria_workbook.py
python aks-aca/src/aa_criteria_workbook_v2.py   # criteria v2, after aa_cluster_snapshot.py
python aks-aca/src/aa_deck_exec_v2.py           # corporate deck v2
```

| Step | Scope by default | Time on 2026-11-10 | Check |
|---|---|---|---|
| `aa_inventory.py` | Every visible subscription; one ARM list per subscription with runtimes (creation date) | 40 s (2026-11-18) | `inventory_errors.csv`; `apim_layer` in `apim_services.csv`; `existed_at_measurement` and `apim_apis_by_cod_app_source` in `inventory_summary.json` |
| `aa_metrics.py` | PROD clusters; the PROD APIM of every layer (`--apim-layers`) | ~75 min (AKS 15, APIM BS 45, UX ~10, Open Banking and Yape ~5 on 2026-11-18) | `metrics_errors.csv`; `apim_layers` in `metrics_summary.json` |
| `aa_container_insights.py` | PROD clusters with the add-on | ~80 min | `ci_errors.csv`, `ci_coverage.csv` |
| `aa_cost.py` | Every environment for AKS and ACA; PROD for App Service and Functions; `AmortizedCost` | ~10 min, plus one query per PROD subscription with App Service | `cost_errors.csv`; `app_service_subscriptions_queried` in `cost_summary.json` |
| `aa_prices.py` | VM sizes in the inventory; eastus2, centralus | < 1 min | `queries_without_rows` in `prices_summary.json` |
| `aa_k8s_snapshot.py` | Every row of the exports; the workbook counts PROD application clusters | < 10 s | `prod_clusters_by_category`, `cluster_names_shared_by_several_clusters`, `prod_application` in `snapshot_summary.json` |
| `aa_signals.py` | PROD clusters with the add-on; the same PROD APIM layers as `aa_metrics.py` | not timed yet (smoke test: 3 clusters + 1 APIM, 1 month, 20 s) | `signals_errors.csv` |
| `aa_microservice_history.py` | PROD clusters with the add-on; sample days not in the history yet | not timed yet | `history_errors.csv`, `history_summary.json` |
| `aa_workbook.py` | Everything above | ~30 s | Open in Excel and save once |

1. `aa_inventory.py` first: the other scripts read its CSVs.
2. `apim_layer` in `apim_services.csv` places each APIM in the north-south chain (UX -> BS -> Core -> mainframe)
   from its custom domain: BS is `apibs.lima.bcp.com.pe` (`apimeu2apigp01` and `apimeu2apigp05`, which took over
   BS traffic in October 2026), UX is `apiux`/`apisux` (`apimeu2apigp02`), Core is `apicr` (`apimeu2apigp03`);
   Open Banking (`apimeu2opbk*`) and Yape (`apimeu2yape*`) run their own; the Open Banking sandbox
   (`sandbox.apisbcp.com`) is `SANDBOX` and never measured. `aa_metrics.py` measures the five layers by default, plus
   two entry points outside APIM (`entry_points_monthly.csv`, `--skip-entry-points` to leave them out): every
   Application Gateway of an AKS application (AGIC sends to the pods, past APIM and the cluster load balancer) and
   the Event Hubs / Service Bus namespaces of applications with no API on APIM. Each application's traffic comes from
   one layer: BS when it has BS traffic, else the busiest of the others. Layers are never added, since a call goes on
   from one gateway to the next. Self-hosted gateways do not report to Azure Monitor, so traffic that only enters
   through one is not measured. An application without measured PROD traffic is `INSUFFICIENT_DATA`, out of the
   migration list ("Sin tráfico medido" on the corporate deck); `no_traffic_reason` says whether it is in use
   (CPU without a measurable entry), in development, an empty cluster or deployed without use. `--apim-environments
   PROD,DESA` on both scripts adds DESA test requests as a one-off check; it is not part of the monthly run
   (~3,600 more APIs). `--apim-names` still picks services
   by name for a smoke test.
3. `aa_metrics.py` next: platform metrics are kept for 93 days, so the last full month is always
   available. Windows that return more than 8 MB are split automatically.
4. After the inventory, `aa_metrics.py`, `aa_container_insights.py`, `aa_cost.py` and `aa_prices.py` can
   run at the same time in separate terminals: they use different APIs and limits. Start `aa_signals.py`
   when `aa_metrics.py` ends, since both call Azure Monitor metrics.
5. `aa_container_insights.py` only covers clusters with the add-on, and only as far back as the
   workspace retention; `ci_coverage.csv` shows empty months.
6. `aa_cost.py` without dates takes the last full month. Keep the same period as the metrics: the
   measured node discount only uses months that both sources cover end to end.
   App Service needs an inventory from 2026-11-17 or later (`runtime` and `plan_id`
   columns); pass `--app-service-environments NONE` to skip it.
7. `aa_metrics.py`, `aa_container_insights.py` and `aa_cost.py` keep each answer of a finished window in
   `results/<source>/_cache/` (`aa.QueryCache`, gzip JSON keyed by a hash of the resource, query and window), so a
   rerun or a resumed run only asks what is missing: failed queries and changed queries are asked again. Metrics
   and Container Insights cache a window 6 hours after it ends; cost, 5 days after the period ends (late usage).
   `--no-cache` asks everything again; deleting the folder does the same for good. `cache` in each summary JSON
   counts hits and queries.
8. `aa_prices.py` needs no sign-in.
9. `aa_k8s_snapshot.py` needs no sign-in either. It reads the platform team's `reportHpa*` and
   `reportDeployment*` exports (CSV with `;` or XLSX) from `.local/aks-aca-inputs/` (`--inputs` for another
   folder), keeps only the columns it needs and joins them with the Container Insights history by cluster
   id, namespace and deployment. The deployment export lists a deployment once per container of its pod; the
   snapshot adds them up (identical rows count once). The exports hold one day of configuration, so pass that day with
   `--snapshot-date`. Run it after `aa_container_insights.py` and `aa_metrics.py` (node size per cluster).
   `ValTrend` in the HPA export covers the last 30 days in ~11-hour averages; the replica history of the
   period comes from Container Insights.
   It classifies every cluster as `APPLICATION`, `INFRASTRUCTURE`, `NO_APPLICATIONS` or `NO_DATA` (rules in
   [docs/data-guide.md](docs/data-guide.md)); only APPLICATION clusters enter the workbook figures. To
   correct one by hand, add `cluster_overrides.csv` (`cluster_id` or `cluster`, `category`, `reason`) to the
   inputs folder. It also writes the inputs of the `Patterns` sheet (microservices per layer and per
   cluster, capabilities repeated across namespaces, time without scaling) and of part of
   `Technical_criteria`; without the exports both sheets skip those rows. `BIA_candidates` crosses the
   Container Apps technical fit per cluster with its BIA; which BIA levels stay on AKS is an editable cell.
   `ACA_by_cluster` compares each PROD application cluster's cost with Container Apps running the replicas
   Container Insights measured (sized by request and by use, 24/7 and with the ADT_019 schedule).
   `APIM_traffic` and `Traffic_profile` keep, for each application, only the layer it is measured on (`traffic_layer`
   in `application_snapshot.csv`, entry points included), so no call is counted twice.
   `Traffic_profile` classifies each API as OFFICE, EXTENDED or 24x7 from `apim_hourly_profile.csv`; the
   windows and threshold are editable cells, so reclassifying needs no new Azure query.
10. Optional: `aa_signals.py` checks the other side of the usage figures: hours with unschedulable pods
   per cluster (only clusters with the autoscaler emit the metric), OOMKilled containers per workload,
   stack and memory per replica split into app and Istio sidecar (last 7 days), and APIM backend latency
   per API and month (the same APIM layers and environments as `aa_metrics.py`; `apim_layer` on every row). It runs on its own after the inventory and writes to
   `results/aa-signals/`; the workbook does not read it yet.
11. Optional: `aa_microservice_history.py` counts the microservices running in each PROD cluster on the 1st and
    the 15th of each month (10:00-11:00 Lima) from Container Insights, as far back as the workspace keeps. It
    writes to `results/aa-history/`, which no script clears: each run adds the missing days and keeps the rest,
    so run it at least once a month. `aa_k8s_snapshot.py` adds one row per export day to the same folder. The
    workbook's `Microservice_history` sheet and slide 4 of the technical deck show the trend and the applications
    that grew the most.
12. `aa_workbook.py` writes `reports/<date>-aks-aca/aks_aca_comparison.xlsx`. Add a `README.md` next to
    it with the scope and gaps of the run (see `reports/2026-11-10-aks-aca/README.md` on the machine that
    ran it).
13. Optional: `aa_deck.py` builds `aks_aca_deck.pptx`, the technical presentation for architects (diagnosis,
    technical criteria, strategic front, proposal and annex; 20 slides, native charts, text in Spanish) from the
    workbook's calculated values and `results/` (signals, and node and workload usage for the CPU vs memory
    slide). Open the workbook in Excel and save it first. Its figures are a draft to check.
14. `aa_deck_exec.py` builds `aks_aca_corporate.pptx`, the corporate deck (slides 1 to 6, text in Spanish, amounts in
    soles with IGV) for a non-technical forum: ways to build an application (runtimes as acronyms, with a legend),
    AKS in PROD (clusters, applications, "micro"-services, monthly and annual cost of PROD plus non-PROD; Azure icon
    from `assets/icons/`), the hypothesis (three signs with a sentence each, motivation, minimum AKS vs Container
    Apps), the two conditions to move to Container Apps (up to 100 TPS in the peak hour, up to 15 "micro"-services)
    plus the off-hours rule, the result per group (conditions failed, where it goes, applications on AKS and ACA)
    and the possible saving with and without discount. ADT_019 is the reference, not the rule: every Container Apps estimate keeps one
    replica per "micro"-service on outside office hours. Runtime counts only take what existed in the measured
    month. It reads only `results/` (including `application_snapshot.csv`), so it needs no saved workbook. AKS
    takes the discount measured on its node VMs (reservations included); `--aks-discount 0.60` fixes it by hand.
    Container Apps takes the bank's 30%.
15. `aa_criteria_workbook.py` writes `aks_aca_criteria.xlsx` next to the deck, plain Excel in six sheets: `Results`
    (one row per destination in clusters and applications, as the result slide, and an editable chart), `Criteria` (one row per application:
    data, conditions, decision, off-hours replicas (nights and weekends) — 1 by default, write 0 to let one scale to zero — and its monthly
    cost on AKS and Container Apps in soles), `Calculator` (one application meter by meter, in the format of the
    team's calculator), `Services` (one row per "micro"-service: application, cluster, request, office replicas, HPA and
    measured use; filter it by application or cluster; `Criteria` sizes each one for Container Apps with `min_vcpu`
    and `vcpu_step` from `Variables`), `Prices` (with links to the sources), `Variables` and `Data`. `Data` and `Services` are Power
    Query connections (queries `Apps` and `Services`): it reads the snapshot CSVs and joins each "micro"-service to its application.
    `Criteria` takes each row from it. To check it on another machine: Enable Content, edit the query and change
    `Folder` (its first step) to that machine's `aks-aca/results`, then Data > Refresh All; no figure should change.
    If a refresh brings more applications, drag the corner of the `Criteria` table down. With the defaults the
    workbook gives the deck's figures. The query block starts from the one Excel wrote for Microsoft's
    connected-workbooks library (`assets/powerquery/`, MIT); only Python is needed.

Smoke tests before a full run: `aa_metrics.py --clusters <one-cluster> --skip-apim --start <day>
--end <day>` and `aa_container_insights.py --clusters <one-cluster> --start <day> --end <day>`, each
with `--output-dir .local/<name>` so `results/` is not overwritten.

Subscriptions with "POC" in their name are left out by the inventory, so every later step skips them
(`--include-poc` keeps them). Expected gaps with a Reader account (2026-11-10): 4 PROD clusters without
Container Insights; one stopped PROD cluster without metrics.

### Criteria v1 and v2

- **v1** is frozen at git tag `v1` (2026-11-24): `aa_deck_exec.py` and `aa_criteria_workbook.py`, one row per
  application, Container Apps costed with today's office replicas. To rebuild it exactly: `git checkout v1`, copy
  `.local/v1-data/results` over `results/` and run the two scripts.
- **v2** works per cluster, with a summary per application, and costs Container Apps from each hour's traffic:
  - `aa_cluster_snapshot.py` attributes every PROD AKS API to its cluster and writes `cluster_criteria_input.csv`
    and `traffic_hourly_profile.csv` (min / average / peak TPS by weekday or weekend and Lima hour, for every
    cluster and application) in `results/aa-cluster-snapshot/`.
  - `aa_v2_model.py` holds the rules and the cost model (its docstring explains them); `aa_deck_exec_v2.py`
    builds `aks_aca_corporate_v2.pptx` from it, and `aa_criteria_workbook_v2.py` writes the same rules as Excel
    formulas in `aks_aca_criteria_v2.xlsx` (Results, Criteria per cluster, Apps, Calculator with the replicas of
    every hour editable, Latency simulator, Traffic, Services, Prices, Variables). With the default variables the
    workbook and the deck give the same figures.
  - Replicas per hour = max(minimum, ceiling(TPS x traffic_share / concurrent_requests)): Container Apps adds a
    replica every `concurrent_requests` requests per second (HTTP scale rule, Azure default 10). The peak of an
    hour is its busiest 5 minutes; only minimum, average and peak are shown, never percentiles.
  - The saving margin (`min_saving`, 10% by default) is applied at the end, not as a criterion.

## Tracing a figure

- `Sources`: each data set with its type (measured, derived, estimated, vendor data), source and
  how to check it in the portal.
- `Queries`: the exact Resource Graph queries, ARM call patterns, Log Analytics queries, Cost
  Management bodies and Retail Prices filters each script ran.
- `Assumptions`: yellow cells are inputs; column E shows the measured reference for each one, when one
  exists. Everything in `Model`, `Break_even` and `Map` is a formula over `Assumptions` and `Prices`.

## Caveats

- Node VMs are covered by reservations or savings plans: with `ActualCost` they cost close to 0 per
  hour (checked on 2026-11-10), so `aa_cost.py` uses `AmortizedCost`. Cost analysis screenshots must
  use the "Amortized cost" view too.
- List prices exclude contract discounts; `discount` in `Assumptions` is the measured ratio for node
  VMs and feeds the "AKS efectivo" line. ACA has no measured discount unless the bank already runs
  Container Apps.
- `Environment Management Hour` and `Environment Private Endpoint` are ACA meters effective
  2026-09-01 in eastus2 at 0.10 USD/h each (Retail Prices API, checked 2026-11-10). Microsoft Learn
  does not say which environments they apply to: the billing page still bills private endpoints and
  planned maintenance as `Dedicated Plan Management`, and private endpoints need a workload-profiles
  environment. The model charges both by default; `aa_cost.py` shows whether the bank's existing ACA
  environments are billed for them.
- The free grant (180,000 vCPU-s, 360,000 GiB-s, 2 million requests) is per subscription and month,
  for the Consumption plan. Learn does not say whether it covers the Consumption profile of a
  workload-profiles environment.
- Requests per microservice come from APIM per API. The API → application mapping uses the backend host
  (`backend_cod_app` and `cod_app_source` in `apim_apis.csv`); only APIs whose `backend_platform` is AKS count as
  AKS traffic, so APIs served by OpenShift or App Service are not added to an AKS application.
- The inventory is current while cost and metrics cover a past month: `existed_at_measurement` on every runtime
  (AKS, Container Apps, App Service, Functions) says whether it existed for the whole month (YES), part of it
  (PARTIAL) or not yet (NO).
- Off-hours wake-ups are estimated from 5-minute buckets: a request inside a bucket counts the whole
  bucket as active, plus a 300 s cooldown per run.
- Office hours do not exclude public holidays.
- Coverage depends on the account: subscriptions, workspaces or APIM services it cannot read are
  missing and are listed in each `*_errors.csv`.
