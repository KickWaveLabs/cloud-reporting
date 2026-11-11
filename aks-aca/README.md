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

## Sources

| Need | Source | Script | Data type |
|---|---|---|---|
| Clusters, node pools, add-ons, ACA apps, APIM APIs | Resource Graph; ARM GET `{apim}/apis` | `aa_inventory.py` | Measured |
| Node CPU and memory, requests per API | Azure Monitor platform metrics (GET) | `aa_metrics.py` | Measured |
| Workloads, replicas, CPU/memory per replica | Container Insights (Log Analytics, GET) | `aa_container_insights.py` | Measured |
| Actual cost of clusters and ACA environments | Cost Management query API, `AmortizedCost` | `aa_cost.py` | Measured |
| List prices | Azure Retail Prices API (public, no sign-in) | `aa_prices.py` | Vendor data |

What is never read: tags (only used to derive the environment), identities, APIM subscription keys and
named values, Container Apps secrets and environment variables, container logs, pod labels.

## ADT_019 rules used in the model

- ACA in Consumption (serverless) mode; dedicated workload profiles need architecture approval.
- 1 replica Monday to Friday 8:00–20:00, scale to zero otherwise; never warm 24/7.
- Max 4 vCPU / 8 GiB per replica; no batch or ETL; consumers must tolerate a cold start (up to ~30 s).
- AKS when total load is above 100 tps, use is constant more than 10 h a day with more than 15
  microservices, or a multi-application platform is needed.

The `Equilibrio` sheet tests these thresholds against the cost model. Today every AKS cluster serves a
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
previous run.

```bash
python aks-aca/src/aa_inventory.py
python aks-aca/src/aa_metrics.py --apim-names apigp01,apigp03,apigp05
python aks-aca/src/aa_container_insights.py
python aks-aca/src/aa_cost.py --from-date 2026-08-01 --to-date 2026-10-31
python aks-aca/src/aa_prices.py
python aks-aca/src/aa_k8s_snapshot.py --snapshot-date 2026-11-10   # optional, needs the exports below
python aks-aca/src/aa_workbook.py
```

| Step | Scope by default | Time on 2026-11-10 | Check |
|---|---|---|---|
| `aa_inventory.py` | Every visible subscription | 20 s | `inventory_errors.csv`; APIM `custom_domains` in `apim_services.csv` |
| `aa_metrics.py` | PROD clusters; the APIM in `--apim-names` | ~60 min (AKS 15, APIM 45) | `metrics_errors.csv` |
| `aa_container_insights.py` | PROD clusters with the add-on | ~80 min | `ci_errors.csv`, `ci_coverage.csv` |
| `aa_cost.py` | Every environment, `AmortizedCost` | ~10 min | `cost_errors.csv` |
| `aa_prices.py` | VM sizes in the inventory; eastus2, centralus | < 1 min | `queries_without_rows` in `prices_summary.json` |
| `aa_k8s_snapshot.py` | Every row of the exports; the workbook counts PROD | < 10 s | `clusters_not_in_inventory`, `hpa_without_ci_history` in `snapshot_summary.json` |
| `aa_workbook.py` | Everything above | ~30 s | Open in Excel and save once |

1. `aa_inventory.py` first: the other scripts read its CSVs.
2. Pick the APIM fragments from `custom_domains` in `apim_services.csv`: BS is `apibs.lima.bcp.com.pe`
   (`apimeu2apigp01` and `apimeu2apigp05`, which took over BS traffic in October 2026), Core is
   `apicr.lima.bcp.com.pe` (`apimeu2apigp03`). `apimeu2apigp02` is the UX layer and stays out.
3. `aa_metrics.py` next: platform metrics are kept for 93 days. Default window: the last 92 days up to
   yesterday (Lima). Windows that return more than 8 MB are split automatically.
4. After the inventory, `aa_metrics.py`, `aa_container_insights.py` and `aa_cost.py` can run at the
   same time in separate terminals: they use different APIs and limits.
5. `aa_container_insights.py` only covers clusters with the add-on, and only as far back as the
   workspace retention; `ci_coverage.csv` shows empty months.
6. `aa_cost.py` without dates takes the last three full months. Keep the period inside the metrics
   window: the measured node discount only uses months that both sources cover end to end.
   It caches each subscription in `results/aa-cost/_cache/<type>_<from>_<to>/`; rerun it to
   resume after throttling (429), delete the folder to fetch again.
7. `aa_prices.py` needs no sign-in.
8. `aa_k8s_snapshot.py` needs no sign-in either. It reads the platform team's `reportHpa*` and
   `reportDeployment*` exports (CSV with `;` or XLSX) from `.local/aks-aca-inputs/` (`--inputs` for another
   folder), keeps only the columns it needs and joins them with the Container Insights history by
   cluster, namespace and deployment. The exports hold one day of configuration, so pass that day with
   `--snapshot-date`. Run it after `aa_container_insights.py`. `ValTrend` in the HPA export covers a short
   recent window; the 3-month replica history comes from Container Insights. It also writes the inputs of
   the `Patrones` sheet (microservices per layer and per cluster, capabilities repeated across namespaces)
   and of part of `Criterios_tecnicos`; without the exports both sheets skip those rows. `Candidatos_BIA`
   crosses the Container Apps technical fit per cluster with its BIA; which BIA levels stay on AKS is an
   editable cell.
   `ACA_por_cluster` compares each PROD cluster's cost today with Container Apps running the replicas
   Container Insights measured (sized by request and by use, 24/7 and with the ADT_019 schedule).
   `Perfil_trafico` classifies each API as office, extended or 24/7 from `apim_hourly_profile.csv`; the
   windows and threshold are editable cells, so reclassifying needs no new Azure query.
9. `aa_workbook.py` writes `reports/<date>-aks-aca/aks_aca_comparison.xlsx`. Add a `README.md` next to
   it with the scope and gaps of the run (see `reports/2026-11-10-aks-aca/README.md` on the machine that
   ran it).

Smoke tests before a full run: `aa_metrics.py --clusters <one-cluster> --skip-apim --start <day>
--end <day>` and `aa_container_insights.py --clusters <one-cluster> --start <day> --end <day>`, each
with `--output-dir .local/<name>` so `results/` is not overwritten.

Subscriptions with "POC" in their name are left out by the inventory, so every later step skips them
(`--include-poc` keeps them). Expected gaps with a Reader account (2026-11-10): 4 PROD clusters without
Container Insights; one stopped PROD cluster without metrics.

## Tracing a figure

- `Fuentes`: each data set with its type (measured, derived, estimated, vendor data), source and
  how to check it in the portal.
- `Consultas`: the exact Resource Graph queries, ARM call patterns, Log Analytics queries, Cost
  Management bodies and Retail Prices filters each script ran.
- `Supuestos`: yellow cells are inputs; column E shows the measured reference for each one, when one
  exists. Everything in `Modelo`, `Equilibrio` and `Mapa` is a formula over `Supuestos` and `Precios`.

## Caveats

- Node VMs are covered by reservations or savings plans: with `ActualCost` they cost close to 0 per
  hour (checked on 2026-11-10), so `aa_cost.py` uses `AmortizedCost`. Cost analysis screenshots must
  use the "Amortized cost" view too.
- List prices exclude contract discounts; `discount` in `Supuestos` is the measured ratio for node
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
- Requests per microservice come from APIM per API. The API → microservice mapping uses the backend
  host and is approximate.
- Off-hours wake-ups are estimated from 5-minute buckets: a request inside a bucket counts the whole
  bucket as active, plus a 300 s cooldown per run.
- Office hours do not exclude public holidays.
- Coverage depends on the account: subscriptions, workspaces or APIM services it cannot read are
  missing and are listed in each `*_errors.csv`.
