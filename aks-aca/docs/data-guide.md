# Data guide: what each CSV holds and how to build your own workbooks

Every figure in the deck comes from a CSV in `aks-aca/results/`. This guide says which script writes
each CSV, what one row is, what the main columns mean, and how to rebuild the deck figures with pivot
tables, without `aa_workbook.py`.

Office hours everywhere: Monday to Friday, 8:00 to 19:59 Lima time. Holidays count as working days.

Period: by default every script reads the last full calendar month (Lima), the same as Cost
Management's "last month". `--start`/`--end` (`--from-date`/`--to-date` in `aa_cost.py`) pick another.

App code: every CSV keyed by cluster, cost owner or API carries `cod_app`, the 4-character application
code. Inventory rows take it from the `codApp` tag, else from the naming convention
(`aksveu2coaap01` -> `COAA`); `cod_app_evidence` says which (`TAG`, `NAME`, or `ENVIRONMENT` for a
Container App that inherits its environment's code). For APIs, `cod_app` is the code in the AKS ingress
host of the backend (`azingressscrmeu2.lima.bcp.com.pe` -> `SCRM`); it is empty for App Service or
external backends. Pivot by `cod_app` to see one application across clusters, cost and traffic.

## The scripts at a glance

| Script | Question it answers | Azure source | Output folder |
|---|---|---|---|
| `aa_inventory.py` | Which clusters, node pools, Container Apps, App Services and APIM APIs exist, and how are they configured? | Resource Graph; ARM GET for the APIM APIs | `aa-inventory` |
| `aa_metrics.py` | How much CPU and memory do the nodes use? How many requests reach each API, and when? | Azure Monitor platform metrics (cluster and APIM resources) | `aa-metrics` |
| `aa_container_insights.py` | Which workloads run, with how many replicas, and how much does each replica use and request? | Log Analytics (Container Insights tables) | `aa-container-insights` |
| `aa_cost.py` | How much did each cluster and Container Apps environment cost in the month? | Cost Management query API, `AmortizedCost` | `aa-cost` |
| `aa_prices.py` | What are the list prices of VMs and Container Apps? | Azure Retail Prices API | `aa-prices` |
| `aa_signals.py` | Are clusters short of capacity or memory? How slow is each API's backend? | Azure Monitor platform metrics; Container Insights | `aa-signals` |
| `aa_microservice_history.py` | How many microservices ran on each sample day, and which applications grew? | Container Insights | `aa-history` |
| `aa_k8s_snapshot.py` | Which microservices exist, with which HPA and BIA, and how did they behave in the month? | No Azure call: the platform team's exports joined with Container Insights | `aa-k8s-snapshot` |

## The CSVs

### aa-inventory

| CSV | One row is | Main columns |
|---|---|---|
| `aks_clusters.csv` | A cluster | `name`, `environment` (from the tag, else the name suffix), `created_at`, `measurement_month`, `existed_at_measurement` (YES = existed the whole month, PARTIAL = created during it, NO = created after it; the same three columns are on every runtime CSV below), `kubernetes_version`, `system_nodes` and `user_nodes` (nodes at the time of the run, not the minimum), `vm_sizes`, `upgrade_channel` (`none` = manual upgrades), `service_mesh`, `gitops_flux`, `container_insights`, plus one column per configuration property |
| `aks_node_pools.csv` | A node pool | `cluster`, `mode` (System/User), `vm_size`, `count` (now), `autoscaling`, `min_count`, `max_count`, `zones`, `max_pods` |
| `aca_environments.csv`, `aca_apps.csv` | A Container Apps environment or app | plan, workload profiles, min and max replicas, CPU and memory per replica |
| `apim_services.csv` | An APIM instance | `sku`, `capacity`, `custom_domains`, `apim_layer` (BS, CORE, UX, OPEN_BANKING, YAPE, SANDBOX, OTHER; from the custom domain) |
| `app_gateways.csv`, `messaging_namespaces.csv` | An Application Gateway; an Event Hubs or Service Bus namespace | `cod_app`, `environment`, `sku` (entry points outside APIM, measured in `entry_points_monthly.csv`) |
| `apim_apis.csv` | An API | `apim_layer`, `api_id`, `path`, `backend_host`, `backend_platform` (AKS, APP_SERVICE, CONTAINER_APPS, OPENSHIFT, APIM, OTHER, NONE), `backend_cod_app` (the owning application) and `cod_app_source` (HOST, HOST_NORMALIZED, K8S_NAMESPACE, BACKEND_RESOURCE, NAMING; rules in `aa_common.resolve_api_cod_app`) |
| `app_service_plans.csv` | An App Service plan | `sku`, `tier`, `workers`, `os`, `sites`, `runtime` (`FUNCTIONS` for a Consumption, Premium or Flex Consumption plan, or one whose sites are all function apps; else `APP_SERVICE`) |
| `app_services.csv` | A web app or function app | `type` (`WEB`/`FUNCTION`), `container` (`YES` = Docker image), `state`, `plan`, `plan_sku`, `plan_id` |

The configured node minimum of a cluster is the sum over its pools of `min_count` when `autoscaling`
is true, else `count`.

### aa-metrics

| CSV | One row is | Main columns |
|---|---|---|
| `aks_usage_monthly.csv` | Cluster x pool x month (usage rows), or cluster x month with `mode = ALLOCATABLE` (capacity row) | Usage rows: `nodes_avg`, `cpu_used_office_avg`, `cpu_used_offhours_avg`, `cpu_used_p95`, `cpu_used_max` (vCPU) and the same for memory (GiB). Capacity rows: `cpu_allocatable_avg`, `mem_allocatable_avg_gib` |
| `apim_api_monthly.csv` | APIM x API x month | `apim_layer` (BS, UX, CORE, OPEN_BANKING, YAPE), `requests`, `office_requests`, `offhours_requests`, `peak_5min_requests` (divide by 300 for requests per second), `peak_hour_requests`, `offhours_wakeups` |
| `apim_hourly_profile.csv` | APIM x API x day type (`weekday`/`weekend`) x Lima hour | `apim_layer`, `avg_requests_per_hour` |
| `entry_points_monthly.csv` | Entry point outside APIM x month | The same columns as `apim_api_monthly`; `apim_layer` is `APP_GATEWAY` (`TotalRequests` of the Application Gateways of AKS applications) or `EVENT_HUB` / `SERVICE_BUS` (`IncomingMessages` of applications with no AKS-backed API: messages, not requests) |
| `entry_points_hourly_profile.csv` | Entry point x day type x Lima hour | `apim_layer`, `avg_requests_per_hour` |

- Usage comes from `node_cpu_usage_millicores` and `node_memory_working_set_bytes`, summed over the
  nodes of each pool, hour by hour.
- Allocatable comes from `kube_node_status_allocatable_cpu_cores` and `_memory_bytes`. It is the
  cluster total for the nodes running in each hour, system and user pools together.
- For CPU, allocatable is almost the full VM: 2 x 8 vCPU nodes give 16.0. For memory AKS reserves a
  part: 2 x 32 GiB nodes give about 61 GiB.
- Use % = sum of usage / sum of allocatable. Both rows types have `environment`, so filter PROD and
  divide the two sums.
- **Never add APIM layers.** One call from a channel goes through the UX, BS and Core APIM, so the same
  request shows in each. Figures of the estate filter `apim_layer` = BS; figures per application come from
  `application_snapshot`, which already took one layer per application (`traffic_layer`).

### aa-container-insights

| CSV | One row is | Main columns |
|---|---|---|
| `aks_workloads_monthly.csv` | Cluster x namespace x workload x month x period (`OFFICE`/`OFFHOURS`) | `kind` (`ReplicaSet` = a deployment, `StatefulSet`, `DaemonSet`, `Job`), `system_namespace`, `hours` (hours the workload existed), `replicas_avg`, `replicas_max`, `cpu_used_avg`/`_p95`/`_max`/`_per_replica`, `cpu_request_*`, `mem_used_gib_*`, `mem_request_gib_*` |
| `aks_services.csv` | A Kubernetes Service | `namespace`, `service`, `service_type` |
| `ci_coverage.csv` | Cluster x month | `status`: `OK`, `EMPTY` (no data in the workspace), `ERROR` |

- Averages are over the hours the workload existed, so a workload that lived one month is not diluted
  over three.
- A rollout briefly counts old and new pods in the same hour, so `replicas_avg` is slightly high.
  This makes "HPA at its minimum" a lower bound.
- Usage and request include every container of the pod, Istio sidecar included.

### aa-cost

`cost_by_resource.csv`: one row per resource x meter x month.

- `owner` is the cluster (or Container Apps environment) the resource belongs to. That covers the
  cluster resource and everything in its node resource group: VMs, disks, load balancers, IPs.
  For App Service it is the plan, which also takes the cost of its sites and slots.
- `owner_type` is `AKS`, `ACA`, `APP_SERVICE` or `FUNCTIONS` (the plan's `runtime`). App Service and
  Functions cover PROD only unless `--app-service-environments` says otherwise; a plan shared by web
  and function apps counts as `APP_SERVICE`.
- `cost_usd` is amortized cost: reservations and savings plans are spread over the month, so node VMs
  do not show as 0.
- `meter_category` splits the cost. For AKS PROD in October 2026: Virtual Machines 76%, Storage 7%,
  Virtual Network 6%, Bandwidth 4%, the AKS tier 3%, Load Balancer 3%.

### aa-prices

`prices.csv`: one row per price used by the cost model. `key` names it (for example the node VM or the
ACA vCPU-second), `retail_price` per `unit`, and `hourly_equivalent`. Only the model sheets use it.

### aa-k8s-snapshot

Configuration comes from the platform team's exports (the export day); behaviour comes from Container
Insights (the period). Every row carries `cluster_id` and `cod_app`; clusters are joined by resource id.

| CSV | One row is | Main columns |
|---|---|---|
| `cluster_snapshot.csv` | A cluster (inventory and export) | `category`, `category_reason`, `in_inventory`, `in_export`, `microservices`, `tool_deployments`, `ci_app_workloads`, `namespaces`, `hpas`, `bia` (highest among its microservices), `statefulsets`, `aca_technical_fit` |
| `deployment_snapshot.csv` | A deployment | `microservice` (`YES` = an application deployment), `platform_component`, `layer`, `bia`, `capabilities`, `cpu_request_vcpu`, `mem_request_gib`, `has_hpa`, `ci_replicas_avg_office`, `ci_cpu_used_per_replica`, `ci_request_minus_export_vcpu`, `node_vcpu`, `node_mem_gib`, `pod_too_big_for_node`, `fits_aca_by_request` (request within 4 vCPU / 8 GiB; empty when the export has no request) |
| `hpa_snapshot.csv` | An HPA | `microservice`, `hpa_min`, `hpa_max`, `replicas_now` (export day), `last_scale_time`, `days_since_last_scale`, `last_scale_bucket`, `ci_replicas_avg_office`, `ci_replicas_max_office`, `office_max_reached_hpa_max`, `behaviour_ci`, `no_scale_cause`, `replicas_needed_by_cpu`, `min_above_need` |
| `capability_domains.csv` | A capability keyword | `namespaces_prod` (domains), `deployments_prod`, both in APPLICATION clusters |
| `application_snapshot.csv` | A PROD application (app code of APPLICATION clusters) | `clusters`, `cost_month_usd`, `bia`, `microservices`, `microservices_desa`, `microservices_new_in_desa` and `microservices_projected` (PROD plus the DESA microservices not in PROD yet, compared by name without the version suffix; empty when the export has no DESA cluster for the application), `hpas_not_scaling_share`, `traffic_status` (MEASURED, ZERO_IN_MONTH, LAYER_NOT_MEASURED, NO_APIM_API), `office_vcpu_used` (office-hours vCPU of its PROD microservices, Container Insights), `traffic_layer` (BS, UX, OPEN_BANKING, YAPE, CORE, APP_GATEWAY, or EVENT_HUB / SERVICE_BUS for an event-driven application with no API on APIM: messages instead of requests), `traffic_note`, `prod_cluster_created` and `prod_cluster_age_days` (newest PROD cluster, age at the start of the month), `no_traffic_reason` (only for INSUFFICIENT_DATA: IN_DEVELOPMENT = PROD cluster under 6 months, new microservices in DESA or at least 1,000 DESA test requests in the month; EMPTY_CLUSTER = no microservice running; NOT_IN_USE = deployed 6 to 12 months without use, ask the team; NOT_IN_USE_OVER_1_YEAR = cluster over a year old without use, candidate to decommission), `desa_requests_month` (test requests of its DESA APIs, only when the DESA APIM was measured as a one-off check; never added to PROD), `apis_with_traffic`, `requests_month`, `office_share`, `peak_tps_upper_bound`, `busiest_hour_tps`, `busiest_hour` (day type and Lima hour), `busiest_hour_in_office`, `busiest_office_hour_tps`, `busiest_offhours_hour_tps`, `weekday_hours_with_traffic`, `offhours_active_share`, `batch_jobs`, `statefulsets`, `microservices_over_aca_size`, `possible_monolith`, `adt_runtime`, `adt_reason` |

#### Definitions

| Term | Definition |
|---|---|
| Application | An app code (`cod_app`) with at least one APPLICATION cluster |
| Microservice | A deployment outside platform namespaces and not a known tool (`aa_common.is_microservice`) |
| Domain | A namespace with at least one microservice |
| Deployment | Any row of the export, unfiltered |

Platform namespaces, tool prefixes, the application's own gateway prefixes and the infrastructure app codes are
defined once, in `aa_common.py`.

#### Cluster category

Only APPLICATION clusters enter the analysis. `cluster_overrides.csv` in the inputs folder (columns
`cluster_id` or `cluster`, `category`, `reason`) corrects a cluster by hand.

| Category | Rule |
|---|---|
| `INFRASTRUCTURE` | App code in `aa_common.INFRASTRUCTURE_APPS` (runners, Terraform agents, SonarQube, observability, shared gateways), or only tool workloads (for example chaos testing) |
| `APPLICATION` | At least one microservice in the export or in Container Insights, or only the application's own gateway |
| `NO_APPLICATIONS` | In the export or with Container Insights data, but without a single microservice: created before its application, or left empty |
| `NO_DATA` | Not in the export and without Container Insights data |

#### Application runtime (ADT_019)

`application_snapshot.csv` applies the ADT_019 tree to each application ([adt-019-runtime.md](adt-019-runtime.md)):

| `adt_runtime` | Rule |
|---|---|
| `AKS` | Busiest hour above 100 TPS; or traffic more than 10 weekday hours and more than 15 microservices (PROD or projected with DESA); or traffic in more than 50% of the off-hours time (`offhours_active_share`: cannot scale to zero, and ADT_019 never keeps Container Apps warm 24/7); or a microservice requests more than 4 vCPU / 8 GiB |
| `ACA_WITH_REVIEW` | Within the thresholds but with BIA ALTA or MUY ALTA, no APIM traffic mapped, or 15 microservices or fewer in PROD and more once DESA is added (ADT_019 asks to review AKS when a Container Apps application passes 15). StatefulSets and batch jobs are listed in their columns but are not a criterion for now |
| `ACA` | Within every threshold |
| `INSUFFICIENT_DATA` | No PROD traffic in the month on any layer (`traffic_status` other than MEASURED). Left out of the migration list; `no_traffic_reason` splits it (IN_USE_NOT_MEASURED: its microservices use at least 0.5 vCPU in office hours, so traffic enters where nothing measures it, such as a self-hosted gateway; IN_DEVELOPMENT; EMPTY_CLUSTER; NOT_IN_USE; NOT_IN_USE_OVER_1_YEAR). The reason repeats `traffic_note` and the DESA microservices |

- Traffic: each API maps to an app code through its backend (`backend_cod_app`); a self-hosted gateway backend
  (`<service>.<namespace>.svc.cluster.local`) takes the code of the PROD cluster that runs that namespace. Only
  AKS backends count. One layer per application (`traffic_layer`): BS when it has BS traffic in the month, else
  the busiest of UX, Open Banking, Yape, Core, its Application Gateway (`APP_GATEWAY`) and the Event Hubs / Service
  Bus messages of an application with no AKS-backed API on APIM (`entry_points_monthly.csv`); never added, since a call goes on from one gateway to the next. `traffic_status`
  says why an application has no traffic and `traffic_note` what the data shows: a new cluster (created less than
  6 months before the month), only event consumers (`async-*`), its own ingress load balancer (reached without
  APIM) or no microservice running. `busiest_hour_tps` adds
  the application's APIs hour by hour (average weekday or weekend profile of the month); `peak_tps_upper_bound`
  adds each API's 5-minute peak, so it overstates the real peak.
- An hour has traffic when it carries at least 10% of the busiest hour; `weekday_hours_with_traffic` counts them.
- `busiest_hour_in_office` = NO means the application's typical busiest hour falls outside the ADT_019 warm
  window (weekdays 8:00-19:59): the office-hours model of Container Apps would meet its peak with the replica
  off or cold. `busiest_office_hour_tps` and `busiest_offhours_hour_tps` are the busiest hour inside and outside
  that window; `busiest_hour_tps` is the larger of the two.
- `offhours_active_share` = off-hours 5-minute intervals with traffic / every off-hours interval of the month,
  for the application's busiest API (a floor: two APIs active in the same interval count once). It measures
  coverage, not load: 0.1 TPS spread over the night keeps a replica warm just as much as 10 TPS.
- `possible_monolith` = a single microservice in PROD and no more once DESA is added (an application with more
  microservices in DESA is growing). The tree sends internal monoliths to App Service, but a single
  deployment is not proof of a monolith: confirm with the application team.
- The tree's questions on Azure-event functions and multi-application platforms are not tested: every cluster
  serves one application today.

#### HPA codes

`behaviour_ci` compares the office-hours average of replicas in the period with the HPA limits:

| Code | Rule |
|---|---|
| `NEVER_SCALES` | Average <= 1.1 x min |
| `AT_MAX` | Average >= 0.9 x max |
| `FIXED` | min = max |
| `SCALES` | Anything else |
| `NO_DATA` | No Container Insights history |

`last_scale_bucket` groups `days_since_last_scale` (from `lastScaleTime` to the export day): `UNDER_1_MONTH`,
`1_TO_3_MONTHS`, `3_TO_12_MONTHS`, `OVER_1_YEAR`, `NO_DATA`. `no_scale_cause` says why an HPA does not scale:
`FIXED_MIN_MAX` (it cannot; the date is when it was set), `MIN_ABOVE_NEED` (the minimum exceeds what the CPU
needs), `LOAD_BELOW_TARGET` (the load never reaches the target), `SCALES`, `NO_DATA`.

`replicas_needed_by_cpu` = office CPU used by all replicas / (request per replica x 0.7, a typical HPA
target). `min_above_need_beyond_ha` = the minimum is above max(2, needed). `ci_replicas_max_office` is the
hourly maximum and overstates scaling during rolling deployments.

`pod_too_big_for_node` = the request is above half of a node's allocatable CPU or memory (node size from
`aks_usage_monthly`, approximate), so each new replica needs its own node. `ci_request_minus_export_vcpu`
above 0 means the export holds only the application container and Container Insights adds the Istio sidecar.

### aa-history

This folder is never cleared: each run adds rows and keeps the old ones, so it holds the history that Container
Insights drops after its retention.

| CSV | One row is | Main columns |
|---|---|---|
| `microservice_history.csv` | Sample day x cluster (`aa_microservice_history.py`) | `date` (1st or 15th of a month), `cluster_id`, `cod_app`, `environment`, `microservices` (with a pod running at 10:00-11:00 Lima), `namespaces` |
| `export_history.csv` | Export day (`aa_k8s_snapshot.py`) | `snapshot_date`, `prod_application_clusters`, `prod_microservices`, clusters per category |

The workbook's `Microservice_history` sheet sums PROD clusters that are not infrastructure (clusters deleted since
keep their samples) and ranks the top 20 app codes by growth from their first sample to the last one.

### aa-signals

| CSV | One row is | Main columns |
|---|---|---|
| `aks_autoscaler_monthly.csv` | Cluster x month | `hours_with_data` (0 = no autoscaler, the metric is not emitted), `hours_with_unschedulable`, `office_hours_with_unschedulable`, `unschedulable_pods_max` (highest hourly average) |
| `aks_oomkilled_monthly.csv` | Cluster x namespace x workload x month | `oom_kills`, `pods` affected, `system_namespace`, `microservice` (`YES` = an application workload, same rule as `deployment_snapshot`) |
| `aks_stack_memory.csv` | Cluster x namespace x workload (last `--stack-days`, office hours) | `stack` (Quarkus, Spring Boot, Java, .NET, Node.js, Python, Unidentified), `replicas_avg`, `app_mem_gib_per_replica`, `sidecar_mem_gib_per_replica` (Istio; empty when the pod has no sidecar) |
| `apim_latency_monthly.csv` | APIM x API x month (same APIM layers as `aa_metrics.py`) | `apim_layer`, `backend_duration_avg_ms`, `backend_duration_max_ms` (APIM to the backend and back), `duration_avg_ms`, `duration_max_ms` (whole gateway request) |

- Unschedulable pods: `cluster_autoscaler_unschedulable_pods_count`, hourly average. An hour above 0
  had pods waiting for a node at some point.
- OOM kills: `KubePodInventory.ContainerLastStatus` = OOMKilled. One kill is one pod, container and
  restart count with that status, so a status repeated in every inventory snapshot counts once.
- Stack: classified inside the workspace from the command and image of every container of the pod
  (`ContainerInventory`). Most app containers start from the image entrypoint, so the command is empty;
  a Java pod is still recognised by the init container that copies a `javaagent.jar`. Spring Boot and
  Quarkus only show when the command or image names them, so "Java" usually means "framework unknown".
  The framework comes from the repositories (`pom.xml`) or the platform team.
- Latency averages are per request over the whole month (metric interval `FULL`). The maximum is a
  single slow request; use the average to compare with a Container Apps cold start (up to ~30 s).
  The cold start only hits the first call after the replica scaled to zero (outside office hours under
  ADT_019), so latency is a check on condition 3 ("can be turned off outside office hours"), not a criterion of its own.

## Building your own workbooks

Three workbooks cover the technical deck (section 6 covers the corporate deck): `cost_and_capacity_summary.xlsx` (everything with money),
`ms_and_hpa_summary.xlsx` (microservices, HPA and cluster health) and `apim_traffic_summary.xlsx`.

Open each CSV with *Data > From Text/CSV*, comma delimiter. Opening it with a double click splits the
columns wrong in a Spanish Excel. Use Power Query only to add columns; every figure below is a pivot.

Power Query notes:

- The source step stores the number of columns (`Csv.Document(..., [Delimiter=",", Columns=21, ...])`).
  When a script adds a column, the last ones are cut and a later step fails with "The column '...' of the
  table wasn't found". Remove `Columns=N,` from the source step (*Home > Advanced Editor*).
- The CSVs write decimals with a dot. In a Spanish Excel, convert numbers with the `en-US` locale, for
  example `Number.From([x], "en-US")`, or `3.0` becomes 30.
- **Category join.** Only `cluster_snapshot` has `category`. To keep `APPLICATION` clusters in any other
  CSV: *Home > Merge Queries*, the CSV's `cluster_id` with `cluster_snapshot[cluster_id]`, left outer;
  expand only `category`. The row count must not change.
- **Category lookup (no merge).** Same result with one step and no extra query: in the CSV's query open
  *Home > Advanced Editor* and add these steps before `in`, replacing `#"Tipo cambiado"` with the name of
  the last step. `cluster_snapshot` must be a query of the same workbook (connection only is enough).

  ```text
      Lookup = Table.Distinct(Table.SelectColumns(Table.SelectRows(cluster_snapshot,
          each [cluster_id] <> null and [cluster_id] <> ""), {"cluster_id", "category"}), {"cluster_id"}),
      Categories = Record.FromList(Lookup[category], List.Transform(Lookup[cluster_id], Text.Lower)),
      WithCategory = Table.AddColumn(#"Tipo cambiado", "category",
          each Record.FieldOrDefault(Categories, Text.Lower([cluster_id]), "NOT_IN_SNAPSHOT"), type text)
  in
      WithCategory
  ```

  `NOT_IN_SNAPSHOT` marks a cluster the export does not have.
- Pivot calculated fields work on sums: fine for a ratio of two sums (`office_requests / requests`), wrong
  for a per-row value such as a peak (add those as a Power Query column). Function names follow the Excel
  language (`SI`, not `IF`); use `;` if `,` fails.

### 1. Cost and capacity (`cost_and_capacity_summary.xlsx`)

Inputs: `aks_clusters.csv`, `aks_node_pools.csv`, `aks_usage_monthly.csv`, `cost_by_resource.csv`,
`cluster_snapshot.csv`.

| Figure | Pivot or formula |
|---|---|
| PROD clusters | `aks_clusters`, filter `environment` = PROD and `existed_at_measurement` other than NO, count `id` (two clusters can share a name; the inventory is current, the month is past) |
| Monthly cost | `cost_by_resource`, filter `owner_type` = AKS, `environment` = PROD; rows `month`; sum `cost_usd` |
| Cost by type | Same pivot with `meter_category` in rows |
| Cost per cluster | Same pivot with `owner_id` in rows |
| Clusters with 4 nodes or fewer now | `aks_clusters`: helper column `system_nodes + user_nodes`, count <= 4 |
| Clusters under the 4-node minimum | `aks_node_pools`: helper column `IF(autoscaling="true", min_count, count)`; pivot by `cluster`, sum it, count < 4 |
| CPU and memory use in office hours | `aks_usage_monthly`, filter PROD; sum of `cpu_used_office_avg` / sum of `cpu_allocatable_avg` (same for memory, p95 and max) |
| Use per cluster | Same pivot with `cluster` in rows |
| Clusters per microservices bucket (pattern 2) | `cluster_snapshot`, filter PROD, `category` = APPLICATION; helper bucket column on `microservices`; add the cost with `XLOOKUP(cluster_id, owner_id, cost)` |
| Clusters without applications | `cluster_snapshot`, filter PROD, `category` = NO_APPLICATIONS; add the cost the same way |
| BIA (annex) | `cluster_snapshot` with the cost column; rows `bia`, columns `in_export` |

### 2. Microservices and HPA (`ms_and_hpa_summary.xlsx`)

Inputs: the `aa-k8s-snapshot` CSVs, `aks_autoscaler_monthly.csv`, `aks_oomkilled_monthly.csv` and
`microservice_history.csv`. Add `category` to `deployment_snapshot`, `hpa_snapshot`,
`aks_autoscaler_monthly` and `aks_oomkilled_monthly` with the category join.

Standard filter: `environment` = PROD, `category` = APPLICATION and, where the CSV has it,
`microservice` = YES. Build one pivot with it, then copy and paste it for the next figure and change
only the rows; a slicer on `environment` connected to every pivot (*Report connections*) changes them all.

| Figure | Pivot or formula |
|---|---|
| Clusters by category | `cluster_snapshot`, filter PROD; rows `category`; count |
| Microservices per layer (pattern 1) | `deployment_snapshot`, standard filter; rows `layer`; count |
| Repeated capabilities | `capability_domains` as is (already filtered). A capability is a keyword in the deployment name, so the counts are a floor |
| Replicas now (pattern 4) | `hpa_snapshot`, standard filter; rows `replicas_now` grouped by hand (select the labels > right click > *Group*): 0 (possible zombie), 1, 2–3, 4–5, 6–10, 11–20, 21–50, more than 50; count of `hpa` and the same count as % of grand total. One HPA is one deployment |
| Time without scaling | Same filter; rows `last_scale_bucket`, columns `no_scale_cause`; count |
| HPA behaviour | Same filter; rows `behaviour_ci`; count as % of total. Headline "does not scale" = `NEVER_SCALES` + `FIXED`, shown split |
| Inflated minimums | Same filter plus `behaviour_ci` = NEVER_SCALES; rows `min_above_need` or `min_above_need_beyond_ha` |
| How far HPAs reach | Power Query column `office_max_vs_hpa_max` = `ci_replicas_max_office / hpa_max`; pivot rows grouped *Start* 0, *End* 0.99, *By* 0.25; `>0.99` = reached its maximum. An upper bound: Container Insights counts distinct pods per hour, so rollouts and restarts add the replaced pods. `office_max_reached_hpa_max` (YES/NO) is the same test; `trend_reached_max` is the export's lower bound. Exact figure: `scripts/grafana.promql`, section 5 |
| Min, max and average replicas | Same filter; values `hpa_min`, `hpa_max`, `ci_replicas_avg_office`, `ci_replicas_max_office` summarised by average, max and min |
| Fits Container Apps by size | `deployment_snapshot`, standard filter; rows `fits_aca_by_request` (request per replica <= 4 vCPU and <= 8 GiB; empty when the export has no request). For the `NO` rows, max of `cpu_request_vcpu` and `mem_request_gib` shows why |
| Clusters short of capacity | `aks_autoscaler_monthly` with the category lookup, filter PROD, `category` = APPLICATION and one `month`; `hours_with_unschedulable` in the filter area without 0; count of `cluster` and sum of the hours. Share of the time: calculated field `unschedulable_share = hours_with_unschedulable / hours_with_data` (the deck figure). `unschedulable_pods_max` is a peak: summarise it by **max**, never sum |
| OOM kills | `aks_oomkilled_monthly`, standard filter (`microservice` = YES drops platform namespaces and agents such as `orca-sensor`, `otel-*`, `thanos-*`, `logstash`); rows `cluster` or `workload`; sum `oom_kills` |
| Microservices over time | `microservice_history`, filter PROD and leave out the infrastructure app codes (`COAA`, `VSDP`, `HTFC`, `INCT`, `O11Y`, `ELKP`, `APIG`, `GFED`); rows `date`; sum `microservices`; line chart. With `cod_app` in rows and `date` in columns it shows which applications grew |

### 3. APIM traffic (`apim_traffic_summary.xlsx`)

Inputs: `apim_api_monthly.csv`, `apim_hourly_profile.csv`, `apim_latency_monthly.csv`, and
`application_snapshot.csv` for the per-application figures.

Standard filter for the per-API figures: `month` = the measured month and `apim_layer` = BS. Without the
layer filter a request is counted once per gateway it passes (UX, BS, Core).

| Figure | Pivot or formula |
|---|---|
| Requests per API, office and off-hours | `apim_api_monthly`, standard filter; rows `api_id`; sum `requests`, plus calculated fields `office_pct = SI(requests=0, 0, office_requests/requests)` and `offhours_pct` the same with `offhours_requests`, formatted as %. Office = weekday 8:00–19:59 Lima |
| Off-hours share of the estate | The grand total row of `offhours_pct` in the same pivot |
| Peak TPS per API | Power Query column `peak_tps = [peak_5min_requests] / 300` (the busiest 5 minutes of the month, as requests per second). Pivot: standard filter; rows `api_id`; **max** of `peak_tps`, sorted descending. *Value filters > Greater than* 100 keeps the APIs above the ADT_019 threshold; an empty result is a finding too. The 5-minute average dilutes shorter bursts |
| Load per application (ADT_019 100 TPS rule) | `application_snapshot` (PROD application clusters only, one traffic layer per application already chosen). Power Query: types with the `en-US` locale for `busiest_hour_tps`, `peak_tps_upper_bound`, `weekday_hours_with_traffic`, `microservices`, `microservices_projected`; columns `expected_ms = List.Max({[microservices], if [microservices_projected] = null then 0 else [microservices_projected]})` (PROD, or PROD plus DESA when larger), `over_100_tps = if [traffic_status] <> "MEASURED" then "NO_TRAFFIC" else if [busiest_hour_tps] > 100 then "YES" else "NO"` and `constant_and_large = if [weekday_hours_with_traffic] <> null and [weekday_hours_with_traffic] > 10 and [expected_ms] > 15 then "YES" else "NO"`. `traffic_layer` in the filter area shows which gateway measured each application. Pivot 1: rows `over_100_tps`; count of `cod_app` (YES + NO = applications with traffic). Pivot 2: rows `constant_and_large`; count. Pivot 3: rows `cod_app`; max of `busiest_hour_tps` and `peak_tps_upper_bound`, sorted descending. `busiest_hour_tps` adds the application's APIs hour by hour (average weekday or weekend profile) and takes the busiest hour; `peak_tps_upper_bound` adds each API's own 5-minute peak, which happen at different times, so it overstates the load |
| Backend latency | `apim_latency_monthly`, `month` = the measured month and `apim_layer` = BS; distribution of `backend_duration_avg_ms` (for example < 100 ms, 100–500 ms, > 500 ms) |
| Traffic class per API (optional) | Already built in the `Traffic_profile` sheet of `aks_aca_comparison.xlsx`; technical deck only until an architect validates the classes. Office vs off-hours above is the ADT_019 model (warm weekdays 8–20, scale to zero otherwise). To rebuild it: `apim_hourly_profile`: helper `week = avg_requests_per_hour x (5 on weekday, 2 on weekend)` and helper window (office: weekday 8–19; extended: any day 6–21; rest). Pivot `api_id` x window, sum `week`. Office >= 90% → OFICINA; office + extended >= 90% → EXTENDIDO; else 24/7 |

The class is per API. Mapping APIs to clusters needs the API → deployment mapping, which is still pending.

**Traffic per application across every layer (rebuilds `traffic_layer` and `requests_month`).** Two joins and an
append; on the October test run it matched `application_snapshot` for all 117 applications.

1. Load `apim_apis.csv` as a connection-only query, keeping `apim`, `api_id`, `apim_environment`,
   `backend_platform` and `backend_cod_app`.
2. In `apim_api_monthly`: *Home > Merge Queries* with `apim_apis` on **`apim` and `api_id`** (select both columns,
   in the same order, on each side), left outer; expand `apim_environment`, `backend_platform`, `backend_cod_app`.
   Keep the rows with `backend_platform` = AKS and `apim_environment` = PROD (an API in front of App Service or
   OpenShift is not AKS traffic). Use `backend_cod_app` as the application.
3. Load `entry_points_monthly.csv` (Application Gateways and Event Hubs / Service Bus) and add a column
   `backend_cod_app` = `cod_app`. Event Hubs count messages, not requests.
4. *Home > Append Queries as New*: the two queries. Keep `backend_cod_app`, `apim_layer`, `month`, `requests`.
5. Pivot: filter `month`; rows `backend_cod_app`; columns `apim_layer`; sum `requests`.
6. Next to the pivot, per row: the layer is BS when its BS cell is above 0, else the column with the largest
   value; the requests are that cell. For example `=SI(B5>0; B5; MAX(C5:H5))` and
   `=SI(B5>0; "BS"; INDICE($C$4:$H$4; COINCIDIR(MAX(C5:H5); C5:H5; 0)))`, with BS in column B.
7. Check: `BUSCARX` of each application in `application_snapshot` (`traffic_layer`, `requests_month`) gives the
   same values. Never add the columns of one row: the same call passes through several gateways.

**Every TPS in one pivot per application.** `apim_api_monthly` already has `cod_app`. Use the standard filter
(BS only); an application whose `traffic_layer` is not BS has its figures in `application_snapshot` only. Two figures are
per application, not per API, so they enter the per-API rows through a lookup and are summarised by **max**
(each API row of an application carries the same value; a sum would multiply it by the number of APIs):

- Load `application_snapshot.csv` as a connection-only query, with `busiest_hour_tps` typed as a decimal
  with the `en-US` locale.
- In the `apim_api_monthly` query (*Home > Advanced Editor*), add before `in`, replacing `#"Tipo cambiado"`
  with the last step:

  ```text
      Busiest = Table.Distinct(Table.SelectColumns(application_snapshot, {"cod_app", "busiest_hour_tps"}), {"cod_app"}),
      BusiestByApp = Record.FromList(Busiest[busiest_hour_tps], Busiest[cod_app]),
      WithBusiest = Table.AddColumn(#"Tipo cambiado", "busiest_hour_tps",
          each if [cod_app] = null or [cod_app] = "" then null
               else Record.FieldOrDefault(BusiestByApp, [cod_app], null), type number)
  in
      WithBusiest
  ```

| Value in the per-application pivot | Summary | Meaning |
|---|---|---|
| `office_tps` = `office_requests / 950400` (calculated field; 22 working days × 12 h × 3600 s in October 2026) | Sum | Average TPS over office hours |
| `offhours_tps` = `offhours_requests / 1728000` ((744 − 264) h × 3600 s) | Sum | Average TPS outside office hours |
| `busiest_hour_tps` (lookup above) | **Max** | Typical busiest hour of the application; the ADT_019 100 TPS rule |
| `peak_tps` | **Max** | Busiest 5 minutes of the application's busiest API |

Always `office_tps` <= `busiest_hour_tps` <= the 5-minute peak of the whole application. An application in
APIM whose clusters are not APPLICATION has no `busiest_hour_tps`.

### 4. Container Apps price per microservice (Azure pricing calculator)

The workbook's `Break_even` sheet shows "ACA cost per microservice per month (ADT schedule)". To back that
figure with Microsoft's own calculator:

1. Open the [Azure pricing calculator](https://azure.microsoft.com/pricing/calculator/) and add
   **Container Apps**.
2. Region **East US 2** (153 of the 158 PROD clusters and every PROD Container Apps environment; Central US
   has the same Container Apps prices). Plan **Consumption**.
3. The calculator does not ask for hours. It derives the active time as
   `requests × execution time per request / concurrent requests`, as if the replica only ran while serving
   requests. To reproduce the ADT_019 warm replica (on for the whole office schedule), enter values that give
   the office seconds of the month (22 working days × 12 h × 3600 = 950,400 s in October 2026):

   | Field | Value |
   |---|---|
   | Requests per month | 5,000,000 |
   | Execution time per request | 190 ms (950,400 s / 5 M requests) |
   | Concurrent requests per container app | 1 |
   | vCPU / memory | 0.5 / 1 GiB |

   The 190 ms is not the real latency; it only turns the warm replica into the calculator's terms.
4. The calculator subtracts the monthly free grant of the subscription (180,000 vCPU-s, 360,000 GiB-s,
   2 M requests). On 2026-11-17 it gave **10.06 USD**. Add the grant back to get the deck figure:
   180,000 × 0.000024 + 360,000 × 0.000003 + 2 × 0.40 = 6.20, so 10.06 + 6.20 = **16.26 USD**.
5. Save the estimate (*Export* or a shared link) next to the deck as the evidence.

The deck leaves the grant out: it is one per subscription and month, shared by every app in it, so it does not
belong to a single microservice. `grant_apply` = 1 in `Assumptions` applies it once to an application's total.

The same figure by hand, with the list prices of the
[Container Apps pricing page](https://azure.microsoft.com/pricing/details/container-apps/) (East US 2, also
in `results/aa-prices/prices.csv`):

| Item | Formula | USD per month |
|---|---|---|
| Office hours, warm replica (October 2026) | 950,400 s × (0.5 vCPU × 0.000024 + 1 GiB × 0.000003) | 14.26 |
| Requests | 5 × 0.40 per million | 2.00 |
| **Total, as in the calculator without the grant** | | **16.26** |
| Outside office hours (wake-ups and cooldown, measured in APIM; workbook only) | 48,000 s × 0.000015 | 0.72 |

- The workbook uses 21.75 working days (261 a year / 12), so its office hours give 939,600 s and 14.09 USD;
  with the off-hours line its total is about 16.81 USD. The 1% gap with October is the calendar.
- 5 M requests in a month are about 4.7 TPS on average in office hours (90% of them, measured share) and about
  14 TPS in the busiest hour: one replica is enough, so the cost is the warm replica, not the traffic.
- Every second is at the active rate. Azure bills a warm replica at the idle rate only while it serves no
  requests, uses under 0.01 vCPU and receives under 1,000 bytes/s
  ([Container Apps billing](https://learn.microsoft.com/azure/container-apps/billing#consumption-plan)),
  which a Java replica with agents and probes rarely meets.
- Outside office hours the minimum is 0 replicas, so a running replica is above the minimum and always billed
  as active.
- Pure serverless, for reference (no warm replica): 5 M requests × the measured latency
  (`backend_duration_avg_ms` in `apim_latency_monthly`, for example 100 ms) / 10 concurrent requests (the
  default HTTP scale rule) = 50,000 s, about 0.75 + 2.00 = 2.75 USD. Nearly all the ADT_019 cost is keeping the
  replica warm in office hours.
- A microservice that needs more than one replica in office hours adds 14.26 per extra replica. The workbook
  computes replicas from the requests and `rps_rep` (requests per second one replica serves).
- The `Assumptions` sheet keeps every input editable (`req_month`, `office_active_share`, `grant_apply`,
  `cpu_rep`, `mem_rep`).

### 5. Container Apps cost per application (`aca_estimate_by_app.csv`)

`aa_deck_exec.py` writes this CSV next to the corporate deck: one row per PROD application with every input and
step of its monthly Container Apps cost, so the saving slide can be rebuilt by hand or in the
[Azure pricing calculator](https://azure.microsoft.com/pricing/calculator/) (Container Apps, Consumption, East US 2).

| Step | Column | Formula |
|---|---|---|
| Replicas | `replicas_office` | Per microservice: office-hours average in Container Insights (else the HPA minimum, at least 1), added up |
| Size of each replica | (inside `vcpu_seconds`) | Smallest Consumption size that covers the request: 0.25 vCPU steps, 2 GiB per vCPU |
| Time warm | `office_seconds` | Monday to Friday 8:00–20:00 of the month (ADT_019); zero replicas otherwise |
| vCPU-seconds | `vcpu_seconds` | replicas × vCPU × office seconds |
| GiB-seconds | `gib_seconds` | replicas × GiB × office seconds |
| Requests | `requests_with_buffer` | requests of the month × 1.25 (headroom so the saving is not overstated) |
| Compute | `compute_usd`, `memory_usd` | vCPU-s × 0.000024 + GiB-s × 0.000003 (active rate, list price) |
| Requests | `requests_usd` | requests / 1,000,000 × 0.40 |
| Environment | `environment_usd` | 1 environment × 730 h × (0.10 management + 0.10 private endpoint) = 146 |
| List total | `aca_list_usd` | sum of the four lines above |
| With discount | `aca_with_discount_usd` | list × 0.70 (the bank's 30%; Container Apps in consumption has no reservations) |
| In soles | `aca_with_discount_pen_igv` | × 3.41333 × 1.18 |

`aks_cost_usd` is what the application's clusters cost in the month (amortized, reservations included).

- In the calculator, enter `vcpu_seconds` and `gib_seconds` as the active usage, `requests_with_buffer` as the
  requests, and add the environment hours as a separate line; it subtracts the free grant, which the CSV leaves out.
- Most of the cost is keeping replicas warm and the environment: on the October test run the requests were
  about 0.1% of the candidates' Container Apps cost and the environments about 60%. Sharing one environment among
  several applications lowers that line (`ENVIRONMENTS_PER_APP` in `aa_deck_exec.py`).

### 6. Corporate deck figures (`aks_aca_corporate.pptx`)

Every figure of the corporate deck, so it can be explained and rebuilt in Excel. Money: US$ × 3.41333 × 1.18
(exchange rate and IGV); the CSVs keep US$. Month: the measured month (`month` of `cost_by_resource`, the
latest with AKS cost).

Inputs: `application_snapshot.csv` (one row per PROD application: the base of slides 1 and 3 to 7),
`cluster_snapshot.csv`, `hpa_snapshot.csv`, `aks_usage_monthly.csv`, `cost_by_resource.csv`, the inventory CSVs
and `aca_estimate_by_app.csv` (written next to the deck).

Power Query columns on `application_snapshot` (types with the `en-US` locale), used by several slides:

```text
expected_ms   = List.Max({[microservices], if [microservices_projected] = null then 0 else [microservices_projected]})
c1_traffic    = [busiest_hour_tps] <= 100
c2_pieces     = not ([weekday_hours_with_traffic] > 10 and [expected_ms] > 15)
c3_night      = [offhours_active_share] <= 0.5
c4_size       = [microservices_over_aca_size] = 0
group         = if [adt_runtime] <> "INSUFFICIENT_DATA" then [adt_runtime]
                else if [no_traffic_reason] = "IN_DEVELOPMENT" then "IN_DEVELOPMENT"
                else if [no_traffic_reason] = "IN_USE_NOT_MEASURED" then "IN_USE" else "NOT_IN_USE"
```

(an empty `offhours_active_share` counts as 0, as in the script.)

| Slide | Figure | Pivot or formula |
|---|---|---|
| 1 | Applications and microservices on AKS | `application_snapshot`: count of `cod_app`; sum of `microservices` |
| 2 | Runtimes in PROD (cards) | Inventory CSVs, filter `environment` = PROD and `existed_at_measurement` other than NO: `aks_clusters` (count `id`), `aca_apps`, `aca_environments`, `app_services` with `type` = WEB, `app_service_plans` with `runtime` = APP_SERVICE. Functions are left out for now (`SHOW_FUNCTIONS` in `aa_deck_exec.py`) |
| 2 | Cost per runtime | `cost_by_resource`, measured month, `environment` = PROD; rows `owner_type`; sum `cost_usd`, in soles |
| 2 | Clusters by category (table) | `cluster_snapshot`, PROD, `in_inventory` = YES, cluster existed in the month; rows `category`; count and cost (lookup of `owner_id` in `cost_by_resource`) |
| 3 | CPU and memory used | `aks_usage_monthly`, measured month, APPLICATION clusters only (category lookup): sum `cpu_used_office_avg` of the usage rows / sum `cpu_allocatable_avg` of the `ALLOCATABLE` rows; the same with `mem_used_office_avg_gib` and `mem_allocatable_avg_gib` |
| 3 | Autoscalers without changing size | `hpa_snapshot`, PROD, `microservice` = YES, APPLICATION clusters, `last_scale_bucket` other than NO_DATA: (`3_TO_12_MONTHS` + `OVER_1_YEAR`) / count |
| 3 | Applications with 15 microservices or fewer | `application_snapshot`: count with `microservices` between 1 and 15, out of the total |
| 3 | Minimum AKS vs one Container Apps microservice | From the Azure pricing calculator (East US 2, the figures in `CALCULATOR_AKS_USD` and `CALCULATOR_ACA_USD` of `aa_deck_exec.py`). AKS: Azure Kubernetes Service, tier Standard, 4 × D8s v4 Linux, 730 h, 3-year reservation = 426.11 US$ of nodes, plus 4 managed OS disks P10 (71.68) and the Standard tier (73.00); disks and tier × 0.70 (the bank's 30%), nodes as they are: 527.39 US$ = S/ 2,124. Container Apps: the calculator setup of section 4, 16.26 US$ × 0.70 = S/ 46. Cross-check (log only): the median October bill of the 11 PROD clusters with exactly 4 D8s v4 nodes all month was 514.75 US$ (S/ 2,073), within 3% |
| 4 | The four conditions | Text only. Each one is a column above (`c1` to `c4`); it is the ADT_019 tree read the other way round. Condition 2 is shown as "Hasta 15 servicios": `c2_pieces` keeps the ADT_019 rule (more than 15 services and more than 10 hours a day), and on the October data no application with more than 15 services is used 10 hours or less, so both give the same count; when one does, slide 5 names it |
| 5 | Funnel | Rows with `busiest_hour_tps` not empty = "Con tráfico medido". Then filter `c1_traffic` = TRUE, then also `c2_pieces`, `c3_night`, `c4_size`: the count after each filter is one bar. The last bar equals `adt_runtime` in (ACA, ACA_WITH_REVIEW) |
| 6 | Result per group | Pivot rows `group` (leave `IN_DEVELOPMENT` out of the chart; it is the second bullet); count of `cod_app` and sum of `cost_month_usd`, in soles |
| 6 | Single microservice | Count of `possible_monolith` = YES |
| 7 | Saving | `aca_estimate_by_app`, filter `adt_runtime` in (ACA, ACA_WITH_REVIEW). With discount: sum `aks_cost_pen_igv` − sum `aca_with_discount_pen_igv` per month; × 12 per year. List price: sum `aks_cost_usd` / measured discount − sum `aca_list_usd`, in soles |

- The measured discount is what PROD node VMs cost in the month (amortized, reservations included) divided by
  their node-hours at list price; the script logs it and prints it on slides 3 and 7. To check it in the portal:
  Cost analysis, amortized cost, group by *Pricing model* and filter Virtual Machines (Reservation vs OnDemand).
- The saving is a month repeated twelve times: the same replicas, prices and applications. It includes the
  applications "con revisión", so the firm saving is the `ACA` rows alone (filter `adt_runtime` = ACA).
- To project with other inputs, change the columns of `aca_estimate_by_app` and recompute: more replicas
  (`replicas_office`), more environments (146 US$ each per month), more requests (`requests_with_buffer`).

## Metrics not collected today

Ready to add if a question needs them. Each one is a change to the script and a rerun.

| Metric | Source | Answers |
|---|---|---|
| `cluster_autoscaler_cluster_safe_to_autoscale` | AKS platform metrics | Is the autoscaler able to act? |
| `kube_pod_status_ready`, `kube_pod_status_phase` | AKS platform metrics | Pods not ready or pending |
| `node_disk_usage_percentage`, `node_network_in_bytes` / `_out_bytes` | AKS platform metrics | Disk and network use per node |
| `KubeEvents` (reason `FailedScheduling`, `BackOff`) | Container Insights | Why pods were not scheduled or keep restarting |
| `cpuLimitNanoCores`, `memoryLimitBytes` | Container Insights `Perf` | Limits measured, not only from the export |
| `KubeNodeInventory` | Container Insights | Node labels, kubelet version, node conditions |
| `Capacity` | APIM platform metrics | APIM load |

## Documentation

- AKS platform metrics:
  https://learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-containerservice-managedclusters-metrics
- APIM platform metrics:
  https://learn.microsoft.com/azure/azure-monitor/reference/supported-metrics/microsoft-apimanagement-service-metrics
- Platform metrics retention (93 days): https://learn.microsoft.com/azure/azure-monitor/essentials/data-platform-metrics
- Node resource reservations (allocatable vs VM size): https://learn.microsoft.com/azure/aks/node-resource-reservations
- Container Insights tables: https://learn.microsoft.com/azure/azure-monitor/reference/tables/kubepodinventory and
  https://learn.microsoft.com/azure/azure-monitor/reference/tables/perf
- Container Insights queries: https://learn.microsoft.com/azure/azure-monitor/containers/container-insights-log-query
- Resource Graph: https://learn.microsoft.com/azure/governance/resource-graph/overview
- AKS resource properties: https://learn.microsoft.com/azure/templates/microsoft.containerservice/managedclusters
- Cost Management query API: https://learn.microsoft.com/rest/api/cost-management/query/usage
- Amortized cost: https://learn.microsoft.com/azure/cost-management-billing/reservations/view-amortized-costs
- Retail Prices API: https://learn.microsoft.com/rest/api/cost-management/retail-prices/azure-retail-prices
- Kubernetes HPA: https://kubernetes.io/docs/tasks/run-application/horizontal-pod-autoscale/
- Container Apps scaling: https://learn.microsoft.com/azure/container-apps/scale-app
