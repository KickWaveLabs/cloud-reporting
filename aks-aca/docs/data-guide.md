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
| `aa_k8s_snapshot.py` | Which microservices exist, with which HPA and BIA, and how did they behave in the month? | No Azure call: the platform team's exports joined with Container Insights | `aa-k8s-snapshot` |

## The CSVs

### aa-inventory

| CSV | One row is | Main columns |
|---|---|---|
| `aks_clusters.csv` | A cluster | `name`, `environment` (from the tag, else the name suffix), `kubernetes_version`, `system_nodes` and `user_nodes` (nodes at the time of the run, not the minimum), `vm_sizes`, `upgrade_channel` (`none` = manual upgrades), `service_mesh`, `gitops_flux`, `container_insights`, plus one column per configuration property |
| `aks_node_pools.csv` | A node pool | `cluster`, `mode` (System/User), `vm_size`, `count` (now), `autoscaling`, `min_count`, `max_count`, `zones`, `max_pods` |
| `aca_environments.csv`, `aca_apps.csv` | A Container Apps environment or app | plan, workload profiles, min and max replicas, CPU and memory per replica |
| `apim_services.csv` | An APIM instance | `sku`, `capacity`, `custom_domains` (tells BS from Core) |
| `apim_apis.csv` | An API | `api_id`, `path`, `backend_host`, `backend_cod_app` (app code of the AKS ingress behind it) |
| `app_service_plans.csv` | An App Service plan | `sku`, `tier`, `workers`, `os`, `sites` |
| `app_services.csv` | A web app or function app | `type` (`WEB`/`FUNCTION`), `container` (`YES` = Docker image), `state`, `plan`, `plan_sku` |

The configured node minimum of a cluster is the sum over its pools of `min_count` when `autoscaling`
is true, else `count`.

### aa-metrics

| CSV | One row is | Main columns |
|---|---|---|
| `aks_usage_monthly.csv` | Cluster x pool x month (usage rows), or cluster x month with `mode = ALLOCATABLE` (capacity row) | Usage rows: `nodes_avg`, `cpu_used_office_avg`, `cpu_used_offhours_avg`, `cpu_used_p95`, `cpu_used_max` (vCPU) and the same for memory (GiB). Capacity rows: `cpu_allocatable_avg`, `mem_allocatable_avg_gib` |
| `apim_api_monthly.csv` | APIM x API x month | `requests`, `office_requests`, `offhours_requests`, `peak_5min_requests` (divide by 300 for requests per second), `peak_hour_requests`, `offhours_wakeups` |
| `apim_hourly_profile.csv` | APIM x API x day type (`weekday`/`weekend`) x Lima hour | `avg_requests_per_hour` |

- Usage comes from `node_cpu_usage_millicores` and `node_memory_working_set_bytes`, summed over the
  nodes of each pool, hour by hour.
- Allocatable comes from `kube_node_status_allocatable_cpu_cores` and `_memory_bytes`. It is the
  cluster total for the nodes running in each hour, system and user pools together.
- For CPU, allocatable is almost the full VM: 2 x 8 vCPU nodes give 16.0. For memory AKS reserves a
  part: 2 x 32 GiB nodes give about 61 GiB.
- Use % = sum of usage / sum of allocatable. Both rows types have `environment`, so filter PROD and
  divide the two sums.

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
- `owner_type` is `AKS` or `ACA`.
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
| `deployment_snapshot.csv` | A deployment | `microservice` (`YES` = an application deployment), `platform_component`, `layer`, `bia`, `capabilities`, `cpu_request_vcpu`, `mem_request_gib`, `has_hpa`, `ci_replicas_avg_office`, `ci_cpu_used_per_replica`, `ci_request_minus_export_vcpu`, `node_vcpu`, `node_mem_gib`, `pod_too_big_for_node`, `fits_aca_by_request` (request within 4 vCPU / 8 GiB) |
| `hpa_snapshot.csv` | An HPA | `microservice`, `hpa_min`, `hpa_max`, `replicas_now` (export day), `last_scale_time`, `days_since_last_scale`, `last_scale_bucket`, `ci_replicas_avg_office`, `ci_replicas_max_office`, `office_max_reached_hpa_max`, `behaviour_ci`, `no_scale_cause`, `replicas_needed_by_cpu`, `min_above_need` |
| `capability_domains.csv` | A capability keyword | `namespaces_prod` (domains), `deployments_prod`, both in APPLICATION clusters |
| `application_snapshot.csv` | A PROD application (app code of APPLICATION clusters) | `clusters`, `cost_month_usd`, `bia`, `microservices`, `hpas_not_scaling_share`, `apis_with_traffic`, `requests_month`, `office_share`, `peak_tps_upper_bound`, `busiest_hour_tps`, `weekday_hours_with_traffic`, `batch_jobs`, `statefulsets`, `microservices_over_aca_size`, `possible_monolith`, `adt_runtime`, `adt_reason` |

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
| `AKS` | Busiest hour above 100 TPS; or traffic more than 10 weekday hours and more than 15 microservices; or a microservice requests more than 4 vCPU / 8 GiB |
| `ACA_WITH_REVIEW` | Within the thresholds but with StatefulSets, batch jobs, BIA ALTA or MUY ALTA (consult architecture governance), or no APIM traffic mapped |
| `ACA` | Within every threshold |

- Traffic: each API maps to an app code through the AKS ingress host of its backend. `busiest_hour_tps` adds
  the application's APIs hour by hour (average weekday or weekend profile of the month); `peak_tps_upper_bound`
  adds each API's 5-minute peak, so it overstates the real peak.
- An hour has traffic when it carries at least 10% of the busiest hour; `weekday_hours_with_traffic` counts them.
- `possible_monolith` = a single microservice. The tree sends internal monoliths to App Service, but a single
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

### aa-signals

| CSV | One row is | Main columns |
|---|---|---|
| `aks_autoscaler_monthly.csv` | Cluster x month | `hours_with_data` (0 = no autoscaler, the metric is not emitted), `hours_with_unschedulable`, `office_hours_with_unschedulable`, `unschedulable_pods_max` (highest hourly average) |
| `aks_oomkilled_monthly.csv` | Cluster x namespace x workload x month | `oom_kills`, `pods` affected, `system_namespace` |
| `aks_stack_memory.csv` | Cluster x namespace x workload (last `--stack-days`, office hours) | `stack` (Quarkus, Spring Boot, Java, .NET, Node.js, Python, Unidentified), `replicas_avg`, `app_mem_gib_per_replica`, `sidecar_mem_gib_per_replica` (Istio; empty when the pod has no sidecar) |
| `apim_latency_monthly.csv` | APIM x API x month | `backend_duration_avg_ms`, `backend_duration_max_ms` (APIM to the backend and back), `duration_avg_ms`, `duration_max_ms` (whole gateway request) |

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

## Building your own workbooks

Open each CSV with *Data > From Text/CSV*, comma delimiter. Opening it with a double click splits the
columns wrong in a Spanish Excel. Three workbooks cover the deck.

### 1. Cost and capacity

Inputs: `aks_clusters.csv`, `aks_node_pools.csv`, `aks_usage_monthly.csv`, `cost_by_resource.csv`.

| Figure | Pivot or formula |
|---|---|
| PROD clusters | `aks_clusters`, filter `environment` = PROD, count `id` (two clusters can share a name) |
| Monthly cost | `cost_by_resource`, filter `owner_type` = AKS, `environment` = PROD; rows `month`; sum `cost_usd` |
| Cost by type | Same pivot with `meter_category` in rows |
| Cost per cluster | Same pivot with `owner_id` in rows |
| Clusters with 4 nodes or fewer now | `aks_clusters`: helper column `system_nodes + user_nodes`, count <= 4 |
| Clusters under the 4-node minimum | `aks_node_pools`: helper column `IF(autoscaling="true", min_count, count)`; pivot by `cluster`, sum it, count < 4 |
| CPU and memory use in office hours | `aks_usage_monthly`, filter PROD; sum of `cpu_used_office_avg` / sum of `cpu_allocatable_avg` (same for memory, p95 and max) |
| Use per cluster | Same pivot with `cluster` in rows |

### 2. Microservices and HPA

Inputs: the four `aa-k8s-snapshot` CSVs, plus the cost per cluster from workbook 1.

| Figure | Pivot or formula |
|---|---|
| Clusters by category | `cluster_snapshot`, filter PROD; rows `category`; count. Every figure below keeps only `APPLICATION` clusters (join by `cluster_id`) |
| Microservices per layer (pattern 1) | `deployment_snapshot`, filter PROD and `microservice` = YES; rows `layer`; count |
| Repeated capabilities | `capability_domains` as is |
| Clusters per microservices bucket (pattern 2) | `cluster_snapshot`, filter PROD, `category` = APPLICATION; helper bucket column on `microservices`; add the cost with `XLOOKUP(cluster_id, owner_id, cost)` |
| Clusters without applications | `cluster_snapshot`, filter PROD, `category` = NO_APPLICATIONS; add the cost the same way |
| Replicas now (pattern 4) | `hpa_snapshot`, filter PROD and `microservice` = YES; helper bucket column on `replicas_now` |
| Time without scaling | Same filter; rows `last_scale_bucket`, columns `no_scale_cause`; count |
| HPA behaviour | Same filter; rows `behaviour_ci`; count as % of total |
| Inflated minimums | Same filter plus `behaviour_ci` = NEVER_SCALES; rows `min_above_need` or `min_above_need_beyond_ha` |
| BIA (annex) | `cluster_snapshot` with the cost column; rows `bia`, columns `in_export` |
| Fits Container Apps by size | `deployment_snapshot`; rows `fits_aca_by_request` |
| Clusters short of capacity | `aks_autoscaler_monthly`, filter PROD; count clusters with `hours_with_unschedulable` > 0; sum the hours |
| OOM kills | `aks_oomkilled_monthly`, filter PROD, `system_namespace` = NO and outside platform namespaces (`orcasecurity` included: its sensor is a security agent, not an app); rows `cluster` or `workload`; sum `oom_kills` |

### 3. APIM traffic

Inputs: `apim_api_monthly.csv`, `apim_hourly_profile.csv`, `apim_apis.csv`.

| Figure | Pivot or formula |
|---|---|
| Requests per API and month | `apim_api_monthly`; rows `api_id`, columns `month`; sum `requests` |
| APIs above 100 tps | Same CSV, last full month, count `peak_5min_requests` > 30,000 |
| Backend latency | `apim_latency_monthly`, last full month; distribution of `backend_duration_avg_ms` (for example < 100 ms, 100–500 ms, > 500 ms) |
| Off-hours share | Sum of `offhours_requests` / sum of `requests` |
| Traffic class per API | `apim_hourly_profile`: helper `week = avg_requests_per_hour x (5 on weekday, 2 on weekend)` and helper window (office: weekday 8–19; extended: any day 6–21; rest). Pivot `api_id` x window, sum `week`. Office >= 90% → OFICINA; office + extended >= 90% → EXTENDIDO; else 24/7 |

The class is per API. Mapping APIs to clusters needs the API → deployment mapping, which is still pending.

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
