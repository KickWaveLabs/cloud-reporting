# Data guide: what each CSV holds and how to build your own workbooks

Every figure in the deck comes from a CSV in `aks-aca/results/`. This guide says which script writes
each CSV, what one row is, what the main columns mean, and how to rebuild the deck figures with pivot
tables, without `aa_workbook.py`.

Office hours everywhere: Monday to Friday, 8:00 to 19:59 Lima time. Holidays count as working days.

## The scripts at a glance

| Script | Question it answers | Azure source | Output folder |
|---|---|---|---|
| `aa_inventory.py` | Which clusters, node pools, Container Apps and APIM APIs exist, and how are they configured? | Resource Graph; ARM GET for the APIM APIs | `aa-inventory` |
| `aa_metrics.py` | How much CPU and memory do the nodes use? How many requests reach each API, and when? | Azure Monitor platform metrics (cluster and APIM resources) | `aa-metrics` |
| `aa_container_insights.py` | Which workloads run, with how many replicas, and how much does each replica use and request? | Log Analytics (Container Insights tables) | `aa-container-insights` |
| `aa_cost.py` | How much did each cluster cost each month? | Cost Management query API, `AmortizedCost` | `aa-cost` |
| `aa_prices.py` | What are the list prices of VMs and Container Apps? | Azure Retail Prices API | `aa-prices` |
| `aa_k8s_snapshot.py` | Which microservices exist, with which HPA and BIA, and how did they behave in 3 months? | No Azure call: the platform team's exports joined with Container Insights | `aa-k8s-snapshot` |

## The CSVs

### aa-inventory

| CSV | One row is | Main columns |
|---|---|---|
| `aks_clusters.csv` | A cluster | `name`, `environment` (from the tag, else the name suffix), `kubernetes_version`, `system_nodes` and `user_nodes` (nodes at the time of the run, not the minimum), `vm_sizes`, `upgrade_channel` (`none` = manual upgrades), `service_mesh`, `gitops_flux`, `container_insights`, plus one column per configuration property |
| `aks_node_pools.csv` | A node pool | `cluster`, `mode` (System/User), `vm_size`, `count` (now), `autoscaling`, `min_count`, `max_count`, `zones`, `max_pods` |
| `aca_environments.csv`, `aca_apps.csv` | A Container Apps environment or app | plan, workload profiles, min and max replicas, CPU and memory per replica |
| `apim_services.csv` | An APIM instance | `sku`, `capacity`, `custom_domains` (tells BS from Core) |
| `apim_apis.csv` | An API | `api_id`, `path`, `backend_host` |

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

| CSV | One row is | Main columns |
|---|---|---|
| `deployment_snapshot.csv` | A deployment | `platform_component` (`SÍ` for platform namespaces), `layer` (first word of the name), `bia`, `capabilities`, `cpu_request_vcpu`, `mem_request_gib`, `has_hpa`, `ci_replicas_avg_office`, `ci_cpu_used_per_replica`, `fits_aca_by_request` (request within 4 vCPU / 8 GiB) |
| `hpa_snapshot.csv` | An HPA | `replicas_now` (export day), `hpa_min`, `hpa_max`, `ci_replicas_avg_office`, `behaviour_ci`, `replicas_needed_by_cpu`, `min_above_need`, `min_above_need_beyond_ha` |
| `cluster_snapshot.csv` | A cluster | `in_inventory`, `in_export` (`NO` = not in the export, so its 0 microservices mean "unknown"), `microservices`, `namespaces`, `hpas`, `bia` (highest among its microservices), `statefulsets`, `aca_technical_fit` |
| `capability_domains.csv` | A capability keyword | `namespaces_prod` (domains), `deployments_prod` |

`behaviour_ci` compares the 3-month office average of replicas with the HPA limits:

| Code | Rule |
|---|---|
| `NUNCA_ESCALA` | Average <= 1.1 x min |
| `LLEGA_AL_MAX` | Average >= 0.9 x max |
| `FIJO` | min = max |
| `ESCALA` | Anything else |
| `SIN_DATO` | No Container Insights history |

`replicas_needed_by_cpu` = office CPU used by all replicas / (request per replica x 0.7, a typical HPA
target). `min_above_need_beyond_ha` = the minimum is above max(2, needed).

## Building your own workbooks

Open each CSV with *Data > From Text/CSV*, comma delimiter. Opening it with a double click splits the
columns wrong in a Spanish Excel. Three workbooks cover the deck.

### 1. Cost and capacity

Inputs: `aks_clusters.csv`, `aks_node_pools.csv`, `aks_usage_monthly.csv`, `cost_by_resource.csv`.

| Figure | Pivot or formula |
|---|---|
| PROD clusters | `aks_clusters`, filter `environment` = PROD, count `name` |
| Monthly cost | `cost_by_resource`, filter `owner_type` = AKS, `environment` = PROD; rows `month`; sum `cost_usd` |
| Cost by type | Same pivot with `meter_category` in rows |
| Cost per cluster | Same pivot with `owner` in rows and `month` = last full month |
| Clusters with 4 nodes or fewer now | `aks_clusters`: helper column `system_nodes + user_nodes`, count <= 4 |
| Clusters under the 4-node minimum | `aks_node_pools`: helper column `IF(autoscaling="true", min_count, count)`; pivot by `cluster`, sum it, count < 4 |
| CPU and memory use in office hours | `aks_usage_monthly`, filter PROD; sum of `cpu_used_office_avg` / sum of `cpu_allocatable_avg` (same for memory, p95 and max) |
| Use per cluster | Same pivot with `cluster` in rows |

### 2. Microservices and HPA

Inputs: the four `aa-k8s-snapshot` CSVs, plus the cost per cluster from workbook 1.

| Figure | Pivot or formula |
|---|---|
| Microservices per layer (pattern 1) | `deployment_snapshot`, filter PROD and `platform_component` = NO; rows `layer`; count |
| Repeated capabilities | `capability_domains` as is |
| Clusters per microservices bucket (pattern 2) | `cluster_snapshot`, filter PROD, `in_inventory` = SÍ, `in_export` = SÍ; helper bucket column on `microservices`; add the cost with `XLOOKUP(cluster, owner, cost)` |
| Replicas now (pattern 4) | `hpa_snapshot`, filter PROD and `platform_component` = NO; helper bucket column on `replicas_now` |
| HPA behaviour | Same filter; rows `behaviour_ci`; count as % of total |
| Inflated minimums | Same filter plus `behaviour_ci` = NUNCA_ESCALA; rows `min_above_need` or `min_above_need_beyond_ha` |
| BIA (annex) | `cluster_snapshot` with the cost column; rows `bia`, columns `in_export` |
| Fits Container Apps by size | `deployment_snapshot`; rows `fits_aca_by_request` |

### 3. APIM traffic

Inputs: `apim_api_monthly.csv`, `apim_hourly_profile.csv`, `apim_apis.csv`.

| Figure | Pivot or formula |
|---|---|
| Requests per API and month | `apim_api_monthly`; rows `api_id`, columns `month`; sum `requests` |
| APIs above 100 tps | Same CSV, last full month, count `peak_5min_requests` > 30,000 |
| Off-hours share | Sum of `offhours_requests` / sum of `requests` |
| Traffic class per API | `apim_hourly_profile`: helper `week = avg_requests_per_hour x (5 on weekday, 2 on weekend)` and helper window (office: weekday 8–19; extended: any day 6–21; rest). Pivot `api_id` x window, sum `week`. Office >= 90% → OFICINA; office + extended >= 90% → EXTENDIDO; else 24/7 |

The class is per API. Mapping APIs to clusters needs the API → deployment mapping, which is still pending.

## Metrics not collected today

Ready to add if a question needs them. Each one is a change to the script and a rerun.

| Metric | Source | Answers |
|---|---|---|
| `cluster_autoscaler_unschedulable_pods_count`, `cluster_autoscaler_cluster_safe_to_autoscale` | AKS platform metrics | Do clusters run short of capacity? |
| `kube_pod_status_ready`, `kube_pod_status_phase` | AKS platform metrics | Pods not ready or pending |
| `node_disk_usage_percentage`, `node_network_in_bytes` / `_out_bytes` | AKS platform metrics | Disk and network use per node |
| `KubeEvents` (reason `OOMKilled`, `FailedScheduling`, `BackOff`) | Container Insights | Pods killed for memory or not scheduled: evidence that requests are too low |
| `cpuLimitNanoCores`, `memoryLimitBytes` | Container Insights `Perf` | Limits measured, not only from the export |
| `KubeNodeInventory` | Container Insights | Node labels, kubelet version, node conditions |
| `Duration`, `BackendDuration` | APIM platform metrics | Latency per API: what a cold start in Container Apps would add |
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
