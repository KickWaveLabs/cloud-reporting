# Cost analysis views for AKS

Screenshots for the "current cost" slides. They show the bank's own billing as **amortized cost**,
the same data `aa_cost.py` reads, so each figure in the deck can be checked in the portal. Node VMs are
covered by reservations or savings plans: in the default "Actual cost" view they show close to 0, so
switch the view's cost type to "Amortized cost" before taking each screenshot.

Period: August to October 2026, monthly. Currency: USD. Scope: the subscription of each cluster
(the account is Reader per subscription, not at management group).

A cluster's cost lives in two resource groups: the cluster's own (the `Kubernetes service` resource:
tier, Defender, cost analysis add-on) and the node resource group `MC_<rg>_<cluster>_<region>`
(VMs, OS disks, load balancer, public IPs). Cost analysis cannot OR two filters, so take two views
or filter by both resource group names.

## Views

| # | View | Scope | Filters | Group by | Chart | Slide |
|---|---|---|---|---|---|---|
| 1 | Cost of one PROD cluster | Subscription | Resource group name = cluster RG and node RG | Meter category | Column, monthly | Current cost |
| 2 | What a cluster pays for | Same | Same | Meter | Table | Current cost |
| 3 | AKS cost per cluster | Subscription (repeat per subscription with PROD clusters) | Resource group name starts with `MC_` | Resource group name | Column, monthly | Current cost |
| 4 | Namespace cost (only if the cost analysis add-on is on) | Subscription | Kubernetes clusters view | Namespace | Table | One app per cluster |

Steps for view 1:

1. Cost Management > Cost analysis > change scope to the cluster's subscription.
2. View: Accumulated cost → switch to "Cost by resource" or a custom view; Granularity: Monthly;
   date range Aug 1 – Oct 31, 2026.
3. Add filter Resource group name: the cluster RG and its `MC_` RG.
4. Group by Meter category. Chart: Column (stacked).
5. Check: the total for each month equals `Measured_AKS` column "Monthly cost (USD)" × months for that
   cluster (the workbook shows the monthly average).

## Before taking the screenshots

- Hide subscription ids and names of people if they appear in the view (owners in resource group
  names or tags).
- Keep the date range and currency visible in the capture.
- Name the files `aks-cost-<view>-<cluster>.png` and keep them with the workbook, not in git.
