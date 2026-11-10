# Backup cost screenshots from Cost analysis

Portal views that show the backup cost exactly as Azure bills it. Each screenshot is evidence for one
cost slide; the workbook gives the reference value to check it against.

## Common setup

1. Portal > Cost Management > Cost analysis.
2. Scope: the management group that holds the bank's subscriptions (change it with "Scope").
   A subscription scope shows only that subscription.
3. Date: Aug–Oct 2026 (custom). Granularity: Monthly. Chart: Column (stacked).
4. Add the filters of the view below with "Add filter". "Service name" is the meter category.
5. Save each view ("Save as") with the name in the heading, so it can be reopened next month.
6. Take the screenshot with the scope, the date, the filter chips and the "Actual cost" total visible.

The October value of each view should match the October bar. Reference values are in the workbook:
sheet `Gasto_medido` (cost by meter) and sheet `Por_ambiente` (cost by environment).

## Views

| # | View | Filters | Group by |
|---|---|---|---|
| 1 | Backup - vaults | Service name = Backup | Meter |
| 2 | Backup - manual disk snapshots | Resource type = microsoft.compute/snapshots | Meter |
| 3 | Backup - Azure Backup snapshots | Resource type = microsoft.compute/restorepointcollections | Meter |
| 4 | Backup - SQL PITR | Service name = SQL Database; Meter subcategory = Single/Elastic Pool PITR Backup Storage | Meter subcategory |
| 5 | Backup - SQL Hyperscale PITR | Service name = SQL Database; Meter subcategory = SingleDB Hyperscale - Storage; Meter = every `Backup - ... Data Stored` | Meter |
| 6 | Backup - SQL LTR | Service name = SQL Database; Meter subcategory = LTR Backup Storage, SingleDB Hyperscale - LTR Backup Storage | Meter subcategory |
| 7 | Backup - PostgreSQL, MySQL, Cosmos | Service name = Azure Database for PostgreSQL, Azure Database for MySQL, Azure Cosmos DB; Meter contains "Backup" | Service name |

Notes:

- View 3 is part of Azure Backup (instant restore points of protected VMs), not manual snapshots.
  Show it next to view 1, not next to view 2.
- View 5 needs the Meter filter: without it, the Hyperscale storage subcategory also includes the
  database's own data, which is not backup.
- View 4 alone is not the SQL backup cost. Views 4, 5 and 6 together are.

## Same views by environment

Repeat any view with Group by = Tag > `environment`. This is the measured split by environment.

- Resources without the tag show as "untagged". Restore point collections (view 3) and LTR servers
  (view 6) usually have no tag, so most of their cost lands there.
- Compare with sheet `Por_ambiente`. That sheet splits vault and snapshot cost from the inventory
  (by protected VM and by source disk), so small differences with the tag view are expected.

## If a number does not match

- Group the same view by Subscription and look for the difference there.
- Check the scope: the inventory scripts read only the subscriptions the account can see.
- Check the month: the workbook uses October 2026 only.
