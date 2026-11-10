# Third-party comparison, production only

Sheet `Comparacion_terceros` compares the whole vault estate against the third-party tariffs. To
compare production only, point its measured block at the production figures of sheet `Por_ambiente`.

Run the workbook first (`bu_insights_workbook.py`), open it in Excel, save it, and work on a copy
(for example `backup_insights_prod.xlsx`) so the original keeps the estate-wide comparison.

The formulas below use names the workbook defines (Formulas > Name Manager), so they do not depend on
row numbers:

| Name | What it is |
|---|---|
| `prod_vault_cost` | Vault cost assigned to PROD, USD/month |
| `prod_vms_360` / `all_vms_360` | VMs with 360-day retention: PROD / all environments |
| `prod_snapshot_gb` | GB of manual snapshots whose source disk is PROD |
| `prod_snapshot_cost` | Manual snapshot cost assigned to PROD, USD/month |

## What changes

| Cell in `Comparacion_terceros` | Key | Estate formula (today) | Production formula |
|---|---|---|---|
| C20 | `vault_cost` | Total cost of every vault | `=prod_vault_cost` |
| C19 | `vault_tb` | TB stored in every vault | `=SUMIFS(Gasto_medido!$D:$D,Gasto_medido!$A:$A,"Backup",Gasto_medido!$B:$B,"*Data Stored")/1024*$C$20/(Costo_calibrado!$E$2+Costo_calibrado!$F$2+Costo_calibrado!$G$2)` |
| C21 | `vault_opt_cost` | Cost minus orphans (O1) and GFS retention (O2) | `=MAX(0,$C$20-SUMIFS(Oportunidades!$E:$E,Oportunidades!$A:$A,"O2")*prod_vms_360/MAX(1,all_vms_360))` |

Why:

- `prod_vault_cost` is the non-orphan part of each vault, split by the disk GB of the PROD VMs it
  protects. Orphan items are already outside it, so O1 is not subtracted again.
- O2 applies only to the PROD share of the VMs with 360-day retention.
- TB is prorated by cost, the same rule the sheet already uses for the optimized volume.

Everything below these cells (`vault_opt_tb`, `base_tb`, `azure_base`, the scenario table and the
break-even column) recalculates by itself.

## Including disk snapshots

Leave `include_snapshots` (C10) at 0 unless the third party would also replace manual disk
snapshots. If it would, set C10 to 1 and also change:

| Cell | Key | Production formula |
|---|---|---|
| C23 | `snap_tb` | `=prod_snapshot_gb/1024` |
| C24 | `snap_cost` | `=prod_snapshot_cost` |

The snapshot GB is the provisioned size of each snapshot, and Azure bills the used size, so this TB
figure is an upper bound.

## Checks before presenting

- `vault_cost` (C20) matches the PROD bar of the "Backup - vaults" view grouped by tag `environment`
  (see [cost-screenshots.md](cost-screenshots.md)), within the untagged share.
- The scenario "Azure Backup actual" shows the same US$ per TB as the estate version: the rate does
  not change, only the volume.
- Confirm with the vendor whether the tariff is per protected TB (source size) or per stored TB.
  Azure here is stored TB; if the vendor charges per protected TB, the volumes are not comparable.
