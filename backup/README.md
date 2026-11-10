# Azure backup reporting

Read-only inventory of how the Azure estate is backed up across four fronts (database, storage,
compute, native), what it costs, and where to save or close gaps. Output: an Excel annex and two
slide decks.

## Layout

| Path | Content | In git |
|---|---|---|
| `src/` | Python scripts (`bu_*.py`) and `requirements.txt` | Yes |
| `src/decks/` | `build_decks.js` and its `package.json` | Yes |
| `results/backup-<front>/` | CSV and JSON written by each inventory script | No |
| `reports/<date>-backup/` | `backup_insights.xlsx`, `metrics.json`, the two `.pptx` and a README with that run's scope | No |

Guides:

- [docs/cost-screenshots.md](docs/cost-screenshots.md): Cost analysis views for each backup cost, overall
  and by environment tag.
- [docs/prod-comparison.md](docs/prod-comparison.md): third-party comparison limited to production.
- [docs/sql-deleted-database-restore.md](docs/sql-deleted-database-restore.md): what PITR and LTR keep
  after a database or server is deleted, where to see it and how to restore it.
- [docs/backup-export.md](docs/backup-export.md): whether vault, PITR and LTR backups can be downloaded
  or moved to another vendor, and how.
- [docs/deck-and-prompts.md](docs/deck-and-prompts.md): slide order, the source of each slide and the
  Copilot prompts to build the deck on the corporate laptop.

The workbook's `Por_ambiente` sheet splits every indicator and the vault and snapshot cost by the
`environment` tag (PROD, CERT, DESA, untagged), without Databricks, ORCA, ARO or AKS node resources.

`results/` and `reports/` hold estate data, so they stay on the machine that ran the scripts. Every
script writes to these folders by default, wherever it is run from; `--output-dir`, `--results` and
`--workbook` override them.

## Run order

Python 3.10+ and an `az login` session (the scripts use it through `DefaultAzureCredential`).
Commands run from the repository root and work the same in PowerShell and bash.

```bash
uv pip install -r backup/src/requirements.txt
python backup/src/bu_compute_inventory.py
python backup/src/bu_database_inventory.py
python backup/src/bu_storage_inventory.py
python backup/src/bu_native_protection_inventory.py
python backup/src/bu_restore_compat.py
python backup/src/bu_backup_cost.py --from-date 2026-10-01 --to-date 2026-10-31
python backup/src/bu_insights_workbook.py --output-dir backup/reports/<date>-backup
```

Then:

1. **Open `backup_insights.xlsx` in Excel, save it and close it.** The workbook is written with
   formulas only; Excel computes and stores their values on save. Skipping this step leaves every
   saving, risk count and calibrated cost empty. `bu_deck_metrics.py` stops with an error if it finds
   the workbook in that state.
2. `python backup/src/bu_deck_metrics.py --workbook backup/reports/<date>-backup/backup_insights.xlsx --output backup/reports/<date>-backup/metrics.json`
3. Build the decks with Node 18+:
   ```bash
   cd backup/src/decks
   npm install
   node build_decks.js ../../reports/<date>-backup/metrics.json ../../reports/<date>-backup
   ```

Without Node on the laptop, run step 3 in Azure Cloud Shell: clone the repository there, upload
`metrics.json`, run the three commands with that path, and download the two `.pptx`. Only
`metrics.json` leaves the laptop; steps 1 and 2 must still run where Excel is.

Runtime on 350 subscriptions: compute ~1.5 min, database ~4 min (includes SQL server locks and backups of deleted databases), native ~1 min, restore <1 min,
storage ~25 min, cost ~70 min (Cost Management throttles; the cost script caches per subscription and
resumes if interrupted).

## Tracing a figure

- `Fuentes` sheet: each indicator with its value, data type (measured, derived, assumption, vendor
  data), source, output file and how to check it in the portal.
- `Consultas` sheet: the exact Resource Graph queries, ARM call patterns and Cost Management request
  bodies each script ran, with run date, rows and errors.
- `Oportunidades`: source of the quantity and of the cost for each saving.
- `Supuestos` and `Comparacion_terceros`: yellow cells are assumptions or vendor data; everything else
  is a formula.
- Cost is `ActualCost` from the Cost Management query API, the same data as Cost analysis in the
  portal (with contract discounts). Effective rates are billed cost divided by billed quantity.

## Caveats

- Savings are estimates. Backup rates are measured; Redis and App Service Plan deltas, the retention
  reduction, the removable share of snapshots and the share of geo-replication without a DR
  requirement are assumptions to validate.
- Orphan savings are a range: the floor is measured (vaults where every item is orphaned, plus
  protected-instance fees); the upper value prorates mixed vaults by item count.
- Coverage depends on the account: subscriptions it cannot read are missing from every figure. Each
  run's README in `reports/` records how many subscriptions it saw.
- Backups taken inside the guest OS (SQL scripts, Velero/OADP) are not visible to Azure APIs.
