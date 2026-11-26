# Backup guidelines: sources behind every figure

Personal reference for the guidelines deck (`src/bu_guidelines_deck.py`). The deck shows only the
regulation (slide 6). Use this file when someone asks where a limit, an RTO or a price comes from. All
links are public and were checked in November 2026.

## Regulation (also on slide 6)

| Norm | What the deck uses | Link |
|---|---|---|
| Ley 26702, art. 183 | Books and documents kept 10 years; microfilm or "analogous media" allowed | SPIJ (no direct link): <https://spij.minjus.gob.pe/>, search «Ley 26702». Text with notes up to May 2026 |
| Res. SBS 2660-2015 (consolidated v5) | Arts. 49 and 54: operations register, quarterly backup copy for 5 years, 48 business hours. Arts. 55 and 58: KYC, counterparties, unusual operations, 10 years | <https://intranet2.sbs.gob.pe/dv_int_cn/1540/v5.0/Adjuntos/2660-2015.r.pdf> |
| Res. SBS 504-2021 (consolidated v5) | Art. 12, num. 3, lit. f) and h): backup strategy and tests. Arts. 23 and 25: cloud and processing abroad | <https://intranet2.sbs.gob.pe/dv_int_cn/2046/v5.0/Adjuntos/504-2021.R.pdf> |
| Circular BCRP 0011-2026 | Art. 46.5: 2-hour recovery. Annex V, section VI: daily backup, yearly restore test, 10 years | <https://www.bcrp.gob.pe/docs/Transparencia/Normas-Legales/Circulares/2026/circular-0011-2026-bcrp.pdf> |

Local copies of the official texts, downloaded on 2026-11-16, are in [regulation/](regulation/):
[Ley 26702](regulation/ley-26702-art-183.pdf) (cover and art. 183 only; the full SPIJ PDF is 59 MB),
[Res. SBS 2660-2015](regulation/res-sbs-2660-2015.pdf), [Res. SBS 504-2021](regulation/res-sbs-504-2021.pdf)
and [Circular BCRP 0011-2026](regulation/circular-bcrp-0011-2026.pdf). The examples on slide 6 come from
these texts: art. 183 gives none; SBS 2660-2015 art. 49.1 to 49.3 (operations and thresholds), 55 and 58;
SBS 504-2021 art. 12.3 f) and h); BCRP Anexo V sec. VI and the definitions in art. 2.

## Granularity and snapshots (slide 2)

| Figure | Source |
|---|---|
| Layers of the application stack (canal, backoffice, infra SW base, integration platforms without data at rest, security); logs follow their own hot/archive retention | Bank's Confluence page "Best-practices para la gestión de tiempos de retención de logs" (2024, internal) |
| VM Instant Restore snapshots: Standard 1–5 days (default 2), Enhanced 1–30 days (default 7); Enhanced first snapshot is a full copy, then incremental; cost grows with daily churn (2%–7%) | <https://learn.microsoft.com/en-us/azure/backup/backup-instant-restore-capability> |
| Incremental disk snapshots hold only the changes since the previous one and are billed by used size | <https://learn.microsoft.com/en-us/azure/virtual-machines/disks-incremental-snapshots> |

## Service limits (slide 3)

| Figure | Source |
|---|---|
| VM backup: daily (Enhanced every 4 h), daily/weekly/monthly/yearly retention up to 99 years | <https://learn.microsoft.com/en-us/azure/backup/backup-support-matrix-iaas> |
| Azure Files: snapshot and vaulted backup | <https://learn.microsoft.com/en-us/azure/backup/azure-file-share-backup-overview> |
| Archive: monthly and yearly points only, 3 months in standard tier, 6 months left; not on ZRS vaults; 180-day minimum | <https://learn.microsoft.com/en-us/azure/backup/archive-tier-support> |
| SQL Database: PITR 1–35 days, LTR up to 10 years, LTR immutability, free backup storage up to max size | <https://learn.microsoft.com/en-us/azure/azure-sql/database/automated-backups-overview> |
| SQL Managed Instance: PITR 1–35 days, LTR up to 10 years, no LTR immutability, GRS = 2 x published price | <https://learn.microsoft.com/en-us/azure/azure-sql/managed-instance/automated-backups-overview> |
| PostgreSQL Flexible: PITR 7–35 days, RPO up to 5 min, geo backup only at creation, LTR weekly/monthly/yearly, 1 TiB max | <https://learn.microsoft.com/en-us/azure/postgresql/backup-restore/concepts-backup-restore> |
| PostgreSQL Flexible: LTR in a Backup vault up to 10 years, instance fee per 250 GB | <https://learn.microsoft.com/en-us/azure/backup/backup-azure-database-postgresql-flex-overview> |
| MySQL Flexible: PITR 1–35 days, log every 5 min; Azure Backup LTR paused | <https://learn.microsoft.com/en-us/azure/mysql/flexible-server/concepts-backup-restore> |
| Cosmos DB: continuous 7, 30 or 35 days; 7 days free; changes backed up within 100 s | <https://learn.microsoft.com/en-us/azure/cosmos-db/continuous-backup-restore-introduction> |
| Cosmos DB: LTR in a Backup vault up to 10 years, billed from July 1, 2026 | <https://learn.microsoft.com/en-us/azure/backup/backup-azure-cosmos-db-overview> |

## RTO and RPO (slides 3 and 5)

| Figure on the slide | Official figure | Source |
|---|---|---|
| VM vault: ~10 h per TB | "A transfer rate of 30 MBps is considered optimal"; 2 TB of allocated data can take up to 20 h. 1,024 GB at 30 MB/s = 9.7 h | <https://learn.microsoft.com/en-us/azure/backup/backup-azure-vm-backup-faq> |
| VM RPO: 24 h (4 h Enhanced) | Same FAQ, "minimum RPO and RTO" | Same |
| SQL: normally < 12 h | Managed Instance table: "Restore usually takes less than 12 hours" | <https://learn.microsoft.com/en-us/azure/azure-sql/managed-instance/automated-backups-overview> |
| SQL: RPO 5 to 10 min | Log backups about every 10 minutes; business continuity table | <https://learn.microsoft.com/en-us/azure/azure-sql/database/business-continuity-high-availability-disaster-recover-hadr-overview> |
| PostgreSQL / MySQL: minutes to a few hours | "The overall recovery time usually takes from few minutes up to a few hours" | PostgreSQL backup page above |
| Cosmos DB: ~30 to 90 min per TB | "Restoring 1 TB of data typically takes 30-90 minutes" (same-account restore) | <https://learn.microsoft.com/en-us/azure/cosmos-db/restore-in-account-continuous-backup-frequently-asked-questions> |
| Archive: rehydrate up to 15 h | Standard priority "might take up to 15 hours"; High under 1 hour for objects under 10 GB | <https://learn.microsoft.com/en-us/azure/storage/blobs/archive-rehydrate-overview> |
| Vendors: ~2.3 h per TB at 1 Gbps | Calculation, not a vendor figure: 1,024 GB x 8 bits / 1 Gbps = 8,192 s. The bank's real link speed changes it | — |

## Backup redundancy names (slides 3 and 5)

The deck shows LRS and ZRS only, the options the bank uses; replication and geo-redundant backup are out of
its scope (review feedback, 2026-11-16). Keep in mind that LRS and ZRS backups do not survive the loss of a
whole region, and that Archive is not available on ZRS vaults. The geo options below stay for reference and
for `bu_guidelines_cost.py`.

| Service | Options | Source |
|---|---|---|
| SQL Database and Managed Instance, PITR and LTR | LRS, ZRS, RA-GRS, RA-GZRS. There is no plain GRS meter; RA-GRS and RA-GZRS cost the same (PITR 0.20, LTR 0.05 US$/GB-month list in East US 2) | Retail Prices API meters `RA-GRS Data Stored`, `RA-GZRS Data Stored`; <https://learn.microsoft.com/en-us/azure/azure-sql/database/automated-backups-overview> |
| Recovery Services vault (VM, Files) | LRS, ZRS, GRS. Cross Region Restore needs GRS, "incurs extra charges" (billed with the `Standard RA-GRS Data Stored` meter) and can't be reverted | <https://learn.microsoft.com/en-us/azure/backup/backup-create-recovery-services-vault> |
| PostgreSQL and MySQL Flexible | Local, zone or geo-redundant backup; geo bills twice the backup GB | PostgreSQL and MySQL backup pages above |
| Cosmos DB | Continuous: the region's redundancy, not selectable. Periodic extra copies: LRS, ZRS or RA-GRS | Retail Prices API meters `Periodic Backup - LRS/ZRS/RA-GRS Data Stored`; continuous backup page above |
| Archive (Azure Backup) | LRS and GRS; not supported on ZRS vaults | Archive tier page above |

## Worked example (slide 4)

One real production database, picked by its application description: TTIB, Transferencias Interbancarias
(Azure SQL Database Hyperscale, East US 2). SBS 2660-2015 art. 49 names transfers as an example of the
operations register; which tables hold it is still to be confirmed with the application owner.

| Input | Value | Source |
|---|---|---|
| Used size | 2,435 GB (grew about 1 TB in the last year; the example keeps it constant) | Azure Monitor metric `storage`, 2026-11-15 |
| PITR 7 days, billable | 5,166 GB-month (what exceeds the free allowance of one database size) | October 2026 bill, meter `Backup - RA-GRS Data Stored` |
| Size of one LTR copy | About the whole database: the bill shows 32,954 LTR GB-month for ~16 copies taken while the database grew from 1.5 to 2.5 TB | October 2026 bill, `microsoft.sql/longtermretentionservers` |
| Prices | Hyperscale `Backup - ZRS Data Stored` (PITR) and `Backup ZRS Data Stored` (LTR), East US 2, list less 30%. The real October bill came out 30% below list on LTR and 43% on PITR | Retail Prices API; <https://learn.microsoft.com/en-us/azure/azure-sql/database/hyperscale-automated-backups-overview#backup-storage-costs> |

Copies are simulated week by week: the monthly copy is the month's first weekly backup, the yearly copy the
year's first, and a backup that meets two rules is kept once. The database starts with the copies the base
rule already keeps. Change `EXAMPLE` and `EXAMPLE_SCENARIOS` in the script to use another database.

## Prices (slide 5)

Every price comes from the Azure Retail Prices API; slide 5 shows Central US only, LRS and ZRS, less
the bank's 30% discount (`DISCOUNT` in the script, `--discount` to change it). On the October 2026 bill, SQL
LTR came out about 38% below list, so 30% is on the conservative side. The exact meter names are in the
speaker notes of slide 5.

| Page | Link |
|---|---|
| Pricing calculator | <https://azure.microsoft.com/en-us/pricing/calculator/> |
| Azure Backup (vault, Archive, protected instances) | <https://azure.microsoft.com/en-us/pricing/details/backup/> |
| Azure SQL Database (PITR and LTR) | <https://azure.microsoft.com/en-us/pricing/details/azure-sql-database/single/> |
| Azure SQL Managed Instance (PITR and LTR) | <https://azure.microsoft.com/en-us/pricing/details/azure-sql-managed-instance/single/> |
| PostgreSQL Flexible Server (backup storage) | <https://azure.microsoft.com/en-us/pricing/details/postgresql/flexible-server/> |
| MySQL Flexible Server (backup storage) | <https://azure.microsoft.com/en-us/pricing/details/mysql/flexible-server/> |
| Cosmos DB (continuous backup) | <https://azure.microsoft.com/en-us/pricing/details/cosmos-db/autoscale-provisioned/> |
| Retail Prices API | <https://prices.azure.com/api/retail/prices> |

Vendor tariffs (D3, Opción 1) come from the vendors' quotes in the backup guidelines review; the unit
(protected or stored TB) is still to be confirmed.
