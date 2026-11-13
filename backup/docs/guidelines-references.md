# Backup guidelines: sources behind every figure

Personal reference for the guidelines deck (`src/bu_guidelines_deck.py`). The deck shows only the
regulation (slide 3). Use this file when someone asks where a limit, an RTO or a price comes from. All
links are public and were checked in November 2026.

## Regulation (also on slide 3)

| Norm | What the deck uses | Link |
|---|---|---|
| Ley 26702, art. 183 | Books and documents kept 10 years; microfilm or "analogous media" allowed | SPIJ (no direct link): <https://spij.minjus.gob.pe/>, search «Ley 26702». Text with notes up to May 2026 |
| Res. SBS 2660-2015 (consolidated v5) | Arts. 49 and 54: operations register, quarterly backup copy for 5 years, 48 business hours. Arts. 55 and 58: KYC, counterparties, unusual operations, 10 years | <https://intranet2.sbs.gob.pe/dv_int_cn/1540/v5.0/Adjuntos/2660-2015.r.pdf> |
| Res. SBS 504-2021 (consolidated v5) | Art. 12, num. 3, lit. f) and h): backup strategy and tests. Arts. 23 and 25: cloud and processing abroad | <https://intranet2.sbs.gob.pe/dv_int_cn/2046/v5.0/Adjuntos/504-2021.R.pdf> |
| Res. SBS 00814-2025 | Amends Res. SBS 877-2020: liabilities with RPO zero and RTO up to 3 hours | <https://intranet2.sbs.gob.pe/dv_int_cn/2483/v1.0/Adjuntos/814-2025.R.pdf> |
| Circular BCRP 0011-2026 | Art. 46.5: 2-hour recovery. Annex V, section VI: daily backup, yearly restore test, 10 years | <https://www.bcrp.gob.pe/docs/Transparencia/Normas-Legales/Circulares/2026/circular-0011-2026-bcrp.pdf> |

## Service limits (slide 1)

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

## RTO and RPO (slide 2)

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

## Prices (slide 2)

Every price comes from the Azure Retail Prices API, list price, East US 2 and Central US. The exact meter
names are in the speaker notes of slide 2.

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
