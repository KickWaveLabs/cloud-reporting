# Exporting Azure backups to another vendor

No Azure backup can be downloaded or handed to another vendor as is. The only way out is to restore
it first and then export the restored copy in a standard format.

| Backup | Direct download | How to get it out |
|---|---|---|
| Vault, VM | No (proprietary, encrypted format) | Restore as managed disks, export the disk as VHD, download it |
| Vault, SQL Server in VM | No | "Restore as files": writes `.bak` files to a folder on a VM; copy them from there |
| Vault, Azure Files | No | Restore to a file share and copy the files |
| Azure SQL PITR | No (managed by Microsoft) | Restore to a new database, export it to BACPAC |
| Azure SQL LTR | No | Restore the LTR backup to a new database, export it to BACPAC |
| PostgreSQL / MySQL flexible | No | Restore to a new server, then `pg_dump` / `mysqldump` |
| Cosmos DB | No | Restore to a new account, then export with Data Factory or a migration tool |

## When asked

- Third-party vendors (Veeam, Commvault and similar) cannot read Azure backups. They protect the
  live resources with their own agent or integration.
- When changing vendor, existing Azure backups stay in Azure until they expire, or are restored and
  exported one by one.
- Every export has a cost: the temporarily restored database or disk, storage for the exported
  file, and egress if it leaves Azure.

## Documentation

- Restore SQL Server in Azure VM, including restore as files:
  https://learn.microsoft.com/azure/backup/restore-sql-database-azure-vm
- Restore VM disks from a vault: https://learn.microsoft.com/azure/backup/backup-azure-arm-restore-vms
- Download a VHD: https://learn.microsoft.com/azure/virtual-machines/windows/download-vhd
- Azure SQL PITR restore: https://learn.microsoft.com/azure/azure-sql/database/recovery-using-backups
- Azure SQL LTR restore: https://learn.microsoft.com/azure/azure-sql/database/long-term-backup-retention-configure
- Export Azure SQL to BACPAC: https://learn.microsoft.com/azure/azure-sql/database/database-export
