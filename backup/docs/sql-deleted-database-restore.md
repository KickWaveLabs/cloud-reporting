# Azure SQL Database: backups after a database or server is deleted

What survives a deletion, where to see it, and how to restore it. Applies to Azure SQL Database
(single databases, elastic pools and Hyperscale).

## What survives

| Event | Point-in-time restore (PITR) | Long-term retention (LTR) |
|---|---|---|
| Database deleted | Kept for the database's PITR retention (1–35 days; 7 in most of the estate). Restores to the deletion time or an earlier point within that window. | Kept until each backup expires. |
| Logical server deleted | **Lost** for every database on the server. Cannot be recovered. | Kept until each backup expires. |

- Both keep billing while they exist.
- Every restore creates a **new** database. Nothing is restored in place.
- A `CanNotDelete` lock on production SQL servers protects the PITR of all their databases. Without
  it, deleting the server loses every PITR at once.

## Where to see them in the portal

| Situation | PITR | LTR |
|---|---|---|
| Database deleted, server exists | SQL server > Data management > **Deleted databases** | SQL server > Data management > **Backups** > tab **Available LTR backups** |
| Server deleted | Not available (lost) | No portal page: the backups hang from a server that no longer exists. Use Cloud Shell from the portal (commands below). |

## Commands

Listing is read-only. Restoring creates a database, so it needs write access on the target server
and belongs to a restore drill or a POC, not to reporting.

```bash
# PITR: deleted databases that can still be restored on a server
az sql db list-deleted --subscription <sub> -g <rg> -s <server> -o table

# PITR: restore a deleted database as a new one (same server)
az sql db restore --subscription <sub> -g <rg> -s <server> -n <deleted-db-name> \
  --deleted-time <deletion-time> --time <point-in-time> --dest-name <new-db-name>

# LTR: backups of deleted databases in one region, servers deleted or not
az sql db ltr-backup list --subscription <sub> --location eastus2 --database-state Deleted -o table

# LTR: restore one backup to any existing server of the same subscription
az sql db ltr-backup restore --backup-id "<backup id>" \
  --dest-server <server> --dest-resource-group <rg> --dest-database <new-db-name>
```

With geo-redundant LTR storage the target server can be in another region.

## How it is reported

`bu_database_inventory.py` writes two files to `results/backup-database/`:

- `database_sql_servers.csv`: every logical server with its environment, whether a delete lock covers
  it (on the server, its resource group or the subscription) and how many deleted databases it can
  still restore by PITR.
- `database_deleted_backups.csv`: one row per deleted database with backups left. `backup_type` is
  `PITR` (server exists, restorable for days) or `LTR` (kept until `available_until`).
  `server_exists = NO` means those backups can only be reached with the commands above.

Environment comes from the tags of the server's databases. When the server has none (for example, it
was deleted), it comes from the name suffix: P, C or D plus digits (`msqleu2appxp01` is PROD). ORCA
Security servers are marked as platform and left out of the counts.

`bu_insights_workbook.py` adds both as sheets `D_SQL_servidores` and `D_BD_eliminadas`, and as rows of
`Por_ambiente`.

## Proving it works

Reading configuration does not prove a restore works. To prove it:

1. Pick one deleted database with LTR from `D_LTR_eliminadas`.
2. With the owning team, restore its newest backup (`newest_backup_id`) to a server of the same
   subscription under a new name.
3. Record how long it took, the size, and whether the application team can read the schema.
4. Drop the restored database when done.

The same drill on a small test database (create, wait for the first LTR backup, delete, restore from
Deleted databases and from LTR) covers the PITR case too.
