# Internal guidelines behind the backup deck

Summary of two internal documents read on 2026-11-25 (from screenshots of the bank's wiki), and what each
one means for the backup guidelines deck. Neither gives retention periods for backups: the deck fills that gap.

## LSEG_008: Lineamiento de Respaldo y Retención de Información

Mandatory general guideline (version 6, October 2026), based on Política 4180.223.04 "Respaldo y retención de
información". Scope: every system in the bank's disaster recovery plan. References NIST SP 800-34 Rev. 1.

| Section | What it says | Meaning for the deck |
|---|---|---|
| 1.1 Business impact analysis (BIA) | Each system is classified by impact (low, medium, high, very high). Backup frequency follows the BIA RPO; backup method and medium follow the BIA RTO | The decision tree needs each application's criticality. RPO and RTO set frequency and method, not retention |
| 1.3 Roles | Teams must know the backup and recovery objectives; roles are in norm 4110.782.10 | — |
| 1.4 Third parties | Outsourced storage must meet international standards and SLAs (availability, response times, restore tests) | Covered by Azure's certifications; not on the slides |
| 2.1 A–C Backup design, methods, frequency | Full, differential and incremental; "combine methods per system". The backup strategy belongs to the application owner; the infrastructure (hosting) team runs it | Matches "se decide por aplicación" on slide 2 |
| 2.1 D Naming | Each copy is labelled with server, database or folder, method, sequence, date, start and end time | Azure vaults and LTR keep this metadata themselves |
| 2.1 E Media | Centralized, automated, encrypted, monitored. Authorized media: network storage through the bank's authorized backup solution (on-premises and cloud), online backup, tapes (never taken out of the data centre) | The user confirmed Azure Backup is authorized |
| 3.2 Access control | High and very high criticality: backups logically separated from lower-criticality ones, need-to-know access. All backups encrypted | — |
| 3.3 Encryption | AES-256 or stronger; keys under the key life-cycle guideline | Azure Backup encrypts at rest by default; customer-managed keys are optional |
| 3.4 Storage sites | High and very high impact: copies in an alternate site. Critical systems: copies kept "off-line" against ransomware | See "Cloud equivalents" below |
| 4 Retention | Keep backups "as long as the business requires", by asset type, regulation, criticality, change rate, business continuity, forensics and fraud. Periods per frequency (hourly, daily, weekly, monthly, yearly) go in each system's contingency plan. Expired backups must be deleted under the secure deletion guideline | No numbers: the deck sets them. The deletion rule backs the snapshot expiry on slide 2 |
| 5 Restore | Restore to a known, safe state; validate integrity with CRC or hash | — |
| 5.1 Restore tests | High and very high criticality: full restore test at least once a year. Low and medium: only on request, at least yearly when requested | Slide 1 asks for yearly tests only in regulated scopes; it could add "criticidad alta y muy alta" |

## Best practice: log retention times

Approved in May 2024 by Security and Technology Architecture. It solved the same problem the backup deck
solves: with no clear rule, every application kept every log for 10 years and paid for it.

- Retention is set by three axes: data type (DAC or not, regulatory, legal or contractual, historical),
  application type and log type.
- Application types, with the examples the document gives:
  - Canal: Homebanking, Mobile Banking, Yape.
  - Backoffice: SaaS services, Salesforce, SIGA.
  - Infrastructure and base software: operating systems, databases, mainframe, network devices.
  - Integration platforms: APIs, microservices, middleware ("no guardan datos en reposo").
  - Security platforms: SIEM, EDR, DLP, PAM.
- Two tiers: hot (in the platform, premium cost) and archive (storage account or the bank's backup systems,
  low cost), then deletion.
- Values: transactional DAC logs of channels, hot up to 1 year and archive up to 10 years; backoffice DAC,
  hot 3 months and archive 10 years; non-DAC, infrastructure metrics and integration platforms, hot 3 months
  and no archive; audit and security events with DAC, hot 3 months and archive 1 year.
- Retention of "activos digitales" (contracts, vouchers) is explicitly out of its scope.
- Its diagram draws the application stack in lanes (transaction, channel or backoffice, base infrastructure,
  security) and the life cycle hot → archive → delete.

Reused on slide 2: the application types, "integration holds nothing at rest" and the hot → archive → delete
life cycle for snapshots.

## Cloud equivalents for the LSEG_008 controls

| LSEG_008 asks for | In Azure | Source |
|---|---|---|
| Copy in an alternate site (3.4 a) | ZRS keeps copies in three availability zones (separate data centres) of the region; GRS adds the paired region | <https://learn.microsoft.com/en-us/security/benchmark/azure/mcsb-v2-backup-recovery> |
| Off-line copy against ransomware (3.4 b) | "Logically air-gapped vault": vaulted backups live in a Microsoft-managed subscription and tenant, apart from production; it requires multi-user authorization (Resource Guard), with soft delete on and immutability optional | <https://learn.microsoft.com/en-us/azure/backup/secure-backup> |
| Copy nobody can change | Immutability is a property of the backup (immutable vault, SQL LTR immutability), not of the protected resource | <https://learn.microsoft.com/en-us/azure/backup/backup-azure-immutable-vault-concept> |

Snapshots are not off-line: they sit in the production subscription next to the disk.

## Open for the backup decision tree (ADT)

- The regulated scope follows what an application's database stores, not what the application processes: a
  channel processes transactions, but their trace is in logs (log retention applies) and the result (balances,
  movements) is stored by the system of record, usually Core or mainframe. Only that database needs the
  regulated backup.
- Which layer holds the source record of each regulated scope: UX/channel, BS, Core or mainframe. Only that
  layer needs the regulated backup; layers holding copies or caches follow the base rule or none.
- Where each application's BIA criticality is recorded (business continuity team, to confirm).
- Whether "alternate site" means another zone (ZRS) or another region (GRS) for the bank.
