# AGENTS.md

## Project Overview

Read-only reporting on the bank's cloud estate: resource inventory, metrics, backup status, cost
and efficiency opportunities. Also POCs, only inside resource groups assigned to the user.
Azure today. AWS and GCP get their own section when work on them starts.

## Response Style

- Concise. Bullets, not paragraphs. No filler, no repeated explanations.
- Short sentences. Result first, then the query or script behind it, in one line.
- Say when a result is partial: throttling, missing access, row limits.
- Reply in Spanish. Scripts, comments, commits and docs in English.

## Access

- The user is Reader on most subscriptions, not all of them. Report access errors as gaps.
- Reporting is read-only on every cloud. Changes happen only under a POC.
- Signing in and switching the active account or subscription are the user's to run.

## POCs

- Each POC has one file, `pocs/<name>.md`: cloud, subscription, resource group, role, goal, pending.
- Read it before working on the POC. Update pending when something finishes.
- Change only the resource group named in that file. With no file, stay read-only.
- Show each change command or plan with its subscription and resource group, then wait for approval.
- Role assignments and anything at subscription scope need an explicit request, even if allowed.
- Inside that resource group, data-plane and credential commands are allowed when the POC needs
  them. Never print secret values.
- Pass the account explicitly (`--subscription` on `az`, `--context` on `kubectl`). The machine's
  defaults may be unrelated.

## Out of Scope for Reporting

- People: users, groups, collaborators, role assignments, Entra ID. Leave `tags` out of results
  unless asked, because they can hold owner names and emails.
- Data-plane content: blobs, database rows, queue messages, application logs.
- Commands that return secrets on any cloud, even when they look like reads: `az keyvault secret show`,
  `az account get-access-token`, `az aks get-credentials`, `az acr credential show`,
  `az webapp config appsettings list`, and any `keys list`, `list-keys` or `*-connection-string`.

## Azure Sources

| Need | Source |
|---|---|
| Inventory | `az graph query`: `resources`, `resourcecontainers` |
| Backups | `az graph query`: `recoveryservicesresources` |
| Efficiency | `az graph query`: `advisorresources`, category `Cost` |
| Metrics | `az monitor metrics list` |
| Actual cost | Cost Management query API via `az rest --method post` |

- For reporting, `az rest` is GET only, except POST to `Microsoft.ResourceGraph` and
  `Microsoft.CostManagement/query`.
- The cost API throttles with 429. Keep queries few, and use Advisor when an estimate is enough.

## Keeping Output Small

- Use Resource Graph first: one query spans every subscription. Do not loop over subscriptions.
- `summarize` first. Fetch rows only for the part that matters.
- Project only needed fields and use `-o tsv`. `--first` stops at 1000; page with `--skip-token`.
- Write long output to `.local/` and search it. Do not print it into the conversation.
- A query needed twice becomes a script in `scripts/`.

## Reports

- Write to `reports/` as Markdown, with date, scope and the query used. Do not push them to a remote.
- Each opportunity states the resource count, the estimated saving and what would need to change.

## Maintaining This File

- Keep it short. Rules for one file type go in `.claude/rules/` with `paths:`.
