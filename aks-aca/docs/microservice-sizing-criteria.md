# Criteria to define microservice size

Summary of the bank's Confluence page *Criterios para definir tamaño de microservicios* (ADT space,
author Luis Alberto Rojas Vasquez), transcribed from screenshots on 2026-11-13. The page is the source
of truth; this file is a reading aid.

"Size" here means **functional granularity** (which features go together in one microservice), not
CPU or memory per replica. There is no single formula; the page lists criteria to start from a
domain and split it.

| # | Criterion | Rule | Page example |
|---|---|---|---|
| 1 | Group by functionality | Start from the business domain, modelled with DDD. At the bank (BIAN): **one Service Domain = one microservice** as the starting point; then check the other criteria to split it. | Service Domain with features F1–F6 → one microservice with F1–F6 |
| 2 | Components that change at different rates | Group features by change frequency (weekly, monthly, every 6 months, yearly). | F2 monthly, F4/F6 every 6 months, F1/F3/F5 yearly → three microservices |
| 3 | Keep interdependent transactions together | Features in the same transaction stay in one microservice, to avoid distributed transactions. | F1 + F3 transactional → {F1, F3} and {F2, F4, F5, F6} |
| 4 | Keep features that change together | If a change in one feature forces changes in others, keep them in one microservice. | A change in F1 implies F3 and F5 → {F1, F3, F5} and {F2, F4, F6} |
| 5 | Features that scale differently | A feature that must scale far more than the rest (for example 50x) goes into its own microservice. | F1/F3 below 5x, F5/F6 above 5x, F2/F4 above 100x → three microservices |
| 6 | Different security needs | Features with specific security mechanisms (for example PCI) are separated from those with standard security. | F2/F4/F6 PCI → {F1, F3, F5} and {F2, F4, F6} |
| 7 | Too many teams on the same code | When several teams touch one codebase (dependencies, rework, waiting), split by team. | Teams A, B, C → {F1, F3}, {F2, F4, F6}, {F5} |

The screenshots end after criterion 7; check the page for anything below it.

## How the aks-aca work uses it

- Criterion 5 is the only one the Azure data can test. A deployment whose replicas never leave the
  HPA minimum shows no different scaling need, so on its own it does not justify a separate
  microservice (nor a dedicated cluster). This supports the "does it scale?" criterion.
- Criteria 2, 3, 4, 6 and 7 depend on code, transactions and teams; the read-only data cannot see them.
  Leave them out of the deck's figures.
- The runtime question (AKS, Container Apps or App Service) comes after sizing: the same microservice
  can run on either platform. CPU and memory limits per replica come from ADT_019, not from this page.
- Related: [fat-001-monolith-vs-distributed.md](fat-001-monolith-vs-distributed.md) (when to split at all).
