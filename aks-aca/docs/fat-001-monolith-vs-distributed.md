# FAT_001: when to use a monolith or a distributed architecture

Summary of the bank's Confluence page *FAT_001 - Determinar cuándo usar Monolito o Arquitectura
Distribuida* (ADT space, last updated 31 March 2026), transcribed from screenshots on 2026-11-13.
Quoted labels stay in Spanish so they match the page. The page is the source of truth; this file is a
reading aid.

## Purpose

A decision tree that prescribes monolith, miniservices or microservices. Key factors: simplicity vs
technical complexity, time to market, team autonomy, and total cost (infrastructure plus the people who
build and maintain it). Principle: **"Monolito primero"** (monolith first): start with one deployable
unit, ideally a modular monolith split by domain, and only split later. Splitting too early is costly
(APIs, contracts, data, coordination).

## Characteristics of each style

| Aspect | Monolith | Miniservice | Microservice |
|---|---|---|---|
| Cost | Low | Medium | High |
| Deployment unit | 1 artifact / 1 release | Few units, each fairly large | Many small units, independent releases |
| Simplicity | High | Medium | Low |
| Granularity | - | Coarse | Fine |
| Typical boundary | Internal modules / layers | Large domains or subdomains | Business capability / bounded context |
| Communication | In-process | Network (HTTP/gRPC/events) | Network (HTTP/gRPC/events) |
| Data | Often a shared database | Mixed (shared or per service) | Ownership per service |
| Scaling | Scale everything | Scale by block | Scale per service |
| Operational complexity | Low | Medium | High |
| Latency | Low | Medium | Variable (depends on services in the flow) |
| Coupling | High without modularity | Medium (contracts between services) | Low with good boundaries; rises with many dependencies |
| Observability | Basic metrics/logs; tracing optional | Metrics, alerts, traces across services; SLOs on critical services | Distributed metrics and tracing "mandatory"; SLOs per service and journey |
| Testability | High | Medium | Medium |
| Elasticity | Low | Medium | High |
| Fault tolerance | Low | Medium | High |

Spectrum figure: monolith (one runtime, one data store) → miniservices (a few services sharing a
runtime and data store per block) → microservices (one runtime and data store each). Toward the monolith:
more data integrity, less complexity, **lower cost**. Toward microservices: more development agility,
deployment flexibility and precise scaling, **higher cost**.

### DDD and BIAN equivalences

| DDD | BIAN | Note |
|---|---|---|
| Domain | Service Landscape | Whole business ecosystem |
| Subdomain | Business Domain | Group of related business capabilities |
| Bounded Context | Service Domain (SD) | They meet at the data-model and ubiquitous-language boundary |
| Entity / Aggregate | Business Object | Core logic inside the SD |

| Style | BIAN equivalent |
|---|---|
| Monolith | One or more Business Domains / a group of Service Domains |
| Miniservices | A group of Service Domains inside one Business Domain, justified by data affinity or critical latency (transactionality) |
| Microservices | One Service Domain |

All three sit at the Solution Building Block (SBB) level; one SBB may combine styles.

## Decision tree

Questions in order; the first "Sí" that leads out of the chain decides. Section 2.6.2 of the page
explains each question.

| # | Question (Spanish, as on the page) | Sí | No |
|---|---|---|---|
| 1 | ¿Es un MVP o la solución abarca un dominio sencillo? | Monolito | Next |
| 2 | ¿Es requerida una alta cadencia de despliegue en paralelo, o existen módulos con una cadencia de despliegue significativamente diferente al resto? | Go to 8 | Next |
| 3 | ¿Es una aplicación crítica en la que resulta imprescindible garantizar que un fallo en una funcionalidad no afecte al resto, con aislamiento completo por dominio? | Go to 8 | Next |
| 4 | ¿Existen funcionalidades que, por requisitos específicos (seguridad, cumplimiento normativo u otras condiciones críticas), deban mantenerse estrictamente aisladas? | Go to 8 | Next |
| 5 | ¿Existen funcionalidades con niveles de concurrencia significativamente distintos que requieren estrategias de escalado específicas? | Go to 8 | Next |
| 6 | ¿Existen flujos críticos que requieran consistencia transaccional estricta (ACID) entre múltiples funcionalidades y no toleran consistencia eventual? | Monolito | Next |
| 7 | ¿Es realmente indispensable utilizar distintos lenguajes o frameworks? | Go to 8 | Monolito |
| 8 | ¿Viene de un monolito o es una aplicación nueva? | Miniservicio | Microservicio |

Question 8 is read from a rotated photo of the diagram: "Sí" leads to miniservice, "No" to microservice.
Check the original image before quoting the branch.

## How each question is judged (section 2.6.2)

1. **MVP and simple domain.** An MVP delivers enough value with the least effort to validate hypotheses,
   measure real users, cut waste and speed time to market. A simple domain: CRUD without business logic,
   few rules and external dependencies, no DDD/CQRS/event-driven patterns, infrequent changes.
2. **Deployment frequency (DORA).** Elite: on demand, several times a day. High: daily to weekly.
   Medium: weekly to monthly. Low: less than monthly. Elite or high → distributed; medium or low →
   monolith. Example: 5 developers shipping 10 features a month fit a monolith; 30 features a month
   needs about 20 developers in 3 teams and a distributed architecture. Uneven change rates also count
   (example over six months: customers 1–3 changes, sales 30–50, suppliers 1–3, payments 20–40; split
   sales and payments).
3. **Fault isolation.** Essential for business-critical apps, high change frequency, 24/7 availability,
   and compliance. **At the bank, an application is critical when its BIA is Alta or Muy Alta (recorded
   in CVT).** Small or medium systems with few changes and low fault impact → monolith; large, critical,
   frequently changed, high fault risk → miniservices or microservices.
4. **Differentiated security** (PCI DSS, ISO 27001, sensitive or regulated data). Same requirements for
   all modules → monolith. Only some modules need strict controls → start with a monolith, evolve to
   miniservices, and microservices only as the last stage (example: isolate a PCI DSS payments module).
5. **Differentiated scalability.** Uniform load and controlled growth → monolith. Some modules grow 10x
   more than others or have significant peaks → distributed (example: on payday, interbank transfers and
   service payments reach 30x the traffic of other modules). The page warns to weigh total cost:
   database per service, message brokers, repositories, pipelines, per-service maintenance and
   distributed security.
6. **ACID transactions.** Strict ACID without eventual consistency → monolith. ACID that fits inside one
   miniservice → miniservice. Microservices are not recommended when strict ACID is needed across services.
7. **Different languages or frameworks.** Uniform stack → monolith. Specialised technology per large
   domain, separating 2–5 blocks (for example AI in Python, payments in Java/Spring Boot for PCI DSS,
   reports in Node.js) → miniservices. Microservices allow any technology per service, at the cost of
   orchestration, monitoring and security across heterogeneous services. (The screenshots end before
   "¿Cuándo optar por microservicios?" in this section.)

General consideration (2.6.1): to move from a monolith to a distributed architecture, use the Strangler
pattern to extract features gradually.

## How the aks-aca work uses it

- The corporate deck keeps monolith vs miniservice vs microservice out (decision of 2026-11-12): it is a
  high-level proposal about runtimes, not about architecture style.
- Slide 1 (context) can borrow the spectrum figure: each style needs a runtime, and finer granularity
  means higher cost.
- The BIA rule (Alta or Muy Alta = critical) matches how the workbook flags critical clusters.
- Question 5 (10x and 30x peaks) supports the "does it scale?" criterion: a service whose replicas never
  leave the minimum has no differentiated scaling need.
