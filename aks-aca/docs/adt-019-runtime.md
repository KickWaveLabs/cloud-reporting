# ADT_019: runtime for a workload (compute service decision tree)

Summary of the bank's Confluence page *ADT_019 - Runtime de ejecución - Determinar servicio de cómputo
para la ejecución de carga de trabajo* (ADT space, page 677289451). Version 2.0: created 2025-11-18,
published 2025-12-18, in force 2025-12-20; it absorbed ADT_002. Version 1.0 was in force from 2025-09-28.
Transcribed from page screenshots on 2026-11-13. The page is the source of truth; this file is a reading
aid. Quoted labels stay in Spanish so they match the page.

## Decision tree (section 2.2)

Starts from "Necesidad nueva o migración (*)".

| # | Question | Sí | No |
|---|---|---|---|
| 1 | ¿Es desarrollo interno? | Go to 3 | Go to 2 |
| 2 | ¿Software contenerizable? | AKS, App Service or Container Apps (see third-party apps) | Azure Virtual Machine (new VMs need approval from the architecture CoE) |
| 3 | ¿Es monolito? | **App Service** | Next |
| 4 | ¿Funcionalidad relacionada con eventos de servicios de Azure? | Azure Function App | Next |
| 5 | ¿Alta carga transaccional? (carga total > 100 tps) | **AKS** | Next |
| 6 | ¿La aplicación cuenta con una utilización constante la mayor parte del tiempo y más de 15 microservicios? (> 10 h por día) | **AKS** | Next |
| 7 | ¿Se requiere una plataforma multi-aplicaciones? | **AKS** | **Azure Container Apps** |

## Components (section 2.3)

| Component | What it is | Guideline page |
|---|---|---|
| App Service | Deploys applications quickly from source code or a Docker image | Azure Web Apps - WAPP |
| Azure Virtual Machine | General-purpose Windows or Linux VMs | - |
| Azure Kubernetes Service | Kubernetes platform managed by Microsoft | Azure Kubernetes Services - AKSV |
| Azure Function | Serverless platform to run code without managing the platform | Azure Functions - FNCT |
| Azure Container Apps | Serverless platform on top of AKS, managed by Azure, for containerized apps | Azure ContainerApps - ACAS |

## General considerations (2.6.1)

- **App Service is not promoted for new microservice prescriptions.** (It remains the tree's answer for
  internal monoliths.)
- "Trigger Cloud Native" means native integration between cloud services and Azure Functions (for
  example a Blob Storage trigger). Functions are mainly for logic that reacts to Azure service events
  (Blob Storage, Event Grid).
- **The flow is designed for new initiatives.** Existing applications are re-evaluated under 2.6.4.
- Azure VMs (IaaS) are limited to specific use cases and need architecture approval. Always prefer a
  containerized deployment when the technology supports it.
- Entra ID authentication and managed identities between Azure components.
- TLS 1.2+ in transit; encryption at rest with bank-managed keys (BYOK), as per the guideline.
- Log integration with the SIEM according to the owner table's evaluation.

## Container Apps (2.6.2)

- **Cold start:** apps on Container Apps must have retry logic for cold-start events. The first request
  after the app was off can take up to 30 s; after that it answers instantly. To mitigate, the service may
  stay on in office hours or set schedules, **but never 24/7: that costs more than AKS**. If the app or one
  of its consumers cannot tolerate a cold start, use AKS.
- **Critical service** on Container Apps: contact the Unit Solution Architecture Governance to review the
  case. (FAT_001 defines critical as BIA Alta or Muy Alta.)
- **Serverless mode only.** Dedicated infrastructure needs prior consultation with technology
  architecture.
- **Not for long-running work** such as heavy batch or ETL.
- **Resource limit per container instance:** above 4 vCPU / 8 GB RAM, use AKS. (Part of this sentence was
  hidden by a comment popup in the screenshot; 4 / 8 matches the earlier reading of the page.)
- **Standard model:** hybrid serverless to avoid cold starts in office hours (Monday to Friday, 8 am to
  8 pm). An app that supports full serverless may skip this policy and save more, accepting the cold start.
- Private networking; exposed through the API gateway, same model as AKS.
- **Light scheduled batch at night** that runs for hours or needs many resources (CPU/RAM): consult a
  technology architect before prescribing Container Apps.
- Requires the new "Backend Cloud Native Framework" to build and deploy microservices; compatible only
  with Azure Key Vault.

There is **no minimum replica size** in the ADT. The 0.5 vCPU / 1 GiB in the workbook is the model's
default assumption, not an ADT rule.

## Third-party applications (2.6.3)

If containerizable, AKS, Container Apps and App Service are all viable; use the platform that suits the
product best, and check with the vendor which platform is supported and recommended.

## Migration or architecture change (2.6.4)

- On an architecture change or a significant change in usage (for example transactional load beyond the
  tree's thresholds, or extracting functionality to share with other applications), **re-evaluate whether
  the app stays on its platform using the decision tree.** If in doubt, consult technology architecture.
- An app created on Container Apps with "idle" pods (kept warm against cold start) that grows past 15
  microservices: review with a technology architect whether to move to AKS for efficiency.
- Apps that started on App Service for microservices do not have to move to Container Apps or AKS unless
  their architecture or consumption pattern changes.

## References (2.6.5)

Configuration baselines (LBS) per service: App Services (publish for container, publish for code), Azure
Function Apps, Azure Container Apps, Azure Kubernetes Services.

## Notes from the analysis (not in the ADT)

- **Why the tree applies to today's AKS estate:** the tree targets new initiatives (2.6.1), but 2.6.4 asks
  to re-run it when usage changes. The measured usage (low TPS, replicas that never scale, office-hours
  traffic) is the evidence for that re-evaluation.
- **App Service** in the tree is the destination for internal monoliths, not for microservices.
- A workload-profiles environment's Consumption profile allows 4 vCPU / 8 GiB per replica; a
  "Consumption only" environment caps at 2 vCPU / 4 GiB. Kept as a note; the deck and workbook use the
  ADT's 4 / 8.
- The 4-node production minimum for an AKS cluster (2 system + 2 user, 8 vCPU / 32 GB) comes from the
  architecture team's email, not from the ADT or the LBS.
- Where the workbook applies these rules: `Supuestos` (schedule, limits), `Equilibrio` (100 tps, 10 h and
  15-microservice thresholds), `Criterios_tecnicos` and `ACA_por_cluster`.
- Related: [fat-001-monolith-vs-distributed.md](fat-001-monolith-vs-distributed.md),
  [microservice-sizing-criteria.md](microservice-sizing-criteria.md).
