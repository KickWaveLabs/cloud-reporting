# Application-level comparison: AKS (with Istio) vs Container Apps

What changes for a team that builds and runs the same application on each platform. Order: first the
application (it must behave the same for the business), then the service mesh, where the platforms
differ the most.

Context that applies to both, from ADT_019 and the bank's model:

- Exposure is north-south through API Management (BS and Core). Every call, including calls between
  microservices of the same layer, goes through APIM. Same model on AKS and ACA.
- Identity: Entra ID and managed identities. TLS 1.2+ in transit; encryption at rest with
  bank-managed keys (BYOK).
- ACA requires the bank's Backend Cloud Native Framework, which reads configuration and secrets only
  from Azure Key Vault.

Rows on fast-moving ACA features were checked on Microsoft Learn on 2026-11-10 (sources at the end).
Check them again on the day the deck is built and note the date in the annex.

## 1. Application parity

| Topic | AKS | Container Apps | Effort to move AKS → ACA |
|---|---|---|---|
| Build | Same OCI image, same pipeline up to the push to ACR | Same | None |
| Runtime limits | Any pod size; batch, CronJobs, long processes | Max 4 vCPU / 8 GiB per replica (ADT_019); jobs for short tasks, no batch/ETL | Re-size or keep on AKS |
| Endpoints | Service + Ingress/Gateway behind APIM | App ingress (internal) behind APIM | Map each Service/Ingress to app ingress; APIM backend URL changes |
| Calls between microservices | Through APIM (north-south) | Same | None |
| Data access | Private endpoints from the VNet; workload identity | Private endpoints from the environment's VNet; managed identity | Swap workload identity for app managed identity |
| Producers and consumers (Event Hubs, Service Bus, Kafka) | Pods always running, or KEDA scaler | KEDA scale rules; scale to zero outside office hours | Consumers must tolerate cold start and lag at 8:00; long-lived connections reconnect |
| Cold start | No (pods run 24/7) | Up to ~30 s on the first request after scale to zero (ADT_019) | Retry logic in callers |
| Scaling | HPA/KEDA on node capacity the team sizes | KEDA rules (HTTP, queue, cron), platform capacity | Rewrite HPA as scale rules |
| Health probes | liveness/readiness/startup | Same three probe types | Copy |
| Observability | Container Insights, Prometheus | Log Analytics; managed OpenTelemetry agent (GA, one per environment: App Insights logs and traces, OTLP, Datadog and others) | Point exporters to the new sink |

## 2. Configuration

| Topic | AKS | Container Apps |
|---|---|---|
| Plain configuration | ConfigMaps (env or mounted files) | Environment variables in the app YAML; no ConfigMap object |
| Secrets | Secrets / Key Vault CSI driver | App secrets that reference Key Vault through managed identity; secret volumes |
| Files of configuration | ConfigMap volume | Secret volume, or the framework reads Key Vault |
| Per-environment values | Kustomize overlays / Helm values | One YAML or Bicep parameter file per environment |

Effort: every ConfigMap becomes variables in the app YAML or entries in Key Vault, per application
and environment.

## 3. Release strategies and GitOps

| Topic | AKS + Istio | Container Apps |
|---|---|---|
| Canary by weight | VirtualService weights; automated with Argo Rollouts or Flagger | Revisions with traffic weights, set by the pipeline |
| Canary by header or user | VirtualService match on headers | Not available: traffic split is by weight only; a revision label gives a separate URL |
| Blue/green | Two Deployments + VirtualService switch | Two revisions, move 100% of traffic |
| Automatic rollback on metrics | Argo Rollouts / Flagger analysis | Not native; the pipeline has to check and roll back |
| Traffic mirroring | Istio mirror | Not available |
| GitOps | Argo CD / Flux reconcile the cluster from Git | No in-platform reconciler: the pipeline applies the Microsoft YAML or Bicep; drift is not corrected automatically |

## 4. Service mesh: security

| Capability | AKS + Istio add-on | Container Apps native |
|---|---|---|
| mTLS between services | PeerAuthentication STRICT, per namespace or workload | Peer-to-peer encryption (GA): one setting for the whole environment, off by default; authenticates apps but does not authorize between them |
| Workload identity in the certificate | SPIFFE ID per service account | Not exposed; Dapr service invocation adds mTLS, but Dapr access-control policies are not supported in ACA |
| Certificate management | istiod CA, or plug-in CA from Key Vault; rotation by the mesh | Managed by the platform; no control over the CA |
| Authorization between services | AuthorizationPolicy by identity, path, method | Not available at that level; rely on APIM policies and network rules |
| Egress control | Egress gateway, ServiceEntry | Environment network (UDR, firewall) |

## 5. Service mesh: resilience

| Capability | AKS + Istio | Container Apps native |
|---|---|---|
| Load balancing | Round robin, least request, consistent hash | Platform balancing between replicas |
| Timeout and retry | VirtualService per route | Service discovery resiliency policies (preview, one per app); otherwise APIM policies |
| Circuit breaker (internal) | DestinationRule connectionPool + outlierDetection | Resiliency policy circuit breaker (preview): consecutive errors, interval, max ejection % |
| Circuit breaker (external) | ServiceEntry + DestinationRule | APIM backend circuit breaker, or in code |
| Outlier detection | Ejects failing endpoints | Covered by the same circuit breaker (preview): ejects failing replicas from the pool |
| Fault injection for tests | Yes | No |

Because calls go through APIM, part of the timeout, retry and circuit-breaker policy can live in APIM
on both platforms. That is the common layer to unify; the mesh adds the internal hop (APIM to pod)
and identity-based authorization.

## Key message

For the business both platforms run the same application. For IT, running both means building twice:
two ways to configure, release, secure and make the services resilient, and later two migration paths
for multicloud. The cost comparison in the workbook says where ACA is cheaper; this table says what it
costs IT to keep both.

## Sources (checked 2026-11-10)

- OpenTelemetry agent: https://learn.microsoft.com/azure/container-apps/opentelemetry-agents (GA since
  May 2025; the page description still says preview).
- Traffic splitting and labels: https://learn.microsoft.com/azure/container-apps/traffic-splitting,
  https://learn.microsoft.com/azure/container-apps/revisions
- Peer-to-peer encryption: https://learn.microsoft.com/azure/container-apps/ingress-environment-configuration
- Dapr: https://learn.microsoft.com/azure/container-apps/dapr-overview
- Resiliency and circuit breaker: https://learn.microsoft.com/azure/container-apps/service-discovery-resiliency
- Scale defaults (polling 30 s, cooldown 300 s, min replicas 0): https://learn.microsoft.com/azure/container-apps/scale-app
- Billing and free grant: https://learn.microsoft.com/azure/container-apps/billing
