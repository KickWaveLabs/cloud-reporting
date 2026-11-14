# AKS vs Container Apps decks: outline and Copilot prompts

Two decks, as the manager asked:

1. **Diagnóstico y criterios técnicos**: what PROD looks like today (patterns 1, 2 and 4), what it costs,
   the technical criteria to choose a platform with the measured figure behind each one, and the cost
   comparison under the ADT_019 rules.
2. **Frente estratégico**: security, resilience and multicloud with Istio vs Container Apps, as the
   architecture lead's email asked. Key message: for the business both platforms are equivalent; for IT
   two platforms double the effort.

Both follow [presenting-findings.md](../../docs/presenting-findings.md): evidence first, then the
comparison and the decision.

## Plan update (2026-11-13)

Feedback on the first presentation: the forum wants a corporate deck, readable by a non-technical audience
that knows TPS, CPU, RAM and time to market. The technical deck stays as backup for architects.

- Data: one month (the last full month, October 2026) for every Azure source; configuration from the
  platform team's exports of the export day.
- Only APPLICATION clusters count (categories in [data-guide.md](data-guide.md)); infrastructure clusters are
  mentioned as excluded and clusters without applications are a finding of their own.
- Pattern 3 is dropped. Pattern 4 shows "does not scale" as a total, split into "stays at its minimum"
  (load) and "min = max" (configuration), plus the time without scaling.
- Corporate deck (slides 1–5 by the user; slide 6, effort and migration estimate, by the manager's team):
  1. Context: workloads and the runtimes for each (no monolith vs microservice debate).
  2. The runtime estate today: AKS, Container Apps, App Service.
  3. Criteria to sort the estate (ADT_019 tree; see [adt-019-runtime.md](adt-019-runtime.md)), each with the
     PROD figure: TPS, hours with traffic, microservices, size per instance, scales or not, batch, cold start.
  4. Theoretical distribution after applying the criteria.
  5. Cost optimization: logic and rates, monthly and annual, with and without discount; growth scenario;
     operating effort as a non-monetary benefit.
- Traffic classes (office, extended, 24x7) stay in the technical deck until another architect validates them.

## Plan (2026-11-11)

The two decks became one presentation in four sections. Status:

| Section | Slide | Status | Source |
|---|---|---|---|
| Diagnosis | Summary | Done | Summary; Measured_AKS |
| | Pattern 1: microservices per layer, repeated capabilities | Done | deployment_snapshot, capability_domains |
| | Pattern 2: 99 of 132 clusters with 15 microservices or fewer | Done | cluster_snapshot + cost_by_resource (last full month) |
| | Pattern 3: apps on system nodes, clusters without apps | New, optional | cluster_snapshot + aks_node_pools + aks_workloads_monthly |
| | Pattern 4: HPAs at their minimum | Done | hpa_snapshot |
| | BIA, information only (slide 5 of the current deck) | Done | cluster_snapshot.bia + cost_by_resource |
| | Capacity health: low use is not a lack of capacity (hours with unschedulable pods, OOM kills, CPU vs memory) | Next | aks_autoscaler_monthly, aks_oomkilled_monthly, aks_usage_monthly |
| | Fragmented configuration and clusters under the 4-node production minimum | Later | aks_clusters, aks_node_pools; LBS level from the guidelines team |
| Technical criteria | Criteria AKS vs ACA with the PROD figure, including backend latency vs a cold start | Next | Technical_criteria; each figure from its CSV; apim_latency_monthly |
| | Cost under the ADT_019 rules | Next | Assumptions, Model, Break_even (estimate, prices.csv) |
| Strategic front | Security, resilience, multicloud and key message | If time allows | app-level-comparison.md |
| Proposal | Per-cluster comparison, traffic profile, HPA minimums, quick wins | Last | ACA_by_cluster, Traffic_profile, hpa_snapshot |
| Annex | Sources and method | Later | Sources, Queries |

The deck has no cover slide: it opens with the summary. Every diagnosis figure can be rebuilt with a pivot table on the CSV
named in the Source column; the cost slide is a model estimate and says so.

Inputs, all on the corporate laptop:

- `aks_aca_comparison.xlsx`, regenerated after `aa_k8s_snapshot.py`, opened in Excel and saved (see
  [README](../README.md)).
- The Cost analysis screenshots of [cost-screenshots.md](cost-screenshots.md), Aug–Oct 2026.
- [app-level-comparison.md](app-level-comparison.md).
- The non-compliance level from the security guidelines team (they validate it; the workbook only
  shows configuration dispersion).

Figures below name the workbook sheet they come from. Label every figure from the HPA and deployment
exports with the export date ("configuración al 10-nov-2026").

## Deck 1: Diagnóstico y criterios técnicos (10 slides, no cover)

Order of the deck as built on 2026-11-12: it opens with the summary and keeps BIA right after the patterns,
labelled as information only.

| # | Slide title (conclusion) | Content | Source |
|---|---|---|---|
| 1 | AKS PROD combina alto costo fijo con baja utilización de CPU | Eight figures (clusters, monthly cost, average per cluster, clusters with 4 nodes or fewer, CPU and memory in office hours, clusters with 15 microservices or fewer, HPAs at their minimum); use table (average, p95, maximum); minimum AKS cluster vs one microservice on ACA | Summary; Measured_AKS; Patterns |
| 2 | Patrón 1: [N] microservicios; capacidades repetidas en hasta [N] dominios | Bars by layer; table of repeated capabilities | Patterns §1 |
| 3 | Patrón 2: [N] de [N] clusters tienen 15 microservicios o menos | Column chart of clusters by microservices; cost of those with 3 or fewer; clusters missing from the export as a note | Patterns §2 |
| 3b | Patrón 3 (new, optional): [N] clusters corren sus apps en los nodos de sistema | Clusters without a user pool, without apps, under the 4-node minimum, with one system node; list of clusters without a user pool | Patterns §3 |
| 4 | Patrón 4: [X]% de los HPA se queda en su mínimo durante 3 meses | Replicas histogram; HPA behaviour over 3 months of Container Insights; inflated minimums | Patterns §4 |
| 5 | Anexo BIA: la criticidad prioriza; la decisión de plataforma sigue siendo técnica | Monthly cost by BIA level; share of MUY ALTA; clusters with high BIA, without BIA, total; how to use BIA. Information only | Summary (BIA table); cluster_snapshot.bia |
| 6 | La configuración se fragmenta y mantenerla cuesta horas-hombre: [N] clusters difieren en [M] propiedades, [Z]% no cumple la LBS | Top differing properties; LBS level from the guidelines team; upgrades manual in [N] PROD (qualitative, no amounts) | Config_AKS; Technical_criteria; guidelines team |
| 7 | Cómo decidimos si un AKS puede pasar a Container Apps | The seven decision rules of ADT_019 with today's PROD figure next to each; the full criteria table (what favours each platform, limits, sources) stays in the workbook | Technical_criteria; Summary (BIA) |
| 8 | Con los mínimos de AKS del banco, ACA cuesta menos en infraestructura; AKS gana si el cluster se comparte o el servicio está activo más del ~60% del tiempo | ADT_019 model (1 replica 8–20, scale to zero); line chart AKS vs ACA | Assumptions; Model; Break_even |
| 9 | Peticiones y umbrales del ADT_019 no cambian la decisión hasta [X] M peticiones/mes | Heat map microservices × requests; table rule, threshold, model result | Map; Break_even |
| 10 | Hoja de ruta y decisiones | Quick wins (consolidate small clusters, right-size HPA minimums and requests), ACA where it wins, guidelines baseline; approvals with owner and date | Patterns; Config_K8s; ACA_by_cluster |

Annex, outside the count: where each figure comes from (Sources, Queries, Assumptions). The manager
asked to keep BIA out of the technical argument: slide 5 says it is information only, and it can move to
the annex without changing anything else.

## Deck 2: Frente estratégico (5 slides)

| # | Slide title (conclusion) | Content | Source |
|---|---|---|---|
| 1 | AKS con Istio y Container Apps: seguridad, resiliencia y multicloud | Title | |
| 2 | La aplicación funciona igual; lo que cambia es configuración, releases y GitOps | Build, endpoints, data; ConfigMaps → YAML/Key Vault; canary; Argo CD → pipeline | app-level-comparison §1–3 |
| 3 | Seguridad: Istio da identidad SPIFFE, mTLS STRICT y autorización por servicio; ACA no | Security table, certificate management | app-level-comparison §4 |
| 4 | Resiliencia: timeout, retry, circuit breaker y outlier detection | Resilience table; what APIM covers on both | app-level-comparison §5 |
| 5 | Multicloud y mensaje clave: para el negocio son equivalentes; para TI, dos plataformas duplican el esfuerzo | AKS carries the same model to other clouds, ACA is Azure only; delivery and migration model ACA → AKS (to complete with the platform team) | app-level-comparison, key message |

## Copilot prompts

Run them in order. Check every number Copilot writes against the workbook.

### 1. Brief from the workbook (Copilot chat, workbook attached)

```text
Eres un arquitecto cloud preparando una presentación para la gerencia del banco que compara AKS y Azure Container Apps.
Usa SOLO las cifras del archivo adjunto aks_aca_comparison.xlsx. Si una cifra no está, escribe [DATO PENDIENTE].
No redondees de forma distinta a la del archivo y no calcules cifras nuevas salvo que te lo pida.

Arma un resumen en cuatro secciones, con una tabla por sección (indicador, valor, tipo de dato, hoja y fila de origen):
A. Costo actual: hojas Summary y Measured_AKS (costo mensual por cluster PROD, total, nodos, tamaños de VM).
B. Insights PROD: Measured_AKS (vCPU asignables vs usados en oficina, pico), Workloads (workloads por cluster) y Config_AKS (propiedades con más clusters distintos al valor más frecuente).
C. Comparación: Assumptions (valores de la columna C), Model (costo con 1, 5, 10, 15 y 20 microservicios), Break_even (puntos de equilibrio y tabla de umbrales del ADT_019) y Map (fila de 15 microservicios).
D. Supuestos que más afectan el resultado: lista los de la hoja Assumptions marcados como supuesto y su valor.
Marca cada cifra como medida, derivada, estimada o dato del proveedor según la hoja Sources.
```

Save the answer as a Word document (for example `brief_aks_aca.docx`) after checking it.

### 2. Deck from the brief (PowerPoint, Copilot > Create a presentation from a file)

```text
Crea una presentación en español a partir de brief_aks_aca.docx con este orden:
1) Portada y resumen ejecutivo, 2) Costo actual de AKS, 3) Insights de producción: uso de capacidad, un cluster por aplicación y fragmentación de la configuración,
4) Comparación de costo AKS vs Container Apps con el punto de equilibrio, 5) Frente estratégico: aplicación, configuración, releases, seguridad y resiliencia,
6) Hoja de ruta y decisiones, 7) Anexo de fuentes.
Cada título debe ser la conclusión con su cifra, no un tema (ejemplo: "Usamos 12% de los vCPU que pagamos en horario de oficina").
Una idea por diapositiva. Usa solo cifras del documento. En las notas del orador, indica la hoja y fila de origen de cada cifra y si es medida o estimada.
Deja en la sección de costo actual dos diapositivas con un espacio para pegar capturas de Cost Management.
```

### 3. Slide by slide (PowerPoint, Copilot on the open deck)

Break-even:

```text
Agrega una diapositiva titulada "ACA es más barato hasta [N] microservicios; desde ahí conviene AKS".
Inserta como imagen el gráfico de la hoja Model (lo pego yo) y a la derecha tres cifras: costo ACA por microservicio al mes,
costo del cluster AKS mínimo a precio de lista y con el descuento medido. Pie de página: "Estimado con precios de lista de Azure
(Retail Prices API) y supuestos de la hoja Assumptions; costo AKS efectivo derivado de Cost Management, ago-oct 2026".
```

ADT thresholds:

```text
Agrega una diapositiva con una tabla de cuatro filas: regla del ADT_019, umbral del ADT, resultado del modelo y lectura,
copiando la tabla de la hoja Break_even. Título: "[Lectura de la fila de 15 microservicios]".
Nota del orador: los demás criterios del ADT (4 vCPU / 8 GiB por réplica, batch, cold start, ambiente dedicado) no son de costo
y pueden descartar Container Apps aunque sea más barato.
```

Configuration problem:

```text
Agrega una diapositiva titulada "La configuración base se fragmenta: [N] clusters PROD difieren en [M] propiedades".
Muestra un gráfico de barras horizontales con las cinco propiedades de Config_AKS con más clusters PROD distintos al valor más frecuente.
Debajo, una línea: "Los clusters no adoptan los cambios de la línea base posteriores a su creación; los equipos no tienen tiempo para mantenerlos al día".
Nota del orador: el nivel de incumplimiento de la LBS lo valida el equipo de lineamientos; esta lámina muestra dispersión, no cumplimiento.
```

Maintenance effort:

```text
Agrega una diapositiva titulada "Mantener [N] clusters al día requiere horas-hombre de los equipos".
Dos columnas. AKS: upgrades de versión menor de Kubernetes manuales en [N] clusters PROD; actualización automática
de nodos según el lineamiento, hoy en [N] de [N] clusters PROD; revisiones de Istio en [N] clusters; cada upgrade
se prueba con las aplicaciones. Container Apps: la plataforma actualiza todo, sin versiones visibles.
No pongas montos: el esfuerzo se expresa en horas-hombre y lo estima el equipo de plataforma.
Nota del orador: el lineamiento de actualización automática de nodos reduce el esfuerzo, pero la versión menor
de Kubernetes y las revisiones de Istio siguen siendo manuales.
```

Strategic front:

```text
Agrega una diapositiva titulada "Para el negocio son equivalentes; para TI, dos plataformas duplican el esfuerzo".
Dos columnas, AKS con Istio y Container Apps, con cuatro filas: seguridad (SPIFFE, mTLS, certificados), resiliencia
(balanceo, circuit breaker interno y externo, timeout/retry, outlier detection), releases (canary, GitOps) y configuración.
Usa solo el contenido del documento adjunto app-level-comparison. No agregues capacidades que no estén ahí.
```

### 4. Final check (PowerPoint, Copilot on the finished deck)

```text
Lista en una tabla todas las cifras de la presentación: número de diapositiva, cifra, texto donde aparece.
Marca las que no tienen fuente en las notas del orador.
```

Compare that table with the workbook. Fix by hand; do not ask Copilot to "correct" numbers.

## Before presenting

- Every current-cost figure matches its Cost analysis screenshot for Aug–Oct 2026.
- Every model figure says "estimado" and its assumptions are in the annex.
- The **Verify** rows of app-level-comparison.md were checked on Microsoft Learn, with the date.
- The ACA environment meters (management hour, private endpoint) were confirmed for the bank's setup.
- The LBS non-compliance figure comes from the guidelines team, with its date.
