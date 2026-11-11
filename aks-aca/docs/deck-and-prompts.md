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

Inputs, all on the corporate laptop:

- `aks_aca_comparison.xlsx`, regenerated after `aa_k8s_snapshot.py`, opened in Excel and saved (see
  [README](../README.md)).
- The Cost analysis screenshots of [cost-screenshots.md](cost-screenshots.md), Aug–Oct 2026.
- [app-level-comparison.md](app-level-comparison.md).
- The non-compliance level from the security guidelines team (they validate it; the workbook only
  shows configuration dispersion).

Figures below name the workbook sheet they come from. Label every figure from the HPA and deployment
exports with the export date ("configuración al 10-nov-2026").

## Deck 1: Diagnóstico y criterios técnicos (10 slides)

| # | Slide title (conclusion) | Content | Source |
|---|---|---|---|
| 1 | AKS y Container Apps en PROD: diagnóstico, criterios y costo | Title | |
| 2 | Resumen: [N] microservicios en [N] clusters PROD cuestan US$[X]/mes; [Z]% de los clusters tiene 15 microservicios o menos | Three figures and the decisions requested | Resumen; Patrones |
| 3 | Patrón 1: [N] microservicios PROD; [X] capacidades se repiten en hasta [N] dominios | Bars by layer; table of repeated capabilities | Patrones §1 |
| 4 | Patrón 2: [N] de [N] clusters PROD tienen 15 microservicios o menos; los de 3 o menos cuestan US$[X]/mes y usamos [Z]% de los vCPU pagados | Column chart of clusters by microservices; cost of the smallest; CPU used vs allocatable in office hours | Patrones §2; Medido_AKS |
| 5 | Patrón 4: [X]% de los HPA es fijo o nunca pasó de su mínimo en 3 meses | Replicas histogram; HPA behaviour by the export trend vs by 3 months of Container Insights | Patrones §4; Config_K8s |
| 6 | La configuración se fragmenta y mantenerla cuesta horas-hombre: [N] clusters difieren en [M] propiedades, [Z]% no cumple la LBS | Top differing properties; LBS level from the guidelines team; upgrades manual in [N] PROD (qualitative, no amounts) | Config_AKS; Criterios_tecnicos; guidelines team |
| 7 | Qué decide la plataforma, con el dato de PROD | Table: criterion, favours ACA, favours AKS, measured figure; Container Apps limits (consumption cores per environment, 1,000 replicas per revision) | Criterios_tecnicos |
| 8 | Con los mínimos de AKS del banco, ACA cuesta menos en infraestructura; AKS gana si el cluster se comparte o el servicio está activo más del ~60% del tiempo | ADT_019 model (1 replica 8–20, scale to zero); line chart AKS vs ACA | Supuestos; Modelo; Equilibrio |
| 9 | Peticiones y umbrales del ADT_019 no cambian la decisión hasta [X] M peticiones/mes | Heat map microservices × requests; table rule, threshold, model result | Mapa; Equilibrio |
| 10 | Hoja de ruta y decisiones | Quick wins (consolidate small clusters, right-size HPA minimums and requests), ACA where it wins, guidelines baseline; approvals with owner and date | Patrones; Config_K8s |

Annex, outside the count: where each figure comes from (Fuentes, Consultas, Supuestos), and, for
information only, the same figures by business criticality (BIA): the summary's BIA table and the
Candidatos_BIA sheet. The main deck stays technical.

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
A. Costo actual: hojas Resumen y Medido_AKS (costo mensual por cluster PROD, total, nodos, tamaños de VM).
B. Insights PROD: Medido_AKS (vCPU asignables vs usados en oficina, pico), Workloads (workloads por cluster) y Config_AKS (propiedades con más clusters distintos al valor más frecuente).
C. Comparación: Supuestos (valores de la columna C), Modelo (costo con 1, 5, 10, 15 y 20 microservicios), Equilibrio (puntos de equilibrio y tabla de umbrales del ADT_019) y Mapa (fila de 15 microservicios).
D. Supuestos que más afectan el resultado: lista los de la hoja Supuestos marcados como supuesto y su valor.
Marca cada cifra como medida, derivada, estimada o dato del proveedor según la hoja Fuentes.
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
Inserta como imagen el gráfico de la hoja Modelo (lo pego yo) y a la derecha tres cifras: costo ACA por microservicio al mes,
costo del cluster AKS mínimo a precio de lista y con el descuento medido. Pie de página: "Estimado con precios de lista de Azure
(Retail Prices API) y supuestos de la hoja Supuestos; costo AKS efectivo derivado de Cost Management, ago-oct 2026".
```

ADT thresholds:

```text
Agrega una diapositiva con una tabla de cuatro filas: regla del ADT_019, umbral del ADT, resultado del modelo y lectura,
copiando la tabla de la hoja Equilibrio. Título: "[Lectura de la fila de 15 microservicios]".
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
