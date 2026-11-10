# AKS vs Container Apps deck: outline and Copilot prompts

One deck in the order of [presenting-findings.md](../../docs/presenting-findings.md): current cost,
production insights (including the configuration problem), comparison and decision, roadmap. Cost is
the main focus, as requested.

Inputs, all on the corporate laptop:

- `aks_aca_comparison.xlsx`, regenerated, opened in Excel and saved (see [README](../README.md)).
- The Cost analysis screenshots of [cost-screenshots.md](cost-screenshots.md), Aug–Oct 2026.
- [app-level-comparison.md](app-level-comparison.md), with the **Verify** rows checked that day.
- The non-compliance level from the security guidelines team (they validate it; the workbook only
  shows configuration dispersion).

Figures below name the workbook sheet they come from.

## Outline

| # | Slide title (conclusion) | Content | Source |
|---|---|---|---|
| 0 | AKS y Container Apps: costo, esfuerzo y decisión | Title | |
| 1 | Resumen: [N] clusters PROD cuestan US$[X]/mes; para apps de menos de 15 microservicios, ACA cuesta US$[Y]/mes | Three figures and the decisions requested | Resumen |
| **A** | **Costo actual (evidencia)** | | |
| 2 | AKS PROD cuesta US$[X]/mes en [N] clusters, uno por aplicación | Stacked column by meter category, Aug–Oct | Screenshots 1, 3; Medido_AKS |
| 3 | Un cluster mínimo de 4 nodos cuesta US$[X]/mes aunque la app use [Z]% de la CPU | One cluster: cost by meter + CPU used vs allocatable | Screenshot 2; Medido_AKS |
| **B** | **Insights PROD** | | |
| 4 | Usamos [Z]% de los vCPU que pagamos en horario de oficina | Bar per cluster: allocatable vs used (office average, peak) | Medido_AKS |
| 5 | Cada aplicación tiene su propio cluster: [N] clusters para [N] aplicaciones | One app per cluster vs one app per namespace | Medido_AKS; Workloads |
| 6 | La configuración base se fragmenta: [N] clusters PROD difieren en [M] propiedades | Top properties by clusters with a different value | Config_AKS |
| 7 | [Z]% de los clusters no cumple la LBS (dato del equipo de lineamientos) | Non-compliance level, why: changes after creation are not adopted | Guidelines team; Config_AKS |
| 7b | Mantener [N] clusters al día requiere horas-hombre de los equipos | Qualitative, no amounts: Kubernetes minor upgrades are manual (`upgrade_channel` none in [N] PROD); node auto-upgrade guideline in rollout ([N] of [N] PROD with `NodeImage`); Istio revisions in [N] clusters; in ACA the platform does all of it | Config_AKS; node auto-upgrade guideline |
| **C** | **Comparación de costo (producción)** | | |
| 8 | Con el modelo del ADT, ACA cuesta US$[X] por microservicio al mes | How ACA bills: 1 replica 8–20, scale to zero, requests | Supuestos; Modelo |
| 9 | ACA es más barato hasta [N*] microservicios; desde ahí conviene AKS | Line chart AKS (list, effective) vs ACA (ADT, 24/7) | Modelo chart; Equilibrio |
| 10 | Peticiones y réplicas no cambian la decisión hasta [X] M peticiones/mes | Heat map microservices × requests | Mapa |
| 11 | Los umbrales del ADT_019 frente al costo | Table: rule, ADT threshold, model result, reading | Equilibrio, ADT table |
| 12 | Qué mueve el resultado | Sensitivity: vCPU per replica, share of office time with traffic, 24/7 vs office hours | Supuestos (rerun with 2–3 values) |
| **D** | **Frente estratégico** | | |
| 13 | La aplicación funciona igual en las dos plataformas | Build, endpoints, data, producers and consumers | app-level-comparison §1 |
| 14 | Pasar de AKS a ACA cambia configuración, releases y GitOps | ConfigMaps → YAML/Key Vault; canary; Argo CD → pipeline | app-level-comparison §2–3 |
| 15 | Seguridad: Istio da identidad SPIFFE, mTLS STRICT y autorización por servicio; ACA no | Security table | app-level-comparison §4 |
| 16 | Resiliencia: timeout, retry, circuit breaker y outlier detection | Resilience table; what APIM covers on both | app-level-comparison §5 |
| 17 | Para el negocio son equivalentes; para TI, dos plataformas duplican el esfuerzo | Key message | app-level-comparison, key message |
| **E** | **Cierre** | | |
| 18 | Hoja de ruta | Quick wins (right-size, shared clusters), ACA where it wins, guidelines baseline | |
| 19 | Decisiones que se necesitan | Approvals with owner and date | |
| 20 | Anexo: de dónde sale cada cifra | Measured, derived, estimated, vendor data; ADT_019 rules | Fuentes; Consultas; Supuestos |

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
