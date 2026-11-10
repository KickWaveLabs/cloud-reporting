# Backup deck: outline and Copilot prompts

One deck in the order of [presenting-findings.md](../../docs/presenting-findings.md): cost evidence,
production insights, non-production insights, production-only comparison, decisions.

Inputs, all on the corporate laptop:

- `backup_insights.xlsx`, regenerated, opened in Excel and saved (see [README](../README.md)).
- The Cost analysis screenshots of [cost-screenshots.md](cost-screenshots.md), Aug–Oct 2026.
- A copy of the workbook with the production comparison of [prod-comparison.md](prod-comparison.md).

Figures below name the workbook sheet and row they come from. Read the values from your workbook:
it covers more subscriptions than any other run.

## Outline

| # | Slide title (conclusion) | Content | Source |
|---|---|---|---|
| 0 | Backup en Azure: costo, riesgos y decisiones | Title | |
| 1 | Resumen: US$[X]/mes en backup, US$[Y]/mes recuperables y [N] brechas en PROD | Three figures and the decisions requested | Resumen; Oportunidades total; Riesgos |
| **A** | **Costo actual (evidencia)** | | |
| 2 | El backup cuesta US$[X]/mes y creció [Y]% desde agosto | Stacked column by component, Aug–Oct | Screenshots 1–7; Resumen "Gasto relacionado a backup" |
| 3 | Bóvedas: US$[X]/mes, [Z]% es almacenamiento GRS | Screenshot view 1 | Gasto_medido, category Backup |
| 4 | Snapshots: US$[X]/mes manuales + US$[Y]/mes de Azure Backup | Screenshots views 2 and 3 | Gasto_medido; Por_ambiente snapshot rows |
| 5 | Bases de datos: US$[X]/mes, sobre todo Hyperscale y LTR | Screenshots views 4–7 side by side | Cost analysis by meter subcategory |
| 6 | [Z]% del costo de backup es PROD | Screenshot view 1 grouped by tag `environment` | Por_ambiente cost section |
| **B** | **Insights PROD** | | |
| 7 | [N] VMs PROD sin backup y [M] con backup fallido | Coverage bar + three counts | Por_ambiente VM rows; Riesgos R1, R2 |
| 8 | Pagamos US$[X]/mes por backups de recursos que ya no existen | Orphan vault items + snapshots of deleted disks | Por_ambiente "Ítems huérfanos", "Discos eliminados"; Oportunidades O1, O10 |
| 9 | Los snapshots manuales PROD cuestan US$[X]/mes; [Y]% son completos y antiguos | Snapshot GB and cost, PROD column | Por_ambiente; Oportunidades O12 |
| 10 | [N] VMs guardan 360 puntos diarios: pasar a 30 diarios + 12 mensuales ahorra US$[X]/mes | Today vs proposal | Por_ambiente retention row; Oportunidades O2 |
| 11 | Ningún servidor SQL PROD tiene delete lock: borrar uno pierde el PITR de todas sus bases | Servers without lock; deleted databases with LTR | Por_ambiente SQL rows; D_SQL_servidores; [sql-deleted-database-restore.md](sql-deleted-database-restore.md) |
| 12 | [N] bases PROD eliminadas conservan [M] backups LTR hasta [año] | Table: server, deletion date, expiry | D_BD_eliminadas, `backup_type = LTR` |
| 13 | Brechas que cerrar antes de recortar | Severity table | Riesgos (R1, R2, R3, R7, R9, R11, R13, R14) |
| 14 | Retener 12 años una VM no garantiza restaurarla | Long-term restore risks | Restauracion_LP |
| **C** | **Insights no productivos** | | |
| 15 | CERT y DESA casi no tienen backup de VM, pero sí SQL geo-redundante y Redis Premium | Por_ambiente CERT and DESA columns side by side | Por_ambiente |
| 16 | US$[X]/mes en no productivo no aporta protección | Redis Premium, empty and zone-redundant plans, unattached disks, SQL geo | Oportunidades O4–O8, O11; Por_ambiente |
| 17 | Decisión: CERT sigue la política de PROD o la de DESA, no una propia | Where CERT looks like PROD and where like DESA | Por_ambiente; Politica_minima |
| 18 | Lo que no conviene priorizar | Findings with almost no saving | Oportunidades O9; O5 if small |
| **D** | **Comparación (solo PROD)** | | |
| 19 | Azure optimizado cuesta US$[X]/mes por [Y] TB en PROD | Today, optimized and third-party scenarios | Comparacion_terceros (production copy) |
| 20 | El tercero solo es más barato si [N]% del volumen va a Archive | Break-even column | Comparacion_terceros column G |
| 21 | Un tercero reemplaza el backup en bóveda, no el nativo de PaaS | RPO, RTO and LTR per workload | Comparacion_terceros RPO table |
| 22 | Preguntas abiertas al proveedor | Unit (protected or stored TB), what the tariff includes, egress, rehydration | Comparacion_terceros questions |
| **E** | **Cierre** | | |
| 23 | Hoja de ruta de 90 días | 30/60/90 days: risks first, then savings | Riesgos; Oportunidades |
| 24 | Decisiones que se necesitan | Approvals with owner and date | |
| 25 | Anexo: de dónde sale cada cifra | Measured, derived, estimated, vendor data | Fuentes; Consultas; Supuestos |

## Copilot prompts

Run them in order. Check every number Copilot writes against the workbook.

### 1. Brief from the workbook (Copilot chat, workbook attached)

```text
Eres un arquitecto cloud preparando una presentación para la gerencia del banco sobre el backup en Azure.
Usa SOLO las cifras del archivo adjunto backup_insights.xlsx. Si una cifra no está, escribe [DATO PENDIENTE].
No redondees de forma distinta a la del archivo y no calcules cifras nuevas salvo que te lo pida.

Arma un resumen en cuatro secciones, con una tabla por sección (indicador, PROD, CERT, DESA, total, hoja y fila de origen):
A. Costo actual: hojas Resumen y Gasto_medido (costo de backup de octubre 2026 por componente).
B. Insights PROD: hoja Por_ambiente, columna PROD, y hojas Oportunidades y Riesgos.
C. Insights no productivos: hoja Por_ambiente, columnas CERT y DESA, y oportunidades O4 a O8 y O11.
D. Comparación: hoja Comparacion_terceros (escenarios, costo mensual, US$ por TB y punto de equilibrio).
Al final, lista los supuestos de la hoja Supuestos que afectan los ahorros.
```

Save the answer as a Word document (for example `brief_backup.docx`) after checking it.

### 2. Deck from the brief (PowerPoint, Copilot > Create a presentation from a file)

```text
Crea una presentación en español a partir de brief_backup.docx con este orden:
1) Portada y resumen ejecutivo, 2) Costo actual, 3) Insights de producción, 4) Insights de no productivo (DESA y CERT),
5) Comparación con backup de terceros solo para producción, 6) Hoja de ruta y decisiones, 7) Anexo de fuentes.
Cada título debe ser la conclusión con su cifra, no un tema (ejemplo: "Pagamos US$9k al mes por backups de VMs que ya no existen").
Una idea por diapositiva. Usa solo cifras del documento. En las notas del orador, indica la hoja y fila de origen de cada cifra.
Deja en la sección de costo actual una diapositiva por componente con un espacio para pegar una captura de Cost Management.
```

### 3. Slide by slide (PowerPoint, Copilot on the open deck)

Cost evidence:

```text
Agrega una diapositiva titulada "El backup cuesta US$[X]/mes y creció [Y]% desde agosto".
Deja el 70% del espacio para una captura de pantalla y a la derecha tres cifras: bóvedas, snapshots y bases de datos (octubre 2026).
Pie de página: "Fuente: Cost Management, costo real, agosto a octubre 2026".
```

Production insight:

```text
Agrega una diapositiva titulada "Ningún servidor SQL de producción tiene delete lock".
Explica en tres viñetas: borrar un servidor elimina el PITR de todas sus bases; el LTR sobrevive hasta su vencimiento;
un lock CanNotDelete no tiene costo. Muestra la cifra de servidores PROD sin lock y la de bases PROD eliminadas con LTR del brief.
Nota del orador: fuente hoja Por_ambiente, filas de bases de datos.
```

Non-production:

```text
Agrega una diapositiva que compare CERT con PROD y DESA en una tabla: VMs con backup, SQL con backup geo-redundante,
Redis Premium y storage sin mecanismo de recuperación. Título: "CERT se parece a DESA en VMs y a PROD en bases de datos".
Cierra con la pregunta: "¿CERT sigue la política de PROD o la de DESA?".
```

Comparison:

```text
Agrega una diapositiva con un gráfico de barras horizontales del costo mensual de cada escenario de la hoja
Comparacion_terceros (copia de producción): Azure actual, Azure optimizado, D3 y Opción 1 en Active y con Archive.
Título: "El tercero solo es más barato si [N]% del volumen de producción va a Archive".
Nota: las tarifas de D3 y Opción 1 son USD por TB-mes según el proveedor; falta confirmar si es TB protegido o almacenado.
```

### 4. Final check (PowerPoint, Copilot on the finished deck)

```text
Lista en una tabla todas las cifras de la presentación: número de diapositiva, cifra, texto donde aparece.
Marca las que no tienen fuente en las notas del orador.
```

Compare that table with the workbook. Fix by hand; do not ask Copilot to "correct" numbers.

## Before presenting

- Every cost figure matches its Cost analysis screenshot for October.
- Every estimate says "estimado" and its assumption is in the annex.
- The D3 / Opción 1 names and units are confirmed with whoever shared the tariffs.
- The orphan and LTR findings were checked with at least one application owner, so the deck does not
  propose deleting something that is a legal requirement.
