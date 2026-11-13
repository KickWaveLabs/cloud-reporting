# Backup guidelines: checklist to validate with Compliance

The guidelines deck (`src/bu_guidelines_deck.py`, slide 1) rests on our reading of each regulation. Take
these questions to the Compliance team before the deck goes out. The questions are in Spanish so they can be
pasted as they are.

Each item states what the deck assumes today and what to edit in the script if the answer differs. Unless
noted, the edit is in `slide_guidelines()`: the `scopes` list (the "Cuándo varía" table) and the speaker
notes below it. Rebuild with `python backup/src/bu_guidelines_deck.py`.

## Ley 26702, art. 183: libros y documentos (10 años)

- [ ] **¿El plazo de 10 años se cumple conservando el dato en la aplicación, sin un backup de 10 años?**
  - Deck assumes: yes. The backup stays at the base rule unless the application purges records earlier.
  - If not: set the "Libros y documentos contables" row to `+ A10` without the purge condition.
- [ ] **¿Qué aplicaciones guardan libros y documentos contables, y alguna borra registros antes de 10 años?**
  - Deck assumes: unknown; the answer decides which systems get A10.
  - If listed: no script change. Keep the list with the deck and apply A10 only to those systems.

## Res. SBS 2660-2015, art. 54: registro de operaciones

- [ ] **"Una copia de seguridad al final de cada trimestre, que se compendia en períodos de cinco años": ¿basta con
  conservar las 20 copias trimestrales de los últimos 5 años?**
  - Deck assumes: yes, 20 copies.
  - If they must be kept 10 years (40 copies): change the row to "copia trimestral por 10 años (40 copias)". With
    native SQL LTR that means M120 instead of M60.
- [ ] **¿Qué aplicación guarda el registro de operaciones? ¿Esa copia trimestral ya existe fuera de Azure
  (on-premises u otro repositorio)?**
  - Deck assumes: it has to be produced in Azure.
  - If it already exists: change the row to "Base; la copia trimestral ya existe en <lugar>".
- [ ] **¿La copia trimestral puede estar en almacenamiento Archive, que tarda hasta 15 horas en recuperarse, si la
  SBS da 48 horas hábiles para entregarla?**
  - Deck assumes: yes.
  - If not: replace "Archive inmutable" with "almacenamiento Cool inmutable" in the row.

## Res. SBS 2660-2015, arts. 55 y 58: KYC, contrapartes, operaciones inusuales

- [ ] **¿Los 10 años se cumplen conservando el dato en la aplicación, sin un backup de 10 años?**
  - Deck assumes: yes, same reading as art. 183.
  - If not: set the "KYC, contrapartes y operaciones inusuales" row to `+ A10`.
- [ ] **¿Qué aplicaciones guardan esa información (onboarding, KYC, cumplimiento) y alguna la purga antes de 10
  años?**
  - Deck assumes: unknown.
  - If listed: no script change; A10 applies only to the systems that purge.

## Circular BCRP 0011-2026, Anexo V, sección VI: LBTR

- [ ] **"Conservar los respaldos en medios seguros por un periodo no menor a diez años": ¿con qué frecuencia?
  ¿Basta un respaldo anual (10 copias), o deben ser mensuales o diarios?**
  - Deck assumes: one yearly backup kept 10 years (A10).
  - If monthly: change the row to `+ M120`. If daily, the requirement is about 3,650 copies per system; ask for
    the exact text before changing anything.
- [ ] **¿Qué aplicaciones del banco son "información relacionada con la operatividad del Sistema LBTR"?**
  - Deck assumes: unknown; the answer decides the scope of the row.
  - If listed: no script change.
- [ ] **¿Quién documenta la prueba de restauración anual y dónde queda la evidencia?**
  - Deck assumes: "prueba de restauración anual" without an owner.
  - If an owner is named: add it to the "Sistemas del LBTR" row.

## Res. SBS 504-2021: seguridad de la información

- [ ] **Art. 12, num. 3, lit. f) y h): ¿con qué frecuencia hay que probar la restauración en los sistemas que no
  son regulados?**
  - Deck assumes: the regulation sets no frequency. The yearly test is recommended only for regulated scopes.
  - If a frequency applies to all: add it to the "Cualquier otro dato" row.
- [ ] **Arts. 23 y 25: ¿la autorización vigente para procesar datos en el exterior cubre la copia geo-redundante
  en Central US?**
  - Deck assumes: yes. The workloads already run in East US 2 and Central US.
  - If not: remove "copia geo-redundante en Central US" from the recommendation in the note box at the bottom.

## Res. SBS 00814-2025: continuidad

- [ ] **¿Confirmamos que el RPO cero de los productos pasivos se cumple con replicación (DR) y queda fuera del
  lineamiento de backup?**
  - Deck assumes: yes ("Replicación y failover; no es backup").
  - If not: the requirement cannot be met with backup. Raise it as a DR topic, not in this deck.

## Not in the deck yet

- [ ] **Ley 29733 y DS 016-2024-JUS: ¿hay un plazo máximo o una obligación de supresión que choque con guardar
  backups con datos personales durante 10 años?**
  - Deck assumes: not covered; the deck cites neither norm.
  - If yes: add a row to `scopes` and the norm to `REGULATION_REFERENCES`, with its official link.
- [ ] **¿La inmutabilidad de la bóveda es obligatoria o recomendada para los ámbitos regulados?**
  - Deck assumes: recommended.
  - If mandatory: move it from the note box into the "Cómo se aplica en Azure" column of each regulated row.
- [ ] **¿Hay datos regulados reales en los ambientes no productivos?**
  - Deck assumes: no ("No productivo: PITR 7 días, sin LTR").
  - If yes: either mask the data or apply the same scope rules there.
