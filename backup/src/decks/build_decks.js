// Builds the two backup decks from metrics.json (see ../bu_deck_metrics.py).
//   npm install            (once, in this folder)
//   node build_decks.js <metrics.json> <output_dir> [apply_theme.js]

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const [, , metricsPath, outDir, applyThemePath] = process.argv;
const M = JSON.parse(fs.readFileSync(metricsPath, "utf-8"));

const THEME = {
  name: "Backup Azure",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B1F2A", lt1: "FFFFFF", dk2: "0B2A5B", lt2: "EEF2F7",
    accent1: "0B2A5B", accent2: "FF7800", accent3: "2E6FBF", accent4: "C0392B",
    accent5: "2E8B57", accent6: "8A94A6", hlink: "2E6FBF", folHlink: "6B4FA0",
  },
};
const HEX = {
  navy: "0B2A5B", orange: "FF7800", blue: "2E6FBF", red: "C0392B", green: "2E8B57", gray: "8A94A6",
  light: "EEF2F7", text: "1B1F2A", muted: "5B6475", grid: "D9DEE7", white: "FFFFFF",
};
const W = 13.333;

const fmt = (n) => Number(n || 0).toLocaleString("es-PE", { maximumFractionDigits: 0 });
const usd = (n) => "US$" + fmt(n);
const usdk = (n) => "US$" + (Number(n || 0) / 1000).toLocaleString("es-PE", { maximumFractionDigits: 1 }) + "k";
const pct = (a, b) => {
  if (!b) return "0%";
  const value = (100 * a) / b;
  // Keep one decimal near the ends so 1,733 of 1,737 does not read as 100%.
  const shown = value > 99 && value < 100 ? value.toFixed(1) : Math.round(value);
  return String(shown) + "%";
};
const find = (pairs, key) => ((pairs || []).find((p) => p[0] === key) || [key, 0])[1];
const opp = (id) => M.opportunities.find((o) => o.id === id) || { count: 0, monthly: 0, annual: 0 };
const risk = (id) => M.risks.find((r) => r.id === id) || { count: 0 };
const lt = (id) => M.long_term.find((r) => r.id === id) || { count: 0 };

function newDeck(title, footer) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.title = title;
  pres.author = "Arquitectura Cloud";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  pres.defineSlideMaster({
    title: "TITLE_DARK",
    background: { color: HEX.navy },
    objects: [],
    slideNumber: undefined,
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: HEX.white },
    margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
      { text: { text: footer, options: { x: 0.6, y: 7.05, w: 8, h: 0.3, fontSize: 10, color: HEX.gray } } },
    ],
    slideNumber: { x: 12.3, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: HEX.gray },
  });
  return { pres, C };
}

function title(slide, text, sub) {
  slide.addText(text, {
    x: 0.6, y: 0.35, w: W - 1.2, h: 0.8, fontSize: 30, bold: true, color: HEX.navy, margin: 0,
    isTextBox: true, objectName: "Title", valign: "top", fit: "shrink",
  });
  if (sub) {
    slide.addText(sub, { x: 0.6, y: 1.12, w: W - 1.2, h: 0.4, fontSize: 15, color: HEX.muted, margin: 0, isTextBox: true });
  }
}

function card(slide, x, y, w, h, opts = {}) {
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.12, fill: { color: opts.fill || HEX.light }, line: { color: opts.fill || HEX.light },
    objectName: opts.name || "Card",
  });
}

function stat(slide, x, y, w, value, label, color, size = 40) {
  slide.addText(value, { x, y, w, h: 0.8, fontSize: size, bold: true, color: color || HEX.navy, margin: 0, isTextBox: true });
  slide.addText(label, { x, y: y + 0.82, w, h: 0.7, fontSize: 13, color: HEX.muted, margin: 0, isTextBox: true, valign: "top" });
}

function badge(slide, x, y, text, color) {
  slide.addShape("ellipse", { x, y, w: 0.46, h: 0.46, fill: { color: color || HEX.orange }, line: { color: color || HEX.orange } });
  slide.addText(text, { x, y, w: 0.46, h: 0.46, fontSize: 13, bold: true, color: HEX.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
}

function bullets(slide, items, x, y, w, h, size = 15) {
  slide.addText(items.map((t, i) => ({
    text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 8 },
  })), { x, y, w, h, fontSize: size, color: HEX.text, valign: "top", margin: 0.05, isTextBox: true });
}

function chartOpts(extra) {
  return Object.assign({
    catAxisLabelColor: HEX.muted, valAxisLabelColor: HEX.muted, catAxisLabelFontFace: "+mn-lt",
    valAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
    valGridLine: { color: HEX.grid, size: 0.5 }, catGridLine: { style: "none" },
    showValue: true, dataLabelFontSize: 12, dataLabelFontFace: "+mn-lt", dataLabelColor: HEX.text,
    showLegend: false, titleFontFace: "+mn-lt", titleColor: HEX.text, titleFontSize: 14,
  }, extra);
}

function titleSlide(pres, heading, sub, note) {
  const s = pres.addSlide({ masterName: "TITLE_DARK" });
  s.addShape("ellipse", { x: 10.4, y: -1.2, w: 4.2, h: 4.2, fill: { color: HEX.orange, transparency: 15 }, line: { color: HEX.orange, transparency: 15 } });
  s.addShape("ellipse", { x: 11.6, y: 4.9, w: 2.6, h: 2.6, fill: { color: HEX.blue, transparency: 35 }, line: { color: HEX.blue, transparency: 35 } });
  s.addText(heading, { x: 0.8, y: 2.3, w: 9.6, h: 1.6, fontSize: 42, bold: true, color: HEX.white, margin: 0, isTextBox: true, valign: "bottom" });
  s.addText(sub, { x: 0.8, y: 4.05, w: 9.6, h: 0.9, fontSize: 18, color: "CADCFC", margin: 0, isTextBox: true, valign: "top" });
  if (note) s.addNotes(note);
  return s;
}

function table(slide, header, rows, opts) {
  const head = header.map((h) => ({ text: h, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, fontSize: 12 } }));
  const body = rows.map((r, i) => r.map((c) => (typeof c === "object" && c !== null && c.text !== undefined
    ? { text: String(c.text), options: Object.assign({ fontSize: 12, fill: { color: i % 2 ? HEX.white : HEX.light } }, c.options) }
    : { text: String(c), options: { fontSize: 12, color: HEX.text, fill: { color: i % 2 ? HEX.white : HEX.light } } })));
  slide.addTable([head, ...body], Object.assign({ border: { type: "solid", color: HEX.white, pt: 1 }, valign: "middle", margin: 0.06 }, opts));
}

const LT_SHORT = {
  L1: "Restaura el SO y agentes del día del backup: en 10 años estarán sin soporte ni parches.",
  L2: "La VM ya no existe: sin documentación viva de red ni dependencias.",
  L3: "Las familias de VM nuevas solo aceptan Gen2.",
  L4: "Requiere la versión de clave usada en el backup; si se borra, no restaura.",
  L5: "El motor restaura, pero la app actual puede no leer un esquema de hace 5 años.",
  L6: "Solo restaura a la misma versión mayor; cuando se retira, no hay destino.",
  L7: "Manifiestos con APIs deprecadas fallan al reconstruir.",
  L8: "Restaurar requiere ticket a soporte y la ventana es de 8 h.",
};

function sourcesSlide(pres) {
  const s = pres.addSlide({ masterName: "CONTENT" });
  title(s, "Anexo: de dónde sale cada cifra", "Todo es lectura de Azure; los costos son el costo real facturado, no una calculadora");
  const kind = (text, color) => ({ text, options: { bold: true, color } });
  table(s, ["Cifra", "Fuente", "Tipo", "Cómo verificarla"], [
    ["Inventario y configuración: VMs, discos, snapshots, bóvedas, políticas, versiones", "Azure Resource Graph (consultas KQL)", kind("Medido", HEX.green), "Resource Graph Explorer con la misma consulta, o la vista del recurso"],
    ["Retención SQL, protección de datos de storage, locks, Redis", "API de Azure Resource Manager (solo GET)", kind("Medido", HEX.green), "Vista del recurso: Backups, Data protection, Locks"],
    ["Gasto total, por medidor, por bóveda y por cuenta de storage", "Cost Management Query API, Costo real, oct-2026", kind("Costo real", HEX.green), "Cost Management > Cost analysis, agrupar por Medidor o Recurso"],
    ["Tarifas efectivas: US$7 por instancia, US$0.031/GB GRS, precio por tier de disco", "Costo / cantidad facturada del medidor", kind("Derivado", HEX.navy), "Cost analysis: columnas Costo y Cantidad del medidor"],
    ["Ahorros estimados", "Recursos x tarifa medida; prorrateos por ítem o GB", kind("Estimado", HEX.orange), "Excel: hojas Oportunidades y Supuestos"],
    ["Precio Redis, plan vacío, % de reducción y de geo sin requisito", "Supuesto editable", kind("Supuesto", HEX.red), "Pricing calculator u hoja de precios del contrato"],
    ["Tarifas D3 / Opción 1", "Proveedor", kind("Dato externo", HEX.orange), "Confirmar unidad (TB protegido o almacenado) y alcance"],
  ], { x: 0.6, y: 1.7, w: 12.13, colW: [4.2, 3.2, 1.3, 3.43], fontSize: 12, rowH: 0.56 });
  s.addText("Las consultas exactas, con fecha, filas y errores de cada corrida, están en la hoja Consultas del Excel anexo; la hoja Fuentes detalla cada indicador.",
    { x: 0.6, y: 6.45, w: 12.13, h: 0.45, fontSize: 13, color: HEX.muted, margin: 0, isTextBox: true });
  s.addNotes("Cost Management Query API devuelve el mismo costo real que Cost analysis en el portal, con los descuentos del contrato. Los únicos valores no medidos son los marcados como supuesto o dato del proveedor.");
  return s;
}

const SEV = { ALTA: HEX.red, MEDIA: HEX.orange, BAJA: HEX.gray };

// ---------------------------------------------------------------------------
// Deck 1: current situation
// ---------------------------------------------------------------------------

function deckCurrent() {
  const { pres } = newDeck("Backup en Azure: situación actual", "Backup en Azure · Situación actual · Nov 2026");
  const c = M.compute, d = M.database, n = M.native, v = M.vaults, st = M.storage, cost = M.cost;
  const prodCov = c.coverage.PROD;

  titleSlide(pres, "Backup en Azure: situación actual",
    `Inventario de los 4 frentes de protección · ${fmt(350)} suscripciones visibles · noviembre 2026`,
    "Objetivo: mostrar qué se protege hoy, con qué mecanismo (nativo, Recovery Services, Backup vault o reconstrucción) y qué no. Los insights y oportunidades van en la segunda presentación.");

  // Executive summary
  let s = pres.addSlide({ masterName: "CONTENT" });
  title(s, "Cada frente usa un modelo de protección distinto", "Qué hay y cómo se protege hoy");
  const fronts = [
    ["A", "Bases de datos", fmt(d.total), "bases y servidores", `${pct(d.native, d.total)} con backup nativo del servicio. Ninguna base usa Recovery Services ni Backup vault.`],
    ["B", "Storage", st.available ? fmt(st.total) : "En curso", st.available ? "cuentas, shares, discos y volúmenes" : "inventario en ejecución", st.available ? "Protección dominada por funciones nativas (soft delete, versioning, snapshots)." : "Se completa con la corrida de storage."],
    ["C", "Compute", fmt(c.vms_total), "VMs + " + fmt(c.k8s_total) + " clusters", `${pct(prodCov.with, prodCov.with + prodCov.without)} de las VMs PROD tradicionales con Recovery Services. Ningún cluster con backup.`],
    ["D", "Nativo y reconstrucción", fmt(n.total), "recursos PaaS y red", `Key Vault con purge protection en ${pct(n.kv_purge, n.kv_total)}. ${pct(n.iac, n.total)} con evidencia de IaC.`],
  ];
  fronts.forEach((f, i) => {
    const x = 0.6 + i * 3.08;
    card(s, x, 1.8, 2.88, 4.9);
    badge(s, x + 0.25, 2.05, f[0], HEX.orange);
    s.addText(f[1], { x: x + 0.82, y: 2.05, w: 1.95, h: 0.46, fontSize: 16, bold: true, color: HEX.navy, valign: "middle", margin: 0, isTextBox: true });
    stat(s, x + 0.25, 2.8, 2.45, f[2], f[3], HEX.navy, 34);
    s.addText(f[4], { x: x + 0.25, y: 4.45, w: 2.45, h: 2.05, fontSize: 14, color: HEX.text, valign: "top", margin: 0, isTextBox: true });
  });
  s.addNotes("Resumen de los cuatro frentes. Mensaje: la protección no es homogénea. Las VMs dependen de Recovery Services; las bases de datos, del backup nativo; red, Key Vault y mensajería dependen de funciones nativas e IaC.");

  // Method
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, "Cómo se midió y qué tan completo es el dato", "Lectura automatizada, sin cambios en el ambiente");
  const cols = [
    ["1", "Fuentes", ["Azure Resource Graph para inventario, bóvedas, políticas e ítems protegidos", "ARM en modo lectura para retención SQL, storage y locks", "Cost Management para el costo real de backup (oct-2026)"]],
    ["2", "Clasificación", ["Ambiente por tag environment o nombre de suscripción", "Plataformas excluidas del backup de VM: Databricks, ORCA, ARO y AKS", "codApp, operación y vendor desde tags"]],
    ["3", "Límites", ["Rol Reader: 350 suscripciones visibles. La laptop corporativa ve más VMs (1,444 vs " + fmt(c.vms_total) + ")", `Costo medido en ${cost.calibrated_vaults} bóvedas con ${fmt(cost.calibrated_items)} de ${fmt(v.items)} ítems`, "Backups dentro del SO (scripts SQL, Velero) no son visibles"]],
  ];
  cols.forEach((col, i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.8, 3.9, 4.2);
    badge(s, x + 0.25, 2.05, col[0], i === 2 ? HEX.gray : HEX.navy);
    s.addText(col[1], { x: x + 0.85, y: 2.05, w: 2.8, h: 0.46, fontSize: 18, bold: true, color: HEX.navy, valign: "middle", margin: 0, isTextBox: true });
    bullets(s, col[2], x + 0.25, 2.8, 3.45, 3.7, 14);
  });
  s.addNotes("Las cifras son reproducibles con los scripts bu_*_inventory.py. Todo fue lectura. La diferencia con la laptop corporativa se explica por permisos: hay suscripciones que esta cuenta no ve.");

  // Compute coverage
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, `${pct(prodCov.with, prodCov.with + prodCov.without)} de las VMs PROD tradicionales tienen backup; CERT y DESA casi no`, "VMs tradicionales por ambiente (excluye Databricks, ORCA y ARO)");
  const envs = ["PROD", "CERT", "DESA"];
  s.addChart(pres.charts.BAR, [
    { name: "Con backup", labels: envs, values: envs.map((e) => c.coverage[e].with) },
    { name: "Sin backup", labels: envs, values: envs.map((e) => c.coverage[e].without) },
  ], chartOpts({
    x: 0.6, y: 1.75, w: 7.4, h: 5.0, barDir: "bar", barGrouping: "stacked", chartColors: [HEX.navy, "C9D1DE"],
    showLegend: true, legendPos: "b", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.muted,
    dataLabelPosition: "ctr", dataLabelColor: HEX.white, catAxisOrientation: "maxMin",
  }));
  const vf = (k) => find(c.findings, k);
  [[fmt(vf("PROD_WITHOUT_BACKUP")), "VMs PROD sin backup", HEX.red],
   [fmt(vf("PROD_BACKUP_UNHEALTHY")), "VMs PROD con backup fallido o sin punto en 48 h", HEX.orange],
   [fmt(vf("DEALLOCATED_VM_WITH_BACKUP")), "VMs PROD apagadas que siguen con backup", HEX.navy]].forEach((k, i) => {
    card(s, 8.4, 1.8 + i * 1.68, 4.33, 1.5);
    stat(s, 8.65, 1.92 + i * 1.68, 3.9, k[0], k[1], k[2], 34);
  });
  s.addNotes("La cobertura en PROD es razonable pero no completa. En no productivo casi no hay backup, lo cual es coherente si esas VMs se reconstruyen por IaC; hay que confirmarlo.");

  // Platform classification
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, `${pct(c.platform_excluded, c.vms_total)} de las VMs pertenecen a plataformas que no requieren backup de VM`, "VMs por clase de infraestructura");
  const classLabel = { DATABRICKS_MANAGED: "Databricks", TRADITIONAL: "Tradicional", ORCA_SECURITY: "ORCA Security", ARO_OPENSHIFT: "ARO (OpenShift)", AKS_MANAGED: "AKS", MANAGED_RG_OTHER: "Otro gestionado" };
  s.addChart(pres.charts.BAR, [{ name: "VMs", labels: c.classes.map((x) => classLabel[x[0]] || x[0]), values: c.classes.map((x) => x[1]) }],
    chartOpts({ x: 0.6, y: 1.75, w: 6.6, h: 5.0, barDir: "bar", chartColors: [HEX.navy], dataLabelPosition: "outEnd", catAxisOrientation: "maxMin" }));
  card(s, 7.6, 1.8, 5.13, 4.9);
  s.addText("Kubernetes", { x: 7.9, y: 2.0, w: 4.6, h: 0.45, fontSize: 18, bold: true, color: HEX.navy, margin: 0, isTextBox: true });
  stat(s, 7.9, 2.55, 2.2, fmt(c.k8s_total), `clusters (${fmt(c.aks)} AKS, ${fmt(c.aro)} ARO)`, HEX.navy, 34);
  stat(s, 10.3, 2.55, 2.2, fmt(c.k8s_backup), "con backup", HEX.red, 34);
  bullets(s, [
    `${fmt(c.flux)} clusters usan GitOps (Flux): se pueden reconstruir desde Git`,
    "ARO ahora se clasifica aparte: antes figuraba como VM tradicional",
    "No hay extensión de AKS Backup instalada en ningún cluster",
  ], 7.9, 4.2, 4.6, 2.3, 14);
  s.addNotes("Las VMs de Databricks, ORCA, ARO y AKS son efímeras o las gestiona la plataforma. Sacarlas del denominador evita reportar falsos 'sin backup'.");

  // Vaults
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, `${fmt(v.total)} bóvedas, ${fmt(v.with_items)} en uso: todas GRS y ninguna inmutable`, "Recovery Services y Backup vault");
  [[fmt(v.total), "bóvedas"], [fmt(v.with_items), "con ítems protegidos"], [fmt(v.grs), "con GRS"], [fmt(v.immutable), "inmutables"], [fmt(v.crr), "con restauración cruzada de región"]]
    .forEach((k, i) => {
      card(s, 0.6 + i * 2.48, 1.8, 2.3, 1.75);
      stat(s, 0.8 + i * 2.48, 1.95, 2.0, k[0], k[1], i >= 3 ? HEX.red : HEX.navy, 34);
    });
  s.addChart(pres.charts.BAR, [{ name: "VMs", labels: c.retention.map((x) => x[0]), values: c.retention.map((x) => x[1]) }],
    chartOpts({ x: 0.6, y: 3.85, w: 6.6, h: 2.95, barDir: "col", chartColors: [HEX.orange], dataLabelPosition: "outEnd", showTitle: true, title: `Retención máxima de los ${fmt(c.active_backups)} backups de VM activos` }));
  bullets(s, [
    `${fmt(v.policies)} políticas: ${fmt(v.policies_unused)} sin ítems (en su mayoría por defecto)`,
    `${fmt(c.policy_v1)} de ${fmt(c.active_backups)} VMs usan la política Standard (V1) diaria`,
    `${fmt(c.archive)} usan el tier Archive para puntos de largo plazo`,
  ], 7.6, 4.0, 5.1, 2.8, 14);
  s.addNotes("La configuración de bóvedas es homogénea: GRS por defecto, sin inmutabilidad ni Cross Region Restore. Con GRS sin CRR se paga la copia secundaria pero solo se puede usar si Microsoft declara un desastre regional.");

  // Database
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, "Las bases de datos dependen del backup nativo; ninguna usa bóveda", "Bases y servidores por servicio");
  const svcLabel = { AZURE_SQL_DB: "Azure SQL DB", COSMOS_DB: "Cosmos DB", POSTGRES_FLEX: "PostgreSQL Flex", MYSQL_FLEX: "MySQL Flex", SQL_SERVER_VM: "SQL Server en VM", SQL_MI_DB: "SQL MI" };
  s.addChart(pres.charts.BAR, [{ name: "Bases", labels: d.services.map((x) => svcLabel[x[0]] || x[0]), values: d.services.map((x) => x[1]) }],
    chartOpts({ x: 0.6, y: 1.75, w: 6.2, h: 5.0, barDir: "bar", chartColors: [HEX.navy], dataLabelPosition: "outEnd", catAxisOrientation: "maxMin" }));
  const dbFacts = [
    [`${fmt(d.sql_prod_ltr)} / ${fmt(d.sql_prod)}`, "bases SQL PROD con retención larga (LTR de 1 a 5 años)"],
    ["7 días", "retención PITR en casi todas las bases SQL"],
    [fmt(d.cosmos_periodic), "cuentas Cosmos en modo periódico (ventana de 8 h)"],
    [fmt(d.prod_not_geo), "bases PROD sin copia geo-redundante (incluye todo PostgreSQL y MySQL)"],
  ];
  dbFacts.forEach((f, i) => {
    card(s, 7.2, 1.8 + i * 1.24, 5.53, 1.1);
    s.addText(f[0], { x: 7.45, y: 1.85 + i * 1.24, w: 1.9, h: 1.0, fontSize: 24, bold: true, color: i === 3 ? HEX.red : HEX.navy, valign: "middle", margin: 0, isTextBox: true });
    s.addText(f[1], { x: 9.4, y: 1.85 + i * 1.24, w: 3.2, h: 1.0, fontSize: 13, color: HEX.text, valign: "middle", margin: 0, isTextBox: true });
  });
  s.addNotes(`Todas las bases PaaS usan su backup nativo. Las 374 políticas SQL de Recovery Services no protegen nada. Hay ${fmt(d.sql_vm)} VMs con SQL Server cuyo backup de base de datos no es visible desde Azure.`);

  // Storage
  s = pres.addSlide({ masterName: "CONTENT" });
  if (st.available) {
    const methodLabel = { NATIVE_SOFT_DELETE_ONLY: "Solo soft delete", NONE: "Ninguno", VM_BACKUP: "Backup de la VM", SNAPSHOT_ONLY: "Solo snapshots", VAULT_RSV: "Recovery Services", VAULT_BV: "Backup vault", NATIVE_PITR: "PITR nativo", UNKNOWN: "Sin dato" };
    title(s, "Storage se protege con soft delete o con nada; la bóveda es la excepción", `${fmt(st.total)} unidades: cuentas blob/ADLS, file shares y discos (incluye discos de plataformas gestionadas)`);
    s.addChart(pres.charts.BAR, [{ name: "Unidades", labels: st.methods.map((x) => methodLabel[x[0]] || x[0]), values: st.methods.map((x) => x[1]) }],
      chartOpts({ x: 0.6, y: 1.75, w: 6.6, h: 5.0, barDir: "bar", chartColors: [HEX.navy], dataLabelPosition: "outEnd", catAxisOrientation: "maxMin", showTitle: true, title: "Mecanismo de protección" }));
    [[fmt(st.prod_without_recovery), "cuentas y discos PROD sin ningún mecanismo de recuperación", HEX.red],
     [fmt(st.prod_share_without_backup), "file shares PROD sin Azure Backup", HEX.orange],
     [fmt(st.vault_protected), "unidades protegidas en bóveda", HEX.navy],
     [fmt(st.vm_backup_disks), "discos cubiertos por el backup de su VM", HEX.navy]].forEach((k, i) => {
      const x = 7.6 + (i % 2) * 2.62, y = 1.8 + Math.floor(i / 2) * 2.5;
      card(s, x, y, 2.5, 2.3);
      stat(s, x + 0.2, y + 0.2, 2.15, k[0], k[1], k[2], 30);
    });
  } else {
    title(s, "Storage: inventario en ejecución", "5,354 cuentas de storage, shares, discos y Elastic SAN");
    bullets(s, ["Resource Graph ya muestra 1,856 cuentas ZRS y 197 RA-GRS con resiliencia declarada", "Backup vault protege 1 cuenta blob y Recovery Services 6 file shares", "El detalle de soft delete, versioning y snapshots llega con la corrida completa"], 0.6, 1.9, 12, 4, 16);
  }
  s.addNotes("En storage la primera línea de protección son las funciones nativas: soft delete, versioning, snapshots y redundancia. El backup en bóveda es la excepción.");

  // Native
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, "La protección nativa es sólida en Key Vault y débil en red, DNS y Redis", "Servicios sin Azure Backup: qué los protege hoy");
  const nat = [
    [`${pct(n.kv_purge, n.kv_total)}`, `de ${fmt(n.kv_total)} Key Vaults con purge protection y 90 días de soft delete`, HEX.green],
    [fmt(n.dns_no_lock), "zonas DNS PROD sin delete lock (83 privadas)", HEX.red],
    [fmt(n.netcore_no_lock), "recursos de red core PROD sin delete lock", HEX.orange],
    [`0 / ${fmt(n.redis_total)}`, `Redis con persistencia o geo-réplica (${fmt(n.redis_premium)} son Premium)`, HEX.red],
    [pct(n.iac, n.total), "de los recursos con evidencia de Terraform: base para reconstruir", HEX.navy],
    [fmt(n.plans_empty), `App Service Plans vacíos de ${fmt(n.plans_total)}`, HEX.navy],
  ];
  nat.forEach((f, i) => {
    const x = 0.6 + (i % 3) * 4.1, y = 1.8 + Math.floor(i / 3) * 2.5;
    card(s, x, y, 3.9, 2.3);
    stat(s, x + 0.25, y + 0.2, 3.45, f[0], f[1], f[2], 36);
  });
  s.addNotes("App Service, Functions, Key Vault, APIM, red y mensajería no tienen backup en bóveda. Su protección es nativa (soft delete, réplica, zonas) o depende de poder reconstruirlos por IaC. Los locks casi no existen en red y DNS.");

  // Cost
  s = pres.addSlide({ masterName: "CONTENT" });
  const backupSpend = M.estate["Gasto relacionado a backup oct-2026 (USD)"] || 0;
  title(s, `El backup cuesta ${usdk(backupSpend)} al mes; en bóvedas, el ${pct(cost.calibrated_storage, cost.calibrated_total)} es almacenamiento GRS`, `Costo real oct-2026: bóvedas (${cost.calibrated_vaults}), snapshots de disco (${usdk(M.snapshots.spend)}) y backup de bases de datos`);
  s.addChart(pres.charts.DOUGHNUT, [{ name: "Costo", labels: ["Almacenamiento GRS", "Instancias protegidas", "Otros"], values: [cost.calibrated_storage, cost.calibrated_pi, Math.max(0, cost.calibrated_total - cost.calibrated_storage - cost.calibrated_pi)] }],
    { x: 0.6, y: 1.7, w: 5.6, h: 5.1, holeSize: 60, chartColors: [HEX.navy, HEX.orange, HEX.gray], showPercent: true, showValue: false, dataLabelColor: HEX.white, dataLabelFontSize: 12, showLegend: true, legendPos: "b", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.muted });
  [[usd(cost.calibrated_total), "por mes en las bóvedas medidas"], [`US$${cost.gb_rate}`, "por GB-mes GRS (tarifa efectiva)"], [`US$${cost.pi_rate}`, "por instancia protegida al mes"], [usd(cost.live_item_storage), "de almacenamiento por VM viva al mes"]]
    .forEach((k, i) => {
      const x = 6.7 + (i % 2) * 3.05, y = 1.8 + Math.floor(i / 2) * 2.5;
      card(s, x, y, 2.9, 2.3);
      stat(s, x + 0.2, y + 0.2, 2.55, k[0], k[1], HEX.navy, 28);
    });
  s.addNotes("Se midió con Cost Management en las suscripciones con más ítems. Las tarifas efectivas ya incluyen el descuento corporativo y se usan en el Excel para estimar ahorros. Para el estate completo se corre bu_backup_cost.py.");

  sourcesSlide(pres);

  // Closing
  s = pres.addSlide({ masterName: "TITLE_DARK" });
  s.addText("Situación en tres líneas", { x: 0.8, y: 0.7, w: 11, h: 0.8, fontSize: 34, bold: true, color: HEX.white, margin: 0, isTextBox: true });
  [
    `Compute: cobertura PROD de ${pct(prodCov.with, prodCov.with + prodCov.without)}, pero ${fmt(M.orphans.total)} ítems de backup protegen VMs que ya no existen.`,
    `Bases de datos y storage: protección nativa del servicio; las bóvedas casi no se usan fuera de VMs (${fmt(st.prod_without_recovery || 0)} unidades de storage PROD sin recuperación).`,
    "Nativo: Key Vault bien protegido; red, DNS y Redis dependen de IaC y no tienen locks.",
  ].forEach((t, i) => {
    badge(s, 0.8, 2.0 + i * 1.45, String(i + 1), HEX.orange);
    s.addText(t, { x: 1.5, y: 1.9 + i * 1.45, w: 10.5, h: 0.75, fontSize: 19, color: HEX.white, valign: "middle", margin: 0, isTextBox: true });
  });
  s.addNotes("Cierre: con este punto de partida, la segunda presentación muestra oportunidades de ahorro, brechas y la política mínima propuesta.");

  return pres;
}

// ---------------------------------------------------------------------------
// Deck 2: insights and opportunities
// ---------------------------------------------------------------------------

function deckInsights() {
  const { pres } = newDeck("Backup en Azure: insights y oportunidades", "Backup en Azure · Insights y oportunidades · Nov 2026");
  const c = M.compute, o = M.orphans, d = M.database, n = M.native, v = M.vaults, sv = M.savings, cost = M.cost;

  titleSlide(pres, "Backup en Azure: insights y oportunidades",
    "Ahorro, brechas de protección y riesgos de restauración a largo plazo · noviembre 2026",
    "Esta presentación propone decisiones. Las cifras salen del inventario y del costo real de octubre 2026; los ahorros son estimaciones con supuestos editables en el Excel anexo.");

  // Key message
  let s = pres.addSlide({ masterName: "CONTENT" });
  title(s, "Pagamos por respaldar lo que ya no existe y dejamos huecos en PROD", "Mensaje principal");
  card(s, 0.6, 1.8, 6.0, 4.9, { fill: HEX.navy });
  s.addText(usdk(sv.backup_monthly * 12) + " / año", { x: 0.9, y: 2.2, w: 5.5, h: 1.1, fontSize: 48, bold: true, color: HEX.white, margin: 0, isTextBox: true });
  s.addText("de ahorro estimado en backup: ítems huérfanos, snapshots, retención y VMs apagadas o no productivas", { x: 0.9, y: 3.35, w: 5.4, h: 1.0, fontSize: 16, color: "CADCFC", margin: 0, isTextBox: true, valign: "top" });
  s.addText(`+ ${usdk(sv.adjacent_monthly * 12)} / año en recursos adyacentes: discos sin adjuntar, geo-replicación de storage, Redis Premium y planes vacíos`, { x: 0.9, y: 4.65, w: 5.4, h: 1.4, fontSize: 14, color: "CADCFC", margin: 0, isTextBox: true, valign: "top" });
  [[fmt(o.total), "ítems de backup huérfanos en " + fmt(o.vaults) + " bóvedas", HEX.orange],
   [fmt(risk("R1").count), "VMs PROD sin backup y " + fmt(risk("R2").count) + " con backup fallido", HEX.red],
   [fmt(v.immutable), "bóvedas inmutables de " + fmt(v.total) + ": sin protección ante borrado malicioso", HEX.red]].forEach((k, i) => {
    card(s, 7.0, 1.8 + i * 1.68, 5.73, 1.5);
    stat(s, 7.25, 1.92 + i * 1.68, 5.3, k[0], k[1], k[2], 32);
  });
  s.addNotes("El hallazgo central: el costo y la protección no están alineados. La mayor parte del almacenamiento medido pertenece a ítems huérfanos, mientras hay VMs PROD sin backup y ninguna bóveda inmutable.");

  // Opportunity map
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, "Mapa de oportunidades ordenado por ahorro", "Ahorro mensual estimado con tarifas efectivas del contrato; supuestos editables en el Excel");
  const sorted = [...M.opportunities].filter((x) => x.monthly > 0).sort((a, b) => b.monthly - a.monthly);
  table(s, ["ID", "Oportunidad", "Recursos", "Ahorro / mes", "Ahorro / año", "Confianza"],
    sorted.map((x) => [x.id, x.title, fmt(x.count), usd(x.monthly), usd(x.annual), String(x.confidence).split(":")[0].split("(")[0].trim()]),
    { x: 0.6, y: 1.7, w: 12.13, colW: [0.6, 6.0, 1.1, 1.4, 1.5, 1.53], fontSize: 12, rowH: 0.42 });
  s.addText(`Total: ${usd(sv.total_monthly)} / mes · ${usd(sv.total_monthly * 12)} / año. Sin ahorro relevante: SQL geo en no productivo, bóvedas vacías y políticas sin uso.`,
    { x: 0.6, y: 6.45, w: 12.13, h: 0.4, fontSize: 13, color: HEX.muted, margin: 0, isTextBox: true });
  s.addNotes("La confianza indica si el ahorro sale de costo medido (alta), de tarifa efectiva por recurso (media) o de un precio supuesto a validar (baja).");

  // Orphans
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, `${fmt(o.total)} ítems de backup protegen recursos que ya no existen`, "Antigüedad del último punto de recuperación de los ítems huérfanos");
  const ageLabels = o.age_years.map((x) => (x[0] >= 5 ? "5+ años" : x[0] === 0 ? "< 1 año" : `${x[0]} año${x[0] > 1 ? "s" : ""}`));
  s.addChart(pres.charts.BAR, [{ name: "Ítems", labels: ageLabels, values: o.age_years.map((x) => x[1]) }],
    chartOpts({ x: 0.6, y: 1.75, w: 6.8, h: 5.0, barDir: "col", chartColors: [HEX.orange], dataLabelPosition: "outEnd" }));
  card(s, 7.8, 1.8, 4.93, 2.15, { fill: HEX.navy });
  s.addText(`${usdk(o.measured_floor)} – ${usdk(opp("O1").monthly)} / mes`, { x: 8.05, y: 1.95, w: 4.5, h: 0.8, fontSize: 28, bold: true, color: HEX.white, margin: 0, isTextBox: true });
  s.addText("Piso medido (bóvedas 100% huérfanas + instancias) y estimado con prorrateo", { x: 8.05, y: 2.8, w: 4.5, h: 0.9, fontSize: 13, color: "CADCFC", margin: 0, isTextBox: true, valign: "top" });
  bullets(s, [
    `${fmt(o.deleted)} con la VM eliminada y ${fmt(o.stale)} vivas sin backup hace más de 30 días`,
    `Siguen en estado 'Protected': pagan instancia y almacenamiento GRS`,
    "Acción: validar con el dueño, detener protección y borrar datos salvo requisito legal",
  ], 7.8, 4.2, 4.93, 2.5, 14);
  s.addNotes("Cuando una VM se elimina sin detener el backup, Azure conserva el último punto indefinidamente y sigue cobrando. Es la palanca de ahorro más clara y de menor riesgo.");

  // Snapshots
  const sn = M.snapshots;
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, `Los snapshots de disco cuestan ${usdk(sn.spend)} al mes y el ${pct(sn.full, sn.count)} son completos`, "Snapshots manuales fuera de Azure Backup (oct-2026)");
  s.addChart(pres.charts.DOUGHNUT, [{ name: "GB", labels: ["Disco origen eliminado", "Disco vivo"], values: [sn.orphan_gb, Math.max(0, sn.gb - sn.orphan_gb)] }],
    { x: 0.6, y: 1.7, w: 5.2, h: 5.0, holeSize: 60, chartColors: [HEX.orange, HEX.navy], showPercent: true, showValue: false, dataLabelColor: HEX.white, dataLabelFontSize: 12, showLegend: true, legendPos: "b", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.muted, showTitle: true, title: "GB de snapshots por estado del disco origen", titleFontFace: "+mn-lt", titleFontSize: 14, titleColor: HEX.text });
  [[fmt(sn.count), `snapshots (${fmt(sn.incremental)} incrementales)`, HEX.navy],
   [fmt(sn.orphan_count), "snapshots de discos que ya no existen", HEX.orange],
   [fmt(sn.old_sources), "discos con snapshots de más de 1 año", HEX.orange],
   [usd(opp("O10").monthly + opp("O12").monthly), "ahorro estimado al mes (huérfanos + expiración)", HEX.green]].forEach((k, i) => {
    const x = 6.3 + (i % 2) * 3.23, y = 1.8 + Math.floor(i / 2) * 2.5;
    card(s, x, y, 3.1, 2.3);
    stat(s, x + 0.22, y + 0.2, 2.7, k[0], k[1], k[2], 30);
  });
  s.addNotes("Los snapshots completos se cobran por el tamaño usado completo en cada copia; los incrementales solo por el delta. La mayoría de los snapshots actuales son completos y antiguos, y una parte pertenece a discos eliminados. Es la segunda palanca de ahorro después de los backups huérfanos.");

  // Storage spend
  const sc = M.storage_cost || [];
  if (sc.length) {
    s = pres.addSlide({ masterName: "CONTENT" });
    const geo = sc.reduce((acc, a) => acc + (a.parts["Transferencia y geo-replicación"] || 0), 0);
    title(s, "En las cuentas más caras se paga por lecturas y replicación, no por el dato guardado", "Costo oct-2026 de las cuentas de storage más costosas (muestra de 6 suscripciones), por componente");
    const groups = ["Dato almacenado", "Operaciones y transacciones", "Transferencia y geo-replicación"];
    s.addChart(pres.charts.BAR, groups.map((g) => ({ name: g, labels: sc.map((a) => a.name), values: sc.map((a) => a.parts[g] || 0) })),
      chartOpts({ x: 0.6, y: 1.75, w: 7.6, h: 5.0, barDir: "bar", barGrouping: "stacked", chartColors: [HEX.navy, HEX.blue, HEX.orange],
        showValue: false, showLegend: true, legendPos: "b", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.muted, catAxisOrientation: "maxMin", catAxisLabelFontSize: 11 }));
    card(s, 8.6, 1.8, 4.13, 2.1, { fill: HEX.navy });
    s.addText(usdk(sc.reduce((acc, a) => acc + a.total, 0)) + " / mes", { x: 8.85, y: 1.95, w: 3.7, h: 0.8, fontSize: 30, bold: true, color: HEX.white, margin: 0, isTextBox: true });
    s.addText(`en estas ${sc.length} cuentas; ${usdk(geo)} es transferencia y geo-replicación`, { x: 8.85, y: 2.8, w: 3.7, h: 0.9, fontSize: 13, color: "CADCFC", margin: 0, isTextBox: true, valign: "top" });
    bullets(s, [
      "Mover a Cool/Cold encarecería las lecturas: no es la palanca aquí",
      "Validar si RA-GRS es requisito de DR; ZRS elimina la transferencia de geo-replicación",
      "Confirmar con el dueño si las cuentas 'back*' son destino de backup o backend de una aplicación",
    ], 8.6, 4.15, 4.13, 2.6, 13);
    s.addNotes("El nombre de las cuentas sugiere backup, pero no está confirmado. El costo medido muestra que la mayor parte es operación y replicación, por lo que el ahorro depende de la decisión de DR, no del tier de almacenamiento.");
  }

  // Retention
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, `${fmt(c.retention_360)} VMs guardan 360 puntos diarios: más de lo que un restore real usa`, "Retención actual vs esquema propuesto");
  card(s, 0.6, 1.8, 5.9, 3.9);
  s.addText("Hoy", { x: 0.9, y: 2.0, w: 5.3, h: 0.5, fontSize: 20, bold: true, color: HEX.navy, margin: 0, isTextBox: true });
  bullets(s, ["Un punto diario retenido 360 días", "Restaurar a 'hace 7 meses, un martes' casi nunca se necesita", `${fmt(c.archive)} VMs usan Archive para puntos antiguos`, "Bóveda GRS sin Cross Region Restore"], 0.9, 2.7, 5.3, 3.8, 15);
  card(s, 6.83, 1.8, 5.9, 3.9, { fill: HEX.navy });
  s.addText("Propuesta (GFS)", { x: 7.13, y: 2.0, w: 5.3, h: 0.5, fontSize: 20, bold: true, color: HEX.white, margin: 0, isTextBox: true });
  s.addText([
    { text: "30 puntos diarios", options: { bullet: true, breakLine: true, paraSpaceAfter: 8 } },
    { text: "12 puntos mensuales", options: { bullet: true, breakLine: true, paraSpaceAfter: 8 } },
    { text: "Anuales solo si hay requisito regulatorio, en Archive", options: { bullet: true, breakLine: true, paraSpaceAfter: 8 } },
    { text: `Ahorro estimado ${usd(opp("O2").monthly)} / mes (supuesto 30% menos almacenamiento)`, options: { bullet: true, paraSpaceAfter: 8 } },
  ], { x: 7.13, y: 2.7, w: 5.3, h: 3.8, fontSize: 15, color: HEX.white, valign: "top", margin: 0.05, isTextBox: true });
  s.addNotes("El cambio de política no borra puntos existentes; los nuevos puntos siguen el esquema GFS y los antiguos expiran con la política anterior. Validar RPO y requisito regulatorio por aplicación.");

  // Not worth prioritizing
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, "Lo que no conviene priorizar por ahorro", "Hallazgos reales con impacto económico casi nulo");
  [[fmt(d.sql_nonprod_geo), "bases SQL no productivas con backup geo-redundante", "Costo medido de backup SQL: US$3/mes en la suscripción con más bases. El PITR está dentro del tramo gratuito."],
   [fmt(v.total - v.with_items), "bóvedas vacías", "No facturan. Limpiarlas mejora gobierno y auditoría, no el costo."],
   [fmt(v.policies_unused), "políticas sin ítems", "Políticas por defecto creadas con cada bóveda. Sin costo."]].forEach((k, i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.8, 3.9, 3.9);
    stat(s, x + 0.25, 2.0, 3.45, k[0], k[1], HEX.gray, 40);
    s.addText(k[2], { x: x + 0.25, y: 3.75, w: 3.45, h: 2.7, fontSize: 15, color: HEX.text, valign: "top", margin: 0, isTextBox: true });
  });
  s.addNotes("Es importante decir qué no mueve la aguja para que el equipo no invierta esfuerzo en cambios de bajo retorno, como pasar SQL de geo a local en no productivo.");

  // Adjacent
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, "Fuera del backup: SKUs premium y planes sin uso en no productivo", "Oportunidades adyacentes detectadas en el inventario nativo (precios a validar)");
  [[fmt(n.redis_premium_nonprod), "Redis Premium en CERT/DESA", usd(opp("O6").monthly) + " / mes", "Pasar a Standard donde no se requiera red privada inyectada ni clustering."],
   [fmt(n.plans_empty), "App Service Plans vacíos", usd(opp("O7").monthly) + " / mes", "Eliminar o consolidar."],
   [fmt(n.plans_zr_nonprod), "planes no productivos con redundancia de zona", "Por medir", "Zone redundancy exige mínimo 3 instancias."]].forEach((k, i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.8, 3.9, 3.9);
    stat(s, x + 0.25, 2.0, 3.45, k[0], k[1], HEX.navy, 40);
    s.addText(k[2], { x: x + 0.25, y: 3.6, w: 3.45, h: 0.6, fontSize: 22, bold: true, color: HEX.orange, margin: 0, isTextBox: true });
    s.addText(k[3], { x: x + 0.25, y: 4.3, w: 3.45, h: 2.2, fontSize: 15, color: HEX.text, valign: "top", margin: 0, isTextBox: true });
  });
  s.addNotes("Estas oportunidades no son de backup pero salieron del mismo inventario. El precio de Redis es referencial: reemplazar por la tarifa del contrato antes de presentarlo como compromiso.");

  // Risks
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, "Brechas que cerrar antes de recortar", "Riesgos de protección priorizados");
  const topRisks = M.risks.filter((r) => ["R1", "R2", "R3", "R9", "R11", "R13", "R4", "R7", "R14"].includes(r.id))
    .sort((a, b) => (a.severity === b.severity ? 0 : a.severity === "ALTA" ? -1 : b.severity === "ALTA" ? 1 : 0));
  table(s, ["Severidad", "Riesgo", "Recursos", "Acción"],
    topRisks.map((r) => [{ text: r.severity, options: { bold: true, color: HEX.white, fill: { color: SEV[r.severity] || HEX.gray } } }, r.title, fmt(r.count), r.action]),
    { x: 0.6, y: 1.7, w: 12.13, colW: [1.2, 4.2, 1.1, 5.63], fontSize: 11, rowH: 0.5 });
  s.addNotes("Recortar retención o bóvedas sin cerrar estas brechas traslada el riesgo. El orden sugerido: primero PROD sin backup, locks en DNS e inmutabilidad; luego los ahorros.");

  // Third-party backup comparison
  const tp = M.third_party || {};
  if ((tp.scenarios || []).length) {
    const val = tp.values || {};
    const opt = tp.scenarios.find((x) => x.name === "Azure Backup optimizado") || { monthly: 0, per_tb: 0 };
    const breakevens = tp.scenarios.filter((x) => x.breakeven !== null && x.breakeven !== undefined).map((x) => Math.round(x.breakeven * 100));
    s = pres.addSlide({ masterName: "CONTENT" });
    title(s, "El backup de terceros solo es más barato si gran parte del volumen va a Archive",
      `Costo mensual sobre el alcance comparable: ${fmt(val.base_tb)} TB en bóvedas tras optimizar (tarifas D3 y Opción 1 por TB-mes)`);
    s.addChart(pres.charts.BAR, [{ name: "USD / mes", labels: tp.scenarios.map((x) => x.name), values: tp.scenarios.map((x) => x.monthly) }],
      chartOpts({ x: 0.6, y: 1.75, w: 7.4, h: 5.0, barDir: "bar", chartColors: [HEX.navy], dataLabelPosition: "outEnd",
        dataLabelFormatCode: '"$"#,##0', catAxisOrientation: "maxMin", catAxisLabelFontSize: 11, valAxisLabelFormatCode: '"$"#,##0' }));
    [[`US$${opt.per_tb}`, "por TB-mes en Azure Backup hoy (almacenamiento GRS + instancias)", HEX.navy],
     ["US$49.90 – 56.01", "por TB-mes en Active de los terceros", HEX.orange],
     [`${Math.min(...breakevens)}–${Math.max(...breakevens)}%`, "del volumen tendría que ir a Archive para igualar a Azure optimizado", HEX.navy]].forEach((k, i) => {
      card(s, 8.4, 1.8 + i * 1.68, 4.33, 1.5);
      stat(s, 8.65, 1.92 + i * 1.68, 3.9, k[0], k[1], k[2], 26);
    });
    s.addNotes("Supuesto de unidad: las tarifas son USD por TB-mes; confirmar con el proveedor si es TB protegido o almacenado y qué incluyen (licencia, repositorio, egress, restauración). Azure Backup también tiene tier Archive para puntos mensuales y anuales, que hoy no se usa: la comparación justa es contra Azure optimizado con Archive. Limpiar huérfanos y retención ahorra más que cambiar de herramienta.");

    s = pres.addSlide({ masterName: "CONTENT" });
    title(s, "Un tercero puede reemplazar el backup en bóveda, no el backup nativo de PaaS", "RPO, RTO y retención larga por workload");
    table(s, ["Workload", "RPO hoy", "RTO hoy", "LTR hoy", "¿Lo reemplaza?"],
      tp.rpo.map((r) => [r.workload, r.rpo, r.rto, r.ltr, { text: r.replaces, options: { bold: true, color: /^No/.test(r.replaces) ? HEX.red : /^Sí/.test(r.replaces) ? HEX.green : HEX.orange } }]),
      { x: 0.6, y: 1.65, w: 12.13, colW: [2.3, 3.0, 2.8, 2.2, 1.83], fontSize: 11, rowH: 0.44 });
    const rs = (tp.restore || []).map((r) => `${fmt(r.tb)} TB ≈ ${r.hours} h`).join(" · ");
    s.addText(`Restaurar desde un repositorio fuera de Azure depende del enlace (10 Gbps al 70%): ${rs}. Azure restaura en minutos desde el snapshot tier.`,
      { x: 0.6, y: 6.55, w: 12.13, h: 0.4, fontSize: 12, color: HEX.muted, margin: 0, isTextBox: true });
    s.addNotes("PITR de SQL (log cada 5-10 min), Cosmos continuo y el soft delete de storage son parte del servicio y gratuitos o casi: un tercero no los mejora. Donde sí aporta: VMs, SQL/SAP en VM (hoy sin workload backup), Files, una copia aislada fuera del tenant y retención larga.");
  }

  // Long-term restore
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, "Retener 12 años una VM no garantiza poder restaurarla", "Riesgo de incompatibilidad de versiones y dependencias en restauraciones de largo plazo");
  table(s, ["Escenario", "Recursos", "Por qué puede fallar"],
    M.long_term.map((r) => [r.title, fmt(r.count), LT_SHORT[r.id] || r.why]),
    { x: 0.6, y: 1.7, w: 8.3, colW: [2.9, 0.9, 4.5], fontSize: 11, rowH: 0.5 });
  card(s, 9.2, 1.8, 3.53, 4.9, { fill: HEX.navy });
  s.addText("Principio", { x: 9.45, y: 2.0, w: 3.1, h: 0.5, fontSize: 20, bold: true, color: HEX.orange, margin: 0, isTextBox: true });
  s.addText([
    { text: "Retener el dato, no la imagen.", options: { bold: true, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Largo plazo: LTR de base de datos o exportes en formato abierto en blob inmutable (Archive).", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "VM: máximo 30 días + 12 mensuales. Probar una restauración al año.", options: {} },
  ], { x: 9.45, y: 2.6, w: 3.1, h: 3.9, fontSize: 14, color: HEX.white, valign: "top", margin: 0, isTextBox: true });
  s.addNotes("Una restauración de VM devuelve el sistema operativo, agentes y librerías del día del backup. En 10 años ese SO estará fuera de soporte y quizás no haya tamaño de VM compatible (Gen1). Las bases PaaS restauran sobre el motor actual, pero la aplicación puede no leer un esquema viejo. Las claves de cifrado rotadas o borradas impiden restaurar.");

  // Minimum policy
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, "Política mínima propuesta por servicio y ambiente", "Punto de partida para validar con Riesgos y Cumplimiento");
  const policy = [
    ["VM con estado", "Diario 30 d + 12 mensuales; GRS + CRR; inmutable", "Diario 30 d; ZRS/LRS", "7 d o sin backup (IaC)", "Sin backup (IaC)"],
    ["SQL Database", "PITR 14 d; LTR solo si regulatorio", "PITR 7 d; Zone", "PITR 7 d; Local", "PITR 7 d; Local"],
    ["PostgreSQL / MySQL", "35 d + geo-backup", "14 d", "7 d", "7 d"],
    ["Cosmos DB", "Continuo 7 d", "Continuo 7 d", "Periódico local", "Periódico local"],
    ["SQL Server en VM", "Workload SQL (full + log)", "Workload SQL diario", "Snapshot 7 d", "Sin backup"],
    ["Blob / Files", "Soft delete 14 d + versioning / backup 30 d", "Soft delete 14 d", "Soft delete 7 d", "Soft delete 7 d"],
    ["AKS", "GitOps + backup de PV críticos", "GitOps + PV críticos", "GitOps", "GitOps"],
    ["Red, DNS, Key Vault", "IaC + delete locks; purge protection", "IaC + locks", "IaC", "IaC"],
  ];
  table(s, ["Servicio", "PROD crítico", "PROD no crítico", "CERT", "DESA"], policy,
    { x: 0.6, y: 1.7, w: 12.13, colW: [2.1, 3.5, 2.4, 2.1, 2.03], fontSize: 12, rowH: 0.55 });
  s.addNotes("La política se define por criticidad y ambiente, no por servicio aislado. No productivo se reconstruye por IaC salvo excepción documentada.");

  // Roadmap
  s = pres.addSlide({ masterName: "CONTENT" });
  title(s, "Hoja de ruta de 90 días", "Primero cerrar riesgos sin costo, luego capturar ahorro");
  const phases = [
    ["30 días", HEX.red, ["Delete locks en zonas DNS y red core PROD", "Purge protection en el Key Vault PROD pendiente", `Plan para las ${fmt(risk("R1").count)} VMs PROD sin backup y las ${fmt(risk("R2").count)} con backup fallido`, "Validar huérfanos con dueños de aplicación"]],
    ["60 días", HEX.orange, [`Detener y eliminar ${fmt(o.total)} ítems huérfanos`, "Política GFS en VMs con 360 diarios", "Inmutabilidad + MUA en bóvedas PROD", "Backup continuo en Cosmos PROD"]],
    ["90 días", HEX.navy, ["Política mínima como Azure Policy por ambiente", "Prueba anual de restauración por servicio", "Reporte mensual de costo de backup por bóveda", "Revisión de Redis Premium y planes en no productivo"]],
  ];
  phases.forEach((p, i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.8, 3.9, 4.2);
    s.addShape("roundRect", { x: x + 0.25, y: 2.0, w: 1.6, h: 0.5, rectRadius: 0.1, fill: { color: p[1] }, line: { color: p[1] } });
    s.addText(p[0], { x: x + 0.25, y: 2.0, w: 1.6, h: 0.5, fontSize: 15, bold: true, color: HEX.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
    bullets(s, p[2], x + 0.25, 2.75, 3.45, 3.8, 14);
  });
  s.addNotes("Los primeros 30 días son cambios sin costo y de bajo riesgo. Los ahorros se capturan en 60 días, una vez validados con los dueños.");

  // Decisions
  s = pres.addSlide({ masterName: "TITLE_DARK" });
  s.addText("Decisiones que se necesitan", { x: 0.8, y: 0.7, w: 11, h: 0.8, fontSize: 34, bold: true, color: HEX.white, margin: 0, isTextBox: true });
  [
    "Aprobar la eliminación de backups huérfanos tras validación con dueños",
    "Adoptar la política mínima por ambiente y el esquema GFS para VMs",
    "Definir con Cumplimiento qué datos requieren retención de 10 años y a qué nivel (dato, no imagen)",
    "Priorizar inmutabilidad, CRR y locks en PROD antes de recortar",
  ].forEach((t, i) => {
    badge(s, 0.8, 1.95 + i * 1.2, String(i + 1), HEX.orange);
    s.addText(t, { x: 1.5, y: 1.85 + i * 1.2, w: 10.8, h: 0.7, fontSize: 19, color: HEX.white, valign: "middle", margin: 0, isTextBox: true });
  });
  s.addNotes("Cierre con las decisiones concretas. El Excel anexo tiene el detalle por recurso para ejecutar cada una.");

  sourcesSlide(pres);
  return pres;
}

async function main() {
  fs.mkdirSync(outDir, { recursive: true });
  const decks = [["01_backup_situacion_actual.pptx", deckCurrent()], ["02_backup_insights_oportunidades.pptx", deckInsights()]];
  for (const [name, pres] of decks) {
    const file = path.join(outDir, name);
    await pres.writeFile({ fileName: file });
    if (applyThemePath) {
      const { applyTheme } = require(applyThemePath);
      await applyTheme(file, THEME);
    }
    console.log(file);
  }
}

main().catch((error) => { console.error(error); process.exit(1); });
