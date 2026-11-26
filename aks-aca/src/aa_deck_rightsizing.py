#!/usr/bin/env python3
"""Build the AKS right-sizing deck (Spanish): the real pain is how AKS is used, not AKS.

Five slides from the same figures as the right-sizing workbook (aa_rightsizing.py, default variables):
  1  the real pain: capacity paid and not used, few well-used clusters, the break-even hours
  2  well-used against badly used AKS: Container Apps against AKS by class, with examples
  3  where the waste is, as a share of the AKS bill: minimum clusters, idle capacity, system pool, empty clusters,
     non-production, autoscaling
  4  pros and cons: optimized AKS, shared AKS and Container Apps
  5  what each lever would save a month and a year, at most (the only slide with soles)
Reads only results/. Output: aks_rightsizing.pptx in the newest reports/<date>-aks-aca folder (or --output).

Usage: python aks-aca/src/aa_deck_rightsizing.py [--results DIR] [--output FILE]
"""

import argparse
import calendar
import glob
import os

from pptx import Presentation
from pptx.util import Inches

import aa_common as aa
import aa_deck as d
import aa_deck_exec as exec_deck
import aa_deck_exec_v1 as v1
import aa_rightsizing as rs

MICRO, MICRO_ONE = v1.MICRO, v1.MICRO_ONE
pen = exec_deck.pen
KINDS = [("OPTIMIZED", "Bien aprovechado", "usa 60% o más", d.GREEN),
         ("PARTLY_USED", "A medias", "usa 30% a 60%", d.ORANGE),
         ("OVERSIZED", "Sobredimensionado", "usa menos de 30%", d.RUST)]


def breakeven(results):
    """(AKS US$ per vCPU-hour reserved, Container Apps US$ per vCPU-hour with discount, hours a day)."""
    prices = v1.read_prices(results)
    price = lambda key: prices.get((key, "eastus2"), 0)
    aks = price(f"vm_{v1.MIN_AKS_VM}_{v1.MIN_AKS_RESERVATION}") / 26280 / 8
    aca = 3600 * (price("aca_vcpu_active_s") + 2 * price("aca_mem_active_gib_s")) * (1 - v1.BANK_DISCOUNT)
    return aks, aca, 24 * aks / aca if aca else 0


def slide_pain(deck, summary, hours, ratio, source):
    classes = summary["classes"]
    total = sum(entry["aks"] for entry in classes.values())
    idle = sum(entry["idle"] for entry in classes.values())
    count = sum(entry["clusters"] for entry in classes.values())
    slide = d.base_slide(deck, "El verdadero dolor no es AKS: es cómo lo usamos", source)
    cards = [(d.percent(d.share_of(idle, total)), "del gasto en servidores AKS de negocio paga CPU y memoria que no "
                                                  "se usan", d.RUST),
             (f"{classes['OPTIMIZED']['clusters']} de {count}", "clusters usan en promedio 60% o más de la CPU o "
                                                               "memoria de sus servidores", d.ORANGE),
             (f"{hours:.0f} h", f"al día: una vCPU de Container Apps cuesta {ratio:.0f} veces la de AKS con reserva; "
                                "solo gana si está encendida menos que eso", d.GREEN)]
    width, gap = Inches(3.9), Inches(0.2)
    for index, (value, label, colour) in enumerate(cards):
        d.kpi(slide, Inches(0.6) + index * (width + gap), Inches(1.4), width, Inches(1.9), value, label,
              value_color=colour, value_size=40, label_size=12)
    d.callout(slide, Inches(0.6), Inches(3.6), Inches(12.1), Inches(1.0),
              "Un AKS bien usado es más barato que Container Apps. El ahorro aparece donde AKS paga capacidad que no "
              "usa.", size=16, fill=d.CARD)
    d.text_box(slide, Inches(0.6), Inches(4.85), Inches(12.1), Inches(1.3),
               ["• Uso: CPU y memoria que usan los servidores de cada cluster, en promedio de lunes a viernes de 8:00 a "
                "20:00 (métricas de Azure Monitor de octubre 2026), sobre lo que esos servidores ofrecen. Cuenta la "
                "más llena de las dos.",
                "• Costo: Cost Management de octubre 2026, costo amortizado (las reservas repartidas en su plazo).",
                "• Clusters de aplicaciones de negocio en producción. Lo que ahorraría cada palanca va al final."],
               size=12, color=d.GREY)


def slide_classes(deck, summary, source):
    classes = summary["classes"]
    slide = d.base_slide(deck, "AKS bien usado contra mal usado: Container Apps gana donde sobra capacidad",
                         source)
    total = sum(entry["aks"] for entry in classes.values())
    verdict = lambda aca, aks: "-" if not aks else (f"{aca / aks:.1f} veces más caro" if aca >= aks
                                                    else f"{d.percent(1 - aca / aks)} más barato")
    table = [["Cluster", "Clusters", "Peso en el gasto", "Pasan a Container Apps", "Los que pasan", "Los que se quedan"]]
    for kind, label, meaning, _ in KINDS:
        members = [row for row in summary["rows"] if row["kind"] == kind]
        moving = [row for row in members if row["aca_group"] in ("ACA", "ACA_WITH_REVIEW")]
        staying = [row for row in members if row not in moving and row["aca_month_pen"] != ""]
        cost = lambda group: (sum(v1.to_float(row["aca_month_pen"]) or 0 for row in group),
                              sum(row["aks_pen"] for row in group))
        table.append([f"{label} ({meaning})", d.number(len(members)),
                      d.percent(d.share_of(sum(row["aks_pen"] for row in members), total)), d.number(len(moving)),
                      verdict(*cost(moving)), verdict(*cost(staying))])
    d.table(slide, Inches(0.6), Inches(1.3), Inches(12.1), table, [0.28, 0.09, 0.13, 0.15, 0.17, 0.18],
            highlight_column=4, row_height=0.45, size=12)
    rows = {row["cod_app"]: row for row in summary["rows"]}
    examples = [(code, rows[code]) for code in ("YAPE", "APPV", "XESS") if code in rows]
    width, gap = Inches(3.9), Inches(0.2)
    for index, (code, row) in enumerate(examples):
        aca = v1.to_float(row["aca_month_pen"]) or 0
        verdict = f"{aca / row['aks_pen']:.0f} veces más caro en Container Apps" if aca >= row["aks_pen"] else \
            f"{d.percent(1 - aca / row['aks_pen'])} más barato en Container Apps"
        d.kpi(slide, Inches(0.6) + index * (width + gap), Inches(3.5), width, Inches(1.6), code,
              f"usa {d.percent(row['use'])} de su capacidad · {verdict}",
              value_color=dict((kind, colour) for kind, _, _, colour in KINDS).get(row["kind"], d.ORANGE),
              value_size=28, label_size=11)
    d.text_box(slide, Inches(0.6), Inches(5.3), Inches(12.1), Inches(0.7),
f"Los de uso a medias que pasan están casi todos en el mínimo de {v1.MIN_AKS_NODES} servidores, y la mayor "
               "parte de la memoria que usan es del propio Kubernetes, no de la aplicación: por eso igual ahorran. "
               f"En los clusters grandes la estimación de Container Apps es alta (supone que cada {MICRO_ONE} recibe "
               "todo el tráfico del cluster), pero la dirección no cambia.", size=11, italic=True, color=d.GREY)


def slide_money(deck, summary, total, source):
    """Where the waste is, as a share of the whole AKS bill (production and non-production): no soles until the levers."""
    slide = d.base_slide(deck, "Dónde está el desperdicio", source)
    classes = summary["classes"]
    idle = sum(entry["idle"] for entry in classes.values())
    share = lambda value: d.percent(d.share_of(value, total))
    table = [["Dolor", "Qué pasa", "Cuántos", "Peso en el gasto AKS"],
             ["Mínimo por cluster", f"Clusters dedicados de {v1.MIN_AKS_NODES} servidores o menos: pagan el piso "
                                    "aunque corran poco", f"{d.number(summary['minimum'][0])} clusters",
              share(summary["minimum"][1])],
             ["Capacidad que no se usa", "Se paga por servidores que las aplicaciones no usan",
              f"{d.number(sum(entry['clusters'] for entry in classes.values()))} clusters", share(idle)],
             ["Servidores de sistema", "Parte de la memoria usada es del propio Kubernetes, no de la aplicación",
              d.percent(d.share_of(summary["system_mem"], summary["mem"])) + " de la memoria", "-"],
             ["Clusters vacíos", "Producción creada antes de tener aplicación",
              f"{d.number(summary['empty'][0])} clusters", share(summary["empty"][1])],
             ["No productivos 24/7", "Desarrollo y certificación encendidos todo el mes",
              f"{d.number(summary['nonprod'][0])} clusters", share(summary["nonprod"][1])],
             ["Autoescalado quieto", f"{v1.MICRO_CAP} que no cambian de réplicas hace más de 3 meses",
              f"{d.number(summary['hpas_still'])} de {d.number(summary['hpas'])}", "-"]]
    d.table(slide, Inches(0.6), Inches(1.3), Inches(12.1), table, [0.2, 0.5, 0.15, 0.15], highlight_column=3,
            row_height=0.5, size=12)
    d.text_box(slide, Inches(0.6), Inches(5.0), Inches(12.1), Inches(0.5),
               "Los dolores se superponen (un cluster en el mínimo también tiene capacidad sin usar): cada fila se lee "
               "sola, no se suman.", size=11, italic=True, color=d.GREY)


def slide_options(deck, source):
    slide = d.base_slide(deck, "Pros y contras de cada camino", source)
    table = [["", "AKS optimizado", "AKS compartido", "Container Apps"],
             ["Costo", "vCPU más barata con reserva; paga 24/7", "Reparte el mínimo y los servidores de sistema",
              "Paga por uso; vCPU unas 4 veces más cara encendida"],
             ["Noches y fines de semana", "Sigue pagando los servidores", "Sigue pagando los servidores",
              "Se apaga si no hay uso"],
             ["Operación", "Dimensionar y vigilar cada cluster", "Un cluster para varias aplicaciones: más gobierno",
              "Sin servidores ni versiones de Kubernetes"],
             ["Límites", "Ninguno nuevo", "Aislamiento entre aplicaciones", "4 vCPU / 8 GB por réplica; Quarkus"],
             ["Mejor para", "Carga alta y pareja", "Muchas aplicaciones chicas estables",
              "Aplicaciones chicas que duermen de noche"]]
    d.table(slide, Inches(0.6), Inches(1.3), Inches(12.1), table, [0.18, 0.27, 0.27, 0.28], row_height=0.6, size=12)


def levers(results, month, summary, aca_saving, target=0.6):
    """Saving a month of each lever, at most: right-size each cluster to the target use without going below the
    dedicated minimum; what only sharing a cluster frees (that minimum); non-production off outside office hours;
    Container Apps for the clusters the criteria move; the empty clusters."""
    soles = (1 + v1.IGV) * v1.EXCHANGE_RATE
    floor = v1.minimum_costs(month, v1.read_prices(results))[0] * soles
    own = shared = 0.0
    for row in summary["rows"]:
        if row["use"] is None or row["use"] >= target:
            continue
        needed = row["aks_pen"] * row["use"] / target
        kept = max(needed, min(row["aks_pen"], floor))
        own += row["aks_pen"] - kept
        shared += kept - needed
    year, number = (int(part) for part in month.split("-"))
    days = calendar.monthrange(year, number)[1]
    off = 1 - v1.office_seconds(month) / (days * 86400)
    return [("Ajustar requests y servidores al uso real",
             f"Llevar cada cluster a {d.percent(target)} de uso, sin bajar de su mínimo", own),
            ("Juntar aplicaciones chicas en un AKS compartido", "Lo que el mínimo de cada cluster no deja bajar", shared),
            ("Apagar no productivos fuera de horario", f"El {d.percent(off)} de las horas del mes",
             summary["nonprod"][1] * off),
            ("Container Apps para aplicaciones chicas", "Los clusters que cumplen los criterios", aca_saving),
            ("Crear producción solo cuando la aplicación esté lista", "Los clusters vacíos de hoy", summary["empty"][1])]


def slide_levers(deck, rows, source):
    slide = d.base_slide(deck, "Cuánto ahorraría cada palanca", source)
    table = [["Palanca", "Cómo", "Ahorro / mes", "Ahorro / año"]]
    for lever, how, saving in rows:
        table.append([lever, how, f"hasta {pen(saving)}", f"hasta {pen(saving * 12)}"])
    d.table(slide, Inches(0.6), Inches(1.3), Inches(12.1), table, [0.33, 0.33, 0.17, 0.17], highlight_column=2,
            row_height=0.5, size=12)
    d.note_box(slide, Inches(0.6), Inches(4.8), Inches(12.1), Inches(0.95),
               "Las palancas se superponen (las aplicaciones chicas que pasan a Container Apps ya no se juntan): no se "
               "suman. Con reservas de 3 años, apagar o achicar servidores solo ahorra si la reserva se ajusta a lo que "
               "queda encendido.", size=12)


def build(results, output):
    data, month = rs.rows(results)
    summary = rs.summary(data)
    aks_hour, aca_hour, hours = breakeven(results)
    rs.model.use("v2")
    aca_saving = sum(row["saving_month_pen"] for row in rs.model.evaluate(rs.model.load(results)) if row["goes_to_aca"])
    source = (f"Fuente: métricas de AKS y Cost Management de {v1.month_label(month)} (costo amortizado); precios de "
              f"lista East US 2. {v1.MONEY_NOTE}")
    deck = Presentation()
    deck.slide_width, deck.slide_height = Inches(13.333), Inches(7.5)
    slide_pain(deck, summary, hours, aca_hour / aks_hour if aks_hour else 0, source)
    slide_classes(deck, summary, source)
    slide_money(deck, summary, sum(v1.to_float(row["cost_month_usd"]) or 0 for row in data)
                * (1 + v1.IGV) * v1.EXCHANGE_RATE, source)
    slide_options(deck, source)
    slide_levers(deck, levers(results, month, summary, aca_saving), source)
    exec_deck.english_numbers(deck)
    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    aa.keep_version(output)
    deck.save(output)
    aa.azure.LOG.info("Wrote %s (%d slides)", output, len(deck.slides))


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--results", default=aa.RESULTS_DIR, help="Folder with the aa-<source> result folders")
    parser.add_argument("--output", default=None,
                        help="Output .pptx. Default: aks_rightsizing.pptx in the newest reports/<date>-aks-aca")
    args = parser.parse_args()
    aa.azure.setup_logging(False)
    output = args.output
    if not output:
        folders = sorted(glob.glob(os.path.join(aa.TOPIC_DIR, "reports", "*-aks-aca")))
        output = os.path.join(folders[-1] if folders else os.path.join(aa.TOPIC_DIR, "reports", "aks-aca"),
                              "aks_rightsizing.pptx")
    build(args.results, output)


if __name__ == "__main__":
    main()
