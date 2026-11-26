#!/usr/bin/env python3
"""Build the AKS right-sizing deck (Spanish): the real pain is how AKS is used, not AKS.

Five slides from the same figures as the right-sizing workbook (aa_rightsizing.py, default variables):
  1  the real pain: capacity paid and not used, few well-used clusters, the break-even hours
  2  well-used against badly used AKS: Container Apps against AKS by class, with examples
  3  where the money goes: minimum clusters, idle capacity, system pool, empty clusters, non-production, autoscaling
  4  pros and cons: optimized AKS, shared AKS and Container Apps
  5  what to do: the levers and the size of the problem each one addresses
Reads only results/. Output: aks_rightsizing.pptx in the newest reports/<date>-aks-aca folder (or --output).

Usage: python aks-aca/src/aa_deck_rightsizing.py [--results DIR] [--output FILE]
"""

import argparse
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
    cards = [(d.percent(d.share_of(idle, total)), f"de lo que pagamos en AKS de negocio no se usa: {pen(idle)} al mes",
              d.RUST),
             (f"{classes['OPTIMIZED']['clusters']} de {count}", "clusters aprovechan bien lo que pagan (60% o más)",
              d.ORANGE),
             (f"{hours:.0f} h", f"al día: una vCPU de Container Apps cuesta {ratio:.0f} veces la de AKS con reserva; "
                                "solo gana si está encendida menos que eso", d.GREEN)]
    width, gap = Inches(3.9), Inches(0.2)
    for index, (value, label, colour) in enumerate(cards):
        d.kpi(slide, Inches(0.6) + index * (width + gap), Inches(1.4), width, Inches(1.9), value, label,
              value_color=colour, value_size=40, label_size=12)
    d.callout(slide, Inches(0.6), Inches(3.6), Inches(12.1), Inches(1.0),
              "Un AKS bien usado es más barato que Container Apps. El ahorro aparece donde AKS paga capacidad que no "
              "usa.", size=16, fill=d.CARD)
    d.text_box(slide, Inches(0.6), Inches(4.85), Inches(12.1), Inches(1.0),
               ["• Uso: CPU o memoria usada en promedio en horario de oficina, sobre lo que ofrecen sus servidores.",
                "• Aplicaciones de negocio en producción; costo de octubre 2026 con IGV."], size=12, color=d.GREY)


def slide_classes(deck, summary, source):
    classes = summary["classes"]
    slide = d.base_slide(deck, "AKS bien usado contra mal usado: Container Apps solo gana en los sobredimensionados",
                         source)
    table = [["Cluster", "Clusters", "AKS hoy / mes", "No se usa / mes", "Container Apps sería"]]
    for kind, label, meaning, _ in KINDS:
        entry = classes[kind]
        ratio = entry["aca"] / entry["aks_costed"] if entry["aks_costed"] else 0
        verdict = f"{ratio:.1f} veces más caro" if ratio >= 1 else f"{d.percent(1 - ratio)} más barato"
        table.append([f"{label} ({meaning})", d.number(entry["clusters"]), pen(entry["aks"]), pen(entry["idle"]),
                      verdict])
    d.table(slide, Inches(0.6), Inches(1.3), Inches(12.1), table, [0.32, 0.1, 0.18, 0.18, 0.22], highlight_column=4,
            row_height=0.45, size=13)
    rows = {row["cod_app"]: row for row in summary["rows"]}
    examples = [(code, rows[code]) for code in ("YAPE", "APPV", "XESS") if code in rows]
    width, gap = Inches(3.9), Inches(0.2)
    for index, (code, row) in enumerate(examples):
        aca = v1.to_float(row["aca_month_pen"]) or 0
        verdict = f"{aca / row['aks_pen']:.0f} veces más caro en Container Apps" if aca >= row["aks_pen"] else \
            f"{d.percent(1 - aca / row['aks_pen'])} más barato en Container Apps"
        d.kpi(slide, Inches(0.6) + index * (width + gap), Inches(3.5), width, Inches(1.6), code,
              f"usa {d.percent(row['use'])} · {pen(row['aks_pen'])} al mes en AKS · {verdict}",
              value_color=dict((kind, colour) for kind, _, _, colour in KINDS).get(row["kind"], d.ORANGE),
              value_size=28, label_size=11)
    d.text_box(slide, Inches(0.6), Inches(5.3), Inches(12.1), Inches(0.7),
               f"En los clusters grandes la estimación de Container Apps es alta: supone que cada {MICRO_ONE} recibe "
               "todo el tráfico del cluster. La dirección no cambia: mientras mejor usado está el AKS, más caro sale "
               "Container Apps.", size=11, italic=True, color=d.GREY)


def slide_money(deck, summary, source):
    slide = d.base_slide(deck, "Dónde se va la plata", source)
    classes = summary["classes"]
    idle = sum(entry["idle"] for entry in classes.values())
    table = [["Dolor", "Qué pasa", "Cuántos", "Al mes"],
             ["Mínimo por cluster", f"Clusters dedicados de {v1.MIN_AKS_NODES} servidores o menos: pagan el piso "
                                    "aunque corran poco", f"{d.number(summary['minimum'][0])} clusters",
              pen(summary["minimum"][1])],
             ["Capacidad que no se usa", "Lo que se paga por servidores y no usan las aplicaciones",
              f"{d.number(sum(entry['clusters'] for entry in classes.values()))} clusters", pen(idle)],
             ["Servidores de sistema", "Parte de la memoria que se usa en los clusters es del propio Kubernetes",
              d.percent(d.share_of(summary["system_mem"], summary["mem"])), "incluido arriba"],
             ["Clusters vacíos", "Producción creada antes de tener aplicación",
              f"{d.number(summary['empty'][0])} clusters", pen(summary["empty"][1])],
             ["No productivos 24/7", "Desarrollo y certificación encendidos todo el mes",
              f"{d.number(summary['nonprod'][0])} clusters", pen(summary["nonprod"][1])],
             ["Autoescalado quieto", f"{v1.MICRO_CAP} que no cambian de réplicas hace más de 3 meses",
              f"{d.number(summary['hpas_still'])} de {d.number(summary['hpas'])}", "-"]]
    d.table(slide, Inches(0.6), Inches(1.3), Inches(12.1), table, [0.2, 0.52, 0.13, 0.15], highlight_column=3,
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


def slide_levers(deck, summary, aca_saving, source):
    classes = summary["classes"]
    idle = sum(entry["idle"] for entry in classes.values())
    slide = d.base_slide(deck, "Qué hacer: cada palanca ataca un dolor", source)
    table = [["Palanca", "Ataca", "Tamaño del dolor / mes"],
             ["Ajustar requests y servidores al uso real", "Capacidad que no se usa", pen(idle)],
             ["Juntar aplicaciones chicas en un AKS compartido", "Mínimo por cluster y servidores de sistema",
              pen(summary["minimum"][1])],
             ["Apagar no productivos fuera de horario", "No productivos 24/7", pen(summary["nonprod"][1])],
             ["Container Apps para aplicaciones chicas que duermen de noche", "Mínimo por cluster y noches",
              f"ahorro estimado {pen(aca_saving)}"],
             ["Crear producción solo cuando la aplicación esté lista", "Clusters vacíos", pen(summary["empty"][1])]]
    d.table(slide, Inches(0.6), Inches(1.3), Inches(12.1), table, [0.45, 0.33, 0.22], highlight_column=2,
            row_height=0.5, size=13)
    d.note_box(slide, Inches(0.6), Inches(4.75), Inches(12.1), Inches(0.75),
               "Con reservas de 3 años, apagar o achicar servidores solo ahorra si la reserva se ajusta a lo que queda "
               "encendido.", size=12)


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
    slide_money(deck, summary, source)
    slide_options(deck, source)
    slide_levers(deck, summary, aca_saving, source)
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
