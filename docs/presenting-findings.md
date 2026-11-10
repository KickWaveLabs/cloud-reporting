# Presenting findings

How to turn a reporting topic (backup, AKS vs ACA, cost) into a deck that people act on, and how to
build it with Microsoft 365 Copilot on the corporate laptop. Each topic has its own outline and
prompts in `<topic>/docs/`.

## Order

1. **Evidence of current cost.** Screenshots of Cost Management, as Azure shows them. They prove the
   numbers are the bank's own billing, not an estimate.
2. **Insights for production.** What drives the cost and where protection or resilience is missing.
3. **Insights for non-production (DESA, CERT).** What is spent or configured beyond what the
   environment needs, and which policy CERT follows.
4. **Comparison or decision.** Alternatives, on production scope only, with the break-even point.
5. **Roadmap and decisions requested.**
6. **Annex.** Where every figure comes from and how to check it.

An executive summary slide can open the deck: total monthly cost, the main saving and the main risk.

## What makes a slide land

- The title is the conclusion with a number: "US$9k a month pays for backups of VMs that no longer
  exist", not "Orphan backup items".
- One message per slide. If a slide needs two titles, it is two slides.
- Money and risk together. A saving without a risk is a nice-to-have; a risk with no cost is easy to
  postpone. Put both on the slide when they exist.
- Say what does not move the needle, so nobody spends effort there.
- End with decisions someone can approve, with an owner and a date.

## How to prove each figure

- **Measured vs estimated.** Label every number: measured (billing or inventory), derived (billed
  cost / billed quantity), estimated (rule or assumption), vendor data.
- **Source on the slide.** A footnote with the Cost analysis view, the workbook sheet and row, or
  the script that produced it.
- **Reproducible.** The scripts are read-only and live in the repository; anyone can run them again
  and get the same figure for the same month.
- **Checkable in the portal.** For each cost, the Cost analysis view (scope, filters, group by) that
  shows the same number.
- **Assumptions are visible.** Keep them in one sheet, editable, and list them in the annex.

## Copilot workflow

Copilot writes from what you give it. It does not know the estate and can invent numbers if the
prompt leaves gaps.

1. Put the workbook and a short brief (Word) in OneDrive, so Copilot can open them.
2. In Copilot chat, attach the workbook and ask for a brief with the figures per section. Check every
   number against the workbook before going on.
3. In PowerPoint, create the deck from the brief ("Create a presentation from a file").
4. Refine slide by slide with short prompts. Add the Cost analysis screenshots by hand.
5. Ask Copilot to list every number in the deck with its slide, and check that list against the
   workbook.

Rules for every prompt:

- "Use only the figures in the attached file. If a figure is missing, write [DATO PENDIENTE]."
- Give the slide title you want; Copilot's default titles are descriptive, not conclusions.
- Ask for speaker notes with the source of each figure.
