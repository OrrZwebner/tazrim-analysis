# tazrim-analysis

A Claude skill that turns your Israeli bank and credit-card exports (עובר ושב, פירוט אשראי) into a
monthly cash-flow (תזרים) analysis: a formula-driven Excel workbook, an interactive HTML dashboard
and a Hebrew PDF report — for any number of months, at three depth levels (overview / standard /
deep). Everything runs locally with Claude Code.

## Install

- **claude.ai / Claude app:** download `tazrim-analysis.zip` from the [latest release](../../releases/latest) and upload it in **Settings → Capabilities** (Skills → Upload skill).
- **Claude Code (plugin):** `/plugin marketplace add OrrZwebner/tazrim-analysis` then `/plugin install tazrim-analysis`.
- **Claude Code (manual):** copy `skills/tazrim-analysis/` into `~/.claude/skills/`.

## Usage

1. Download the exports from your bank and credit-card sites into one folder: one עובר ושב export per
   account for the period, and one monthly פירוט אשראי file per billing cycle.
2. Open Claude Code in that folder.
3. Paste this prompt (fill in the placeholders):

```
נתח את התזרים שלי לפי הקבצים בתיקייה הזו לתקופה <חודש התחלה>–<חודש סיום>, ברמה <overview / standard / deep>
```

The skill guides the folder layout, asks a few methodology questions once, and builds the outputs.

## Outputs

- `outputs/תזרים.xlsx` — the workbook: expenses and incomes by category × month with averages,
  interactive filtering, and a "לסיווג ידני" sheet where you label what it could not classify
  (your labels survive every rebuild).
- `outputs/dashboard.html` — a self-contained interactive dashboard; open it in any browser.
- `outputs/דוח.pdf` — a Hebrew report of the period.

## Requirements

- Python ≥ 3.8 and `pip install -r requirements.txt`.
- Microsoft Excel 365 to open the workbook; Google Chrome for the PDF.

## Privacy

Everything runs on your machine; the only network call is the Bank of Israel rate API.
The workbook, dashboard and report contain your data — never commit or share them.

## License

[Personal Use License](LICENSE) © 2026 Orr Zwebner. Free for private, personal use — analysing
your own household's finances — and you may modify it as you like for yourself. Anything beyond
that (commercial use, reselling, offering it or its outputs as a service, redistribution) needs
written permission. Provided as is, without warranty; this is not financial advice.
