---
name: tazrim-analysis
description: >
  Turns Israeli household bank and credit-card exports into a monthly cash-flow (תזרים) analysis:
  parses bank account (עובר ושב) exports (.xlsx / .xls), credit-card monthly statements
  (פירוט אשראי, .xlsx / .xls), a benefits-club export (.docx) and P2P transfer apps (BIT / PayBox);
  classifies every transaction; builds a formula-driven Excel workbook with a "לסיווג ידני" sheet
  for the user's own labels, an interactive dashboard.html and a Hebrew RTL PDF report. Use for
  "ניתוח תזרים", "ניתוח הוצאות והכנסות", "cash-flow analysis", "Israeli bank exports",
  "פירוט אשראי", "עובר ושב", "כרטיס אשראי", "BIT", "לסיווג ידני", any number of months, at
  overview / standard / deep level.
license: Personal-Use
compatibility: >
  Python >= 3.8 with openpyxl, xlrd, pandas, python-docx, matplotlib (pip install -r requirements.txt);
  Microsoft Excel 365 to open the workbook (dynamic arrays); macOS + Excel for the recalculation
  gate (pandas-only fallback elsewhere, reported as skipped); Google Chrome or Chromium for the PDF.
metadata:
  version: 0.1.4
  language: he/en
  author: Orr Zwebner
allowed-tools: Bash, Read, Write, Edit, Glob, Grep, AskUserQuestion, WebSearch
---

# tazrim-analysis

Build a monthly cash-flow (תזרים) analysis of an Israeli household from its own exports: parse
every statement into one normalized schema, classify every row, build a workbook whose every
aggregate is an Excel formula over a `database` sheet, verify it against an independent pandas
recomputation, then write the summary, figures, Hebrew RTL report and single-file dashboard.
User-facing text (sheets, headers, notes, report, dashboard, questions) is Hebrew, RTL. Never
modify a source export; never invent a row; send every transaction-level doubt to the
"לסיווג ידני" sheet, not to chat. Scripts live in
`${CLAUDE_PLUGIN_ROOT}/skills/tazrim-analysis/scripts/` (call that `$S` below); each takes
`--config <path>` (default `./tazrim.config.json`), `--project-dir`, `--level`, `--help`, prints
diagnostics on stderr and exactly one JSON object on stdout; `{"ok": false, "error", "hint"}` on failure.

## 1. Where to save which files

The user's prompt may name two folders: where the exports already are (`<input folder>`) and where
the outputs should go (`<project>`). They need not be the same folder. Treat `<project>` as the project
folder (it holds `tazrim.config.json`, `rules/`, `work/`, `outputs/`); if the exports live elsewhere,
do NOT move them — point every `files` glob in the config at their absolute path (globs may be
absolute) and keep the `inputs/` tree below only as the reference layout. If the user names one
folder only, it is both. Create this tree in `<project>` before anything else:

```
<project>/
├── tazrim.config.json
├── inputs/
│   ├── bank/<account-id>/       one עובר ושב export per account covering the whole window (.xlsx / .xls)
│   ├── cards/<card-id>/         one פירוט אשראי file per billing cycle, named <card-id>_MMYY.* or <last4>_MMYY.*;
│   │                            MMYY = the month AFTER the charge; INCLUDE the cycle after the window (FX rows)
│   ├── benefits/<program>/      the club's purchase-detail export (.docx pasted from the site)
│   ├── p2p/<app>-<person>/      BIT / PayBox screenshots, or a filled <app>_<person>.csv in the normalized schema
│   └── template/                optional: only its category list is read (groups col A, sub-categories col B); never copied
├── rules/                       merchant_rules.csv, user_labels.csv, user_labels_interpreted.csv, fx_overrides.csv
├── work/                        normalized/, database.csv, summary.json, findings.json, figures/, boi_rates/, report.html
├── notes/                       decisions.md, questions_and_answers.md
└── outputs/                     תזרים.xlsx, dashboard.html, דוח.pdf
```

Tell the user what to download, per source. Every click-path below is **ASSUMPTION — verify
once** with the user (the method was built from files that already existed, not from the sites):

- **עובר ושב (bank)** — one export per account, for the whole window plus the first days of the
  following month (so the last card debit is visible): the account-activity page's Excel export
  (`.xlsx` or `.xls`). Save as `inputs/bank/<account-id>/<anything>.xlsx|.xls`; match the layout to a
  format id in `references/bank-formats.md` (format A: sheet `עובר ושב`, header row 8, newest first;
  format B: sheet `Activities`, header row 6, opening-balance row). ASSUMPTION — verify once.
- **פירוט אשראי (cards)** — one file per billing cycle, from the issuer's "עסקאות לחיוב בחודש X"
  page, for every cycle in the window **and the cycle after it** (its FX rows settle inside the
  window). Save as `inputs/cards/<card-id>/<card-id>_MMYY.*` (or `<last4>_MMYY.*`), e.g.
  `1234_0225.xlsx` for the cycle charged in January 2025; a cycle file that lists all of an issuer's
  cards goes under one folder for that issuer (e.g. `inputs/cards/card_cycle/`). The name is a
  convention for humans only — the parsers read the cycle from the file's own header. Two layouts
  are supported (`card_cycle_xlsx`: one file per cycle with all cards, totals in rows 7–8, header row 9;
  `card_blocks_xls` / `card_detail_xlsx`: one statement per card). ASSUMPTION — verify once.
- **Benefits club** — the purchase-detail page pasted into a Word document
  (`.docx`, paragraphs only), one per program: `inputs/benefits/<program>/*.docx`. ASSUMPTION — verify once.
- **BIT / PayBox** — either screenshots of the app's transaction list (`inputs/p2p/<app>-<person>/*.png`)
  that you transcribe into `<app>_<person>.csv` in the normalized schema (see
  `references/bank-formats.md`, "P2P CSV"), or a CSV the user already filled. ASSUMPTION — verify once.
- **Template workbook (optional)** — the household's existing תזרים file; copied, never edited in place.

Then inventory: list the tree; dump the first rows of every workbook (openpyxl / xlrd /
python-docx) to confirm sheet, header row, date encoding and sign; read each card file's own
"N עסקאות לחיוב בחודש X" / `לחיוב ב-dd.mm` header to learn **which cycle it holds** — the file
name is not the charge month; md5 all card files (byte-identical duplicates = one cycle exported
twice); compare bank card-debit rows against the files to find missing cycles; check the tools
named in `compatibility` and state explicitly which are absent.

## 2. Choose the level and the mode

| | overview | standard | deep |
|---|---|---|---|
| Sheets | `database`, `עזר_חודשי`, `הוצאות`, `הכנסות`, `ניתוח נתונים` | + `פילוח`, `פירוט לפי קטגוריה`, `פירוט עסקאות`, `הוצאות משתנות לפי בית עסק`, `לסיווג ידני`, `העברות, BIT ו-PAYBOX`, `רשימות` (hidden) | + `קטגוריות מוצעות` (secondary scheme) |
| Unknown rows | lumped into "אחר / לא מזוהה"; no manual loop | every unknown → "לסיווג ידני" row with a `נדרש:` note; one re-label round | multiple rounds; free text interpreted into `user_labels_interpreted.csv`; WebSearch on unknown merchants |
| P2P / benefits / trips | skipped unless configured | BIT matching if CSV/screenshots given; benefit programs if configured | all, incl. per-trip blocks |
| Report / dashboard | `summary.json` + short report (4 sections), no dashboard | full 8-section report + dashboard | + per-person income split + "ממצאים" from `findings.json` |

Modes (orthogonal to level):
- **A — first run**: config + exports in `inputs/` → everything in `outputs/`.
- **B — re-label**: the user filled "למילוי משתמש (טקסט חופשי)" / "קטגוריה (בחירה מהרשימה)" in
  "לסיווג ידני" and saved → rerun from classify; labels survive.
- **C — add month / extend window / replace a file**: rerun from the parsers; ids are stable
  (derived from source file + row), rows outside the window stay with `in_window=לא`.

Pick the level with the user (question set below); default `standard`.

The fixed schemas every step relies on (`python3 $S/common.py` prints them):

- **Normalized CSV** (`work/normalized/<id>.csv`, one per account / card / program / P2P entry):
  `id, source, card, original_name, txn_date, charge_date, amount_ils, orig_currency, orig_amount,
  details, source_file, row_ref` — `amount_ils > 0` = money out, `< 0` = money in; ids are
  `<PREFIX>-<6 hex>` from (source file basename, row ref), stable across re-parses.
- **database.csv** adds `pay, person, name_clean, type, group_tz, cat_tz, group_new, cat_new,
  month, month_name, amount, in_window, summed, rule_note, linked_id, trip`; `amount` is positive
  for expenses **and** incomes — `type` carries the direction; refunds stay negative.
- **Row types** (`type`): `הוצאה`, `הכנסה` (the only summed types), `העברה פנימית`,
  `חיסכון והשקעות`, `תשלום כרטיס אשראי`, `כפילות`, `הוצאה בהחזר`, `החזר הוצאה`.
- **P2P transcription** (no parser exists): read each screenshot, write one row per transfer in the
  normalized schema with `source = "<app> <person label>"`, `card = "<app>"`, `original_name` = the
  counterparty, `details = "<note> | <status> | יוצא/נכנס"`, `source_file` = the image name, outgoing
  positive, incoming negative, "Withdrawal to bank" rows kept (they become internal). If a member's
  app data is unavailable, their unmatched card rows go to "<person> – העברה לא מזוהה (BIT/PayBox)"
  and the report asks for the screenshots.

## 3. Setup and the methodology questions

1. `cp $S/../assets/config.example.json <project>/tazrim.config.json` and edit it — every key is
   documented in `references/config-schema.md` (load it now). Ids of banks/issuers must be
   `bank_xlsx_a` / `bank_xls_b` / `card_cycle_xlsx` / `card_blocks_xls` / `card_detail_xlsx` / `club_docx` (`python3 $S/formats.py` lists them).
2. `cp $S/../references/merchant-rules-starter.csv <project>/rules/merchant_rules.csv`; the file
   grows per household. Card-debit, FX-transfer and P2P-withdrawal rules are generated from config
   (`bank_debit_pattern`, `fx_settlement.transfer_pattern`, `bank_withdrawal_pattern`) — do not
   hand-write them. Category names must never contain a comma.
3. `pip install -r ${CLAUDE_PLUGIN_ROOT}/requirements.txt`.
4. Ask the methodology questions — **one AskUserQuestion round, at most 8 questions**, each
   self-contained per `references/question-policy.md` (load it before asking), recommended option
   first:
   1. window bounds (first day of the first month .. last day of the last month; never include the
      first day of the following month);
   2. month rule — bank charge month (recommended; only supported value) vs transaction month;
   3. analysis level — overview / standard / deep;
   4. treatment of each transfer class present in the data (own FX sub-account, pension/gemel,
      securities, transfers to other own accounts not in scope, P2P withdrawals, club loads) —
      internal vs expense, one question per class;
   5. what to do with a missing or duplicate cycle file — re-download (recommended) vs reconstruct
      from the bank debit (`--reconstruct-missing-cycles`, synthetic row, flagged);
   6. statistics over months with no charge — count as 0 (recommended; "ממוצע ללא אפס" shown alongside);
   7. category scheme — show the default scheme (`references/category-scheme-default.csv`) or the
      template's category list plus proposed additions **before** asking approval; never adopt
      unseen. The template workbook itself is never copied into the output — only its scheme is read.
   Log every answer in `notes/decisions.md` (date, decision, reason, what it touches) and in
   `notes/questions_and_answers.md`.

## 4. Run order per mode

`python3 $S/run_pipeline.py --mode A|B|C [--level L] [--from step] [--to step] [--dry-run]`
runs the steps below as subprocesses and stops at the first failure. Prefer step by step on a
first run so you can read each JSON before continuing.

| Step | Script | Continue only when the JSON shows |
|---|---|---|
| parse bank | `parse_bank.py` | `balance.breaks == 0` (F1/F2) for every account; read `descriptions` — it drives rule writing |
| parse card cycle files | `parse_card_cycle.py [--offline]` | header strings recorded per file; every non-zero `reconciliation` diff explained (FX rows settle outside the debit); `--` rows skipped and listed; carry-overs and duplicates counted |
| parse per-card statements | `parse_card_blocks.py [--ignore f] [--reconstruct-missing-cycles]` | every block `self_check` passes; bank reconciliation per block; `missing_cycles` empty or agreed with the user |
| parse club docx | `parse_benefits.py` | monthly reconciliation: bank debit < face value; discount rows emitted |
| classify | `classify.py` | totals by type; `unknown` count; `p2p_pairs`; `categories_not_in_scheme` empty |
| build workbook | `build_excel.py [--database] [--out]` | `sheets` match the level and start with `הוצאות`; `dropped_categories` are zero-total rows only; `preserved_labels` on a rebuild |
| transfers sheet | `add_transfers_sheet.py` | three sections with subtotals; `flagged` payees reviewed (standard/deep only; `skipped` at overview) |
| verify | `verify.py [--excel-recalc auto\|always\|never] [--keep-temp]` | no check `fail`; `excel_recalc.ran` true on macOS+Excel, else the Excel checks are `skipped` (say so — never call them passed) |
| summary | `make_summary.py [--strict]` | `recon.residual_ok` true (|residual| < 1); otherwise fix the component that drifted before going on |
| figures | `make_figures.py` | `n` = 5 / 11 / 14 per level; `font` found |
| report | `make_report_html.py [--strict]` | `findings_present` true at standard/deep; `residual_ok`; then `bash $S/render_pdf.sh work/report.html outputs/דוח.pdf` |
| dashboard | `build_dashboard.py` | `source == workbook`; open via `python3 -m http.server --bind 127.0.0.1` and read the browser console |

Mode B = `apply_user_labels.py` (§5) or `run_pipeline.py --mode B` (from classify). Mode C =
`--mode C` (from the parsers) after adding/replacing files; then re-validate `rules/user_labels*.csv`
ids against the new `database.csv`. Hand nothing over before verify shows no failure and the
residual is 0; `python3 $S/make_fixtures.py --out <dir>` builds a synthetic project to test the
install without real data.

## 5. The "לסיווג ידני" loop

- Never ask a transaction-level question in chat. Every in-window expense / income /
  reimbursable row whose category is unknown or whose `rule_note` carries `נדרש:` / `לבדיקה` /
  `לאימות` lands in the sheet automatically with all database columns, a red `נדרש:` column saying
  which decision is missing, and two green user columns at the far end: free text and a category
  dropdown (all scheme categories). When you find an ambiguous row yourself, set its `rule_note`
  to `נדרש: <decision>` through a rule or an override — do not ask.
- The user fills either column, saves, closes the workbook. Then:
  `python3 $S/apply_user_labels.py [--no-run]` → merges the sheet into `rules/user_labels.csv`
  (`id,user_text,user_cat,timestamp`; the sheet wins over older entries, old ids are kept) and reruns
  classify → build → transfers → verify → summary → figures → report → dashboard for the level.
- Deep level: read the free text and write `rules/user_labels_interpreted.csv`
  (`id,type,group_tz,cat_tz,group_new,cat_new,trip,name_clean,note`) — an explicit per-id override;
  the raw text stays in the note. Interpretation rules: a country word → `trip` = the configured trip
  covering that date (else the vacation category with the country in the note); a person's name
  with "החזר" → reimbursement / P2P income; "בייביסיטר", "שכר דירה", "גן" etc. → the matching
  scheme category; a PAYBOX/BIT note naming a purpose → that purpose's category; "פנימי"/"שלנו" →
  `העברה פנימית`; anything you cannot map → leave the row in the sheet with a sharper `נדרש:`.
- Unknown merchants (deep): batch them, WebSearch Hebrew and English (add city / "בע"מ" /
  "כרטיס אשראי חיוב"); record what / proposed category / confidence / URL in
  `notes/merchant_research.md`; write "לא נמצא" rather than guess; a merchant may be registered
  under its street address.
- After every round: the sheet shows only still-open rows; labels and category choices survive
  the rebuild (nothing else typed into the workbook does — it is rebuilt from scratch). Tell the
  user to close Excel **without saving** before reopening.

## 6. Findings and report

After `make_summary.py` succeeds, read `work/summary.json` and write `work/findings.json`:

```json
{"findings": ["…" | {"title": "…", "text": "…"}],
 "actions": [{"what": "…", "why": "…", "how": "…"}],
 "data_gaps": ["…"]}
```

Rules: 6–8 findings, every number quoted from `summary.json` (no arithmetic of your own, no
constants); each action has all three fields; `data_gaps` lists missing cycles, members without
P2P data, estimated conversions, skipped checks. Then `make_report_html.py`, then
`render_pdf.sh` (exit 3 = no Chrome: deliver the HTML and say no PDF was produced). Regenerate the
HTML immediately before every print. Charts of averages are captioned "ממוצע חודשי", never N-month sums.

## 7. Hand-over template (end of every round)

```
| גיליון | מה השתנה |
|---|---|
| הוצאות (הלשונית הראשונה) | … |
| לסיווג ידני | N שורות פתוחות (היו M) |
מספרים: הכנסה ממוצעת ₪X · הוצאה ממוצעת ₪Y · מאזן ₪Z · חיסכון בפועל ₪W  (all from summary.json)
אימות: verify PASS · 0 שגיאות נוסחה · סה"כ תואם pandas · residual 0 · Excel recalculation: ran|skipped
קבצים: outputs/תזרים.xlsx · outputs/dashboard.html · outputs/דוח.pdf (or "PDF not produced")
⚠ סגור את ה-Excel בלי לשמור לפני פתיחה מחדש (העותק בזיכרון ידרוס את הבנייה).
```

## 8. Judgment calls that decide routing

- **Month = bank charge month.** `month = charge_date[:7]`; a purchase early in the month counts
  in the next cycle. Window flags: `in_window = start ≤ charge_date ≤ end`; `summed = in_window ∧
  type ∈ {הוצאה, הכנסה}`; recomputed as the last step of classify.
- **Transfer vs expense** (verbatim rule): a bank transfer / cheque is **not** automatically
  internal: rent standing order, kindergarten cheques, advisors are expenses; internal only when
  the money demonstrably stays in the family's own accounts (pension/gemel deposits, securities,
  own-account transfers, FX conversions, P2P withdrawals to bank, cash withdrawals that fund
  estimates, club loads). Transfers to the household's own FX sub-account are internal; only the
  exchange fee is an expense; the real expense is the card's FX rows.
- **Card debits are never summed** (`type = תשלום כרטיס אשראי`); the statement rows carry the
  expense. Club purchases at face value plus one computed negative discount row per month.
- **Never invent rows.** A cycle without a file is reported and re-downloaded; zero-total
  categories are dropped from the stat sheets; categories are added only with approval.
- **Unknown → לסיווג ידני**, never AskUserQuestion. Chat questions are for methodology only.
- **Empty months count as 0**; every average has "ממוצע ללא אפס" adjacent — in the stat sheets,
  the analysis tables, the pivot KPIs, the per-category blocks and the dashboard.
- **Group subtotal rows** after each group (formulas `+` over member rows), excluded from the
  rows-check; groups sorted by total desc, sub-categories by average desc.
- **P2P**: card-funded payments count on the app payment date; the card row becomes `כפילות`
  with `linked_id`; incoming P2P from friends is income ("החזרים מחברים"), not netted.
- **Payment-method labels always carry the owner** in parentheses; a joint account uses the
  shared label. Card → person mapping is confirmed with the user, not inferred silently.
- **Claims in the report are checked against the data** (e.g. a fee percentage) — cards linked to
  an FX account pay no conversion fee in that currency.

## 9. Failure-modes checklist (run through before every hand-over)

- Header assertion failed → the institution changed its export; fix `formats.py` only
  (`references/bank-formats.md`), never the parser's caller.
- `ValueError '--'` in a cycle file → not-yet-posted rows; they are skipped and listed — confirm the count.
- Charge date in the wrong column → cycle-file columns shift; the parser locates it by pattern.
- Two card files byte-identical → one cycle exported twice; ask for the missing one.
- Balance chain breaks → wrong sign or a skipped row; do not continue.
- Reconciliation diff ≠ 0 per cycle → FX rows settle outside the debit; otherwise a missing block.
- `#NAME?` everywhere → `_xlfn.` / `_xlpm.` / `_xlfn._xlws.` prefix missing (`references/workbook-spec.md`).
- KPI shows 0 / dynamic table blank → SUMPRODUCT / LET not written as `ArrayFormula`.
- Excel refuses the file silently → a `"` inside a category name inside a formula; quotes are doubled; bisect.
- Rules stop matching → a comma in a category name.
- Override lands on the wrong row after adding a file → ids must be stable; re-validate `user_labels*.csv`.
- verify FAIL while checks print pass → total row matched a subtotal prefix; labels are exact.
- "Excel did not save the calculated copy" → temp copy still open; close and delete `work/verify_calc.xlsx`.
- Residual ≠ 0 after a reclassification → a component definition no longer matches the types;
  rerun the identity after every rule change.
- "You added rows on your own" → zero-total template rows rendered; drop them.
- Rebuilt workbook "does not show the new sheets" → Excel's in-memory copy (`~$` lock file); close without saving.
- Dashboard `ReferenceError` / blank page → `node --check` the script; serve over `http.server`, read the console.
- `KeyError` in the report after a rename → narrative must come from `findings.json`, tables key-driven.
- Excel dialog during verify → alerts must be off; never depend on System Events.
- Figures show boxes instead of Hebrew → font search list; install Noto Sans Hebrew on Linux.

## 10. References (load on demand)

- `references/bank-formats.md` — every export format (sheet, header row, exact headers, dates,
  sign, block markers, cycle derivation, quirks), the BOI rate cache and the P2P CSV template.
  Load when parsing, when a header assertion fails, or when transcribing screenshots.
- `references/config-schema.md` — every key of `tazrim.config.json` with type, default, example
  and the script that reads it. Load when creating or validating the config.
- `references/analysis-method.md` — the rules and formulas: month rule, row types and the
  no-double-count table, transfer-vs-expense, FX settlement and fees, statistics, the cash
  reconciliation identity, reimbursables, estimates, trips, merchant research. Load during
  classification, when the residual is not 0, and before writing `findings.json`.
- `references/workbook-spec.md` — sheet-by-sheet spec of the workbook, named ranges, formula
  prefixes, ArrayFormula rules, colour scales, hyperlink rule, preserved columns, verify checks,
  Excel/AppleScript gotchas. Load when building or debugging the workbook.
- `references/question-policy.md` — when AskUserQuestion is allowed and its exact form, with
  good/bad examples. Load before any question.
- `references/category-scheme-default.csv` — the household-neutral scheme (`group,cat,fixed`)
  used when no template is given. Load when building or approving the scheme.
- `references/merchant-rules-starter.csv` — the shipped rules (national chains, utilities,
  statutory bodies, generic bank descriptions). Load at `rules/merchant_rules.csv` creation.
