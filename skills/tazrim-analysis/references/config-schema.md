# `tazrim.config.json` — every key

Loads when creating or validating the config (`assets/config.example.json` is a complete
synthetic example; copy it to `<project>/tazrim.config.json`). All paths are relative to the
project folder (the config's folder, or `--project-dir`). `scripts/common.py` deep-merges the
`DEFAULTS` table under the user's file and validates on every script start; a violation exits 2
with `{"ok": false, "error": "<Hebrew> | <English>", "hint"}`. Globs that match no file are
warnings, not errors (mode C adds files later). Run `python3 scripts/common.py` to print the
defaults as JSON.

Derived from `common.py` (DEFAULTS, `CARD_DEFAULTS`, `BENEFIT_DEFAULTS`, `P2P_DEFAULTS`,
`ACCOUNT_DEFAULTS`, `load_config`) and from each script's reads; a key added to a script must be appended here.

Every example value is synthetic. Thresholds are design defaults (3,000 / 12,000 ₪).

## Top level

| Key | Type | Default | Example | Read by |
|---|---|---|---|---|
| `language` | `"he"` | `"he"` | `"he"` | (reserved; user-facing text is Hebrew) |
| `analysis_level` | `overview` \| `standard` \| `deep` | `"standard"` | `"deep"` | every script via `cfg.level`; `--level` overrides |
| `fixed_categories` | list of category names | `[]` | `["שכר דירה", "ועד בית"]` | build_excel (fixed vs variable, F22), make_summary; the scheme CSV's `fixed` column is the other source |

## `window` (required)

| Key | Type | Default | Example | Read by |
|---|---|---|---|---|
| `window.start` | ISO date | — required | `"2025-01-01"` | common.months, classify (in_window), make_summary, fx_rates (BOI start period) |
| `window.end` | ISO date ≥ start | — required | `"2025-02-28"` (last day of the last month — never the 1st of the next) | same |
| `window.month_rule` | `"charge_date"` | `"charge_date"` | `"charge_date"` | common (only supported value: analysis month = bank charge month) |

N months = every calendar month from start to end inclusive; any N ≥ 1.

## `household` (required)

| Key | Type | Default | Example | Read by |
|---|---|---|---|---|
| `household.label` | string | `"משק הבית"` | `"משק בית לדוגמה"` | build_dashboard (title), make_report_html, make_summary |
| `household.people[]` | list of `{id, label}`, ≥ 1, ids unique, `shared` reserved | — required | `[{"id":"p1","label":"בן/בת זוג א'"},{"id":"p2","label":"בן/בת זוג ב'"}]` | common.owner_label, classify (person / pay), build_excel, make_figures, make_summary, build_dashboard |
| `household.shared_label` | string | `"משותף (עו\"ש)"` | `"משותף"` | common.owner_label for `owner: "shared"` |

## `accounts[]` — bank accounts (parse_bank)

| Key | Type | Default | Example | Read by |
|---|---|---|---|---|
| `id` | unique string | — required | `"bank_a"` | file name `work/normalized/<id>.csv`; `settles_from` targets |
| `bank` | `bank_xlsx_a` \| `bank_xls_b` (formats.BANKS) | — required | `"bank_xlsx_a"` | parse_bank (format dispatch) |
| `label` | string | = id | `"עו\"ש חשבון א'"` | pay labels, report, dashboard |
| `owner` | person id or `"shared"` | `"shared"` | `"p2"` | classify (person), pay label suffix |
| `files` | glob or list of globs | `null` | `"inputs/bank/bank_a/*.xlsx"` | parse_bank |
| `opening_balance` | number or null | `null` | `3000.00` | parse_bank (format B, F2 when the sheet has no opening row), make_summary |

## `cards[]` — credit cards (parse_card_cycle / parse_card_blocks / classify / verify)

| Key | Type | Default | Example | Read by |
|---|---|---|---|---|
| `id` | unique string | — required | `"card_1234"` | normalized file name; summary keys |
| `issuer` | `card_cycle_xlsx` \| `card_blocks_xls` \| `card_detail_xlsx` (formats.ISSUERS) | — required | `"card_cycle_xlsx"` | run_pipeline (which parser), parse_* |
| `last4` | 4 digits string | `""` | `"1234"` | parse_card_cycle (card cell match), parse_card_blocks (block match), verify |
| `label` | string | `"<issuer> <last4>"` | `"ויזה 1234"` | pay labels everywhere |
| `owner` | person id or `"shared"` | — required | `"p1"` | classify (person), labels |
| `files` | glob(s) | `null` | `"inputs/cards/card_cycle/card_*.xlsx"` | the issuer's parser |
| `settles_from` | an `accounts[].id` | — required | `"bank_a"` | parsers (bank reconciliation, F3), classify (scope of the generated card-debit rule) |
| `bank_debit_pattern` | substring of the bank description | `null` | `"חיוב לכרטיס ויזה 1234"` (cycle file) / `"1234 - כרטיס אשראי"` (per-card statement) | classify (→ `תשלום כרטיס אשראי`, F13), parsers (F3), verify, make_summary |
| `cycle_day` | int 1..28 | `10` | `10` (cycle file) / `2` (per-card statement) | parse_card_cycle (cycle date = cycle_day of month MM−1) |
| `fx_settlement` | object or null | `null` | `{"currency": "USD", "transfer_pattern": "העברה לחשבון מט\"ח"}` | see below |
| `fx_settlement.currency` | ISO code | — | `"USD"` | parse_card_cycle (F4 settlement currency), fx_rates, make_summary |
| `fx_settlement.transfer_pattern` | substring | absent | `"העברה לחשבון מט\"ח"` | classify: bank rows matching it → `העברה פנימית` / `המרת מט"ח` |
| `reconstruct_missing_cycles` | bool | `false` | `false` | parse_card_blocks (synthetic row for a missing cycle; also `--reconstruct-missing-cycles`) |

## `benefit_programs[]` — clubs (parse_benefits / classify)

| Key | Type | Default | Example | Read by |
|---|---|---|---|---|
| `id` | unique string | — required | `"club"` | normalized file name |
| `issuer` | `club_docx` | — required | `"club"` | parse_benefits |
| `label` | string | = id | `"מועדון הטבות"` | labels |
| `owner` | person id or `"shared"` | `"shared"` | `"p2"` | classify |
| `files` | glob(s) | `null` | `"inputs/benefits/club/*.docx"` | parse_benefits |
| `settles_from` | an `accounts[].id` | — required | `"bank_b"` | reconciliation (F7) |
| `bank_debit_pattern` | substring | `null` → format default `חיוב מועדון הטבות` | `"חיוב מועדון הטבות"` | parse_benefits (actual debit date), classify (→ `תשלום כרטיס אשראי`), make_summary |
| `default_debit_day` | int | `2` | `2` | parse_benefits (charge date when no bank debit is found) |
| `record_marker` | string | `null` | `"<club label>"` | parse_benefits (paragraph 0 of every benefit record is the club's own label; `null` = any non-empty paragraph followed by a date starts a record) |

## `p2p[]` — BIT / PayBox entries (classify / add_transfers_sheet / build_dashboard)

| Key | Type | Default | Example | Read by |
|---|---|---|---|---|
| `id` | unique string | — required | `"bit_p1"` | — |
| `app` | string | — required | `"BIT"` | source label `"<app> <owner label>"` |
| `owner` | person id | — required | `"p1"` | classify (unmatched card rows → that person's category) |
| `label` | string | `"<app> <owner label>"` | `"BIT בן/בת זוג א'"` | labels |
| `csv` | path | `null` | `"inputs/p2p/bit_p1.csv"` | classify (app rows, skipped at overview) |
| `screenshots` | glob | `null` | `"inputs/p2p/bit_p1/*.png"` | inventory only (Claude transcribes them) |
| `card_marker` | substring in card rows | `"BIT"` | `"BIT"` | classify (F6 pairing), add_transfers_sheet, make_summary, build_dashboard |
| `match_days` | int ≥ 0 | `5` | `5` | classify (F6) |
| `match_tolerance` | number ≥ 0 | `0.01` | `0.01` | classify (F6) |
| `bank_withdrawal_pattern` | substring | absent | `"Withdrawal from my balance"` | classify: bank rows matching it → internal `משיכת P2P לבנק` |

## `categories`

| Key | Type | Default | Example | Read by |
|---|---|---|---|---|
| `categories.scheme` | `"default"` or a CSV path (`group,cat,fixed`) | `"default"` (= `references/category-scheme-default.csv`) | `"rules/scheme.csv"` | common.scheme_path, classify (canonicalisation), build_excel |
| `categories.from_template` | object or null | `null` | see below | classify, build_excel — the template is opened read-only and ONLY its category list is read; no sheet of it is copied into the output workbook |
| `from_template.file` | path to the template workbook | — required when set | `"inputs/template/תזרים.xlsx"` | scheme read (never edited, never copied) |
| `from_template.sheet` | sheet name | — required | `"הכנסות - הוצאות"` | scheme read |
| `from_template.groups_col` / `cats_col` | column letters | — required | `"A"` / `"B"` | scheme read (contiguous rows, group repeats downwards) |
| `from_template.first_row` / `last_row` | ints | — required | `2` / `7` | scheme read |

Legacy keys `from_template.link_avg_col` and `from_template.manual_cells` (v0.1.0: linking the
template's average column into the copied template tab) are accepted and ignored since v0.1.1 with a
one-line warning on stderr — remove them from your config.
| `categories.secondary_scheme` | CSV path or null | `null` | `"rules/scheme_proposed.csv"` | classify (`group_new` / `cat_new`), build_excel (`קטגוריות מוצעות`, deep) |

## `rules` — file paths

| Key | Default | Columns | Read by |
|---|---|---|---|
| `rules.merchant_rules` | `rules/merchant_rules.csv` | `pattern,name_clean,type,group,cat[,group2,cat2],note`; `#` comments; longest pattern wins, ties in file order; no commas in any field | classify |
| `rules.user_labels` | `rules/user_labels.csv` | `id,user_text,user_cat,timestamp` | apply_user_labels (writes), classify (reads), add_transfers_sheet |
| `rules.user_labels_interpreted` | `rules/user_labels_interpreted.csv` | `id,type,group_tz,cat_tz,group_new,cat_new,trip,name_clean,note` (Claude writes it at deep level) | classify, add_transfers_sheet |
| `rules.fx_overrides` | `rules/fx_overrides.csv` | `source_file,row_ref,currency,reason` | parse_card_cycle |

## `estimates[]` — optional synthetic cash rows (classify; off at overview)

| Key | Type | Default | Example |
|---|---|---|---|
| `id` / `name` | string | — | `"cash_market"` / `"שוק – מזומן"` |
| `amount` | number ₪ | `0` | `250` |
| `every_n_months` | int ≥ 1 | `1` | `1` |
| `day` | day of month | `15` | `15` |
| `group` / `cat` | scheme names | `שונות` / `אחר / לא מזוהה` | `"מזון"` / `"ירקות ופירות"` |
| `note` | string | `""` | `"הערכת המשתמש"` |
| `funded_by_cat` | category whose expense rows are retyped internal (no double count) | absent | `"מזומן"` |

Rows get `source = "ידני – הערכה"`, `source_file = "מחושב"`, name `"<name> – הערכה"`.

## `trips[]` — optional trip attribution (classify; off at overview)

| Key | Type | Example |
|---|---|---|
| `label` | string (required) | `"חופשה לדוגמה"` |
| `countries` | list of country words (Hebrew/English as they appear in statements) | `["פורטוגל", "Portugal"]` |
| `start` / `end` | ISO dates, end ≥ start (required) | `"2025-02-03"` / `"2025-02-10"` |

A row joins a trip by country word in name/details, by the trip's currency, or by date range;
vacation rows matching no trip get `trip = "לא משויך – לסיווג ידני"` and a `נדרש:` note.

## `fx`, `verify`, `dashboard`, `thresholds`

| Key | Type | Default | Read by |
|---|---|---|---|
| `fx.rates_dir` | path | `work/boi_rates` | fx_rates (cache) |
| `fx.source` | `boi_sdmx` \| `manual` | `boi_sdmx` | fx_rates (`manual` never calls the API) |
| `fx.manual_rates` | CSV path (`date,ccy,rate`) or null | `null` | fx_rates (fallback; rows flagged `שער ידני`) |
| `verify.excel_recalc` | `auto` \| `always` \| `never` | `auto` (= run when macOS + Excel) | verify (`--excel-recalc` overrides) |
| `dashboard.chartjs` | `cdn` \| `inline` | `cdn` | build_dashboard (`inline` embeds a copy cached under `work/chartjs/`) |
| `thresholds.highlight_expense_avg` | number ₪ | `3000` | build_excel (bold red averages ≥ value in `הוצאות`; a third of it in the variable-merchants sheet) |
| `thresholds.highlight_income_avg` | number ₪ | `12000` | build_excel (`הכנסות`) |
| `thresholds.dynamic_merchant_rows` | int | `250` | build_excel (`פילוח` merchant table height) |
| `thresholds.dynamic_txn_rows` | int | `500` | build_excel (`פילוח` transaction table height) |
| `thresholds.top_merchants` | int | `25` | build_excel, make_summary, make_figures |
| `thresholds.hyperlink_rows` | int | `45` | build_excel (tall hyperlink ranges so a block opens at the top) |

## `work`, `notes`, `outputs` — paths

| Key | Default |
|---|---|
| `work.dir` / `work.normalized` / `work.database` / `work.summary` / `work.findings` / `work.figures` | `work` / `work/normalized` / `work/database.csv` / `work/summary.json` / `work/findings.json` / `work/figures` |
| `notes.dir` / `notes.decisions` / `notes.questions_and_answers` | `notes` / `notes/decisions.md` / `notes/questions_and_answers.md` |
| `outputs.dir` / `outputs.workbook` / `outputs.dashboard` / `outputs.report_html` / `outputs.report_pdf` | `outputs` / `outputs/תזרים.xlsx` / `outputs/dashboard.html` / `work/report.html` / `outputs/דוח.pdf` |

`common.ensure_dirs()` creates `work/`, `work/normalized`, `work/figures`, `fx.rates_dir`,
`outputs/`, `notes/` and the `rules/` folder.

## Validation rules (exit 2)

Missing file / invalid JSON; missing `window.start|end`, `household.people[].id|label`; invalid
level; `end < start`; `month_rule ≠ charge_date`; duplicate or reserved (`shared`) ids across
accounts / cards / programs / p2p / people; unknown `owner`; unknown `bank` / `issuer` id;
`settles_from` not an account id; `cycle_day` outside 1..28; negative `match_days` /
`match_tolerance`; bad enums in `verify`, `dashboard`, `fx`; non-numeric highlight thresholds;
incomplete `from_template`; a trip whose `end` precedes `start`.

## Minimal config (one account, one card, two months)

```json
{
  "window": {"start": "2025-01-01", "end": "2025-02-28"},
  "household": {"people": [{"id": "p1", "label": "בן/בת זוג א'"}]},
  "accounts": [{"id": "bank_a", "bank": "bank_xlsx_a", "files": "inputs/bank/bank_a/*.xlsx"}],
  "cards": [{"id": "card_1234", "issuer": "card_cycle_xlsx", "last4": "1234", "owner": "p1",
             "files": "inputs/cards/card_cycle/card_*.xlsx", "settles_from": "bank_a",
             "bank_debit_pattern": "חיוב לכרטיס ויזה 1234"}]
}
```
