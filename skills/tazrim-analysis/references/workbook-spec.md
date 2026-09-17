# Workbook specification — `outputs/תזרים.xlsx`

Loads when building or debugging the workbook (`scripts/build_excel.py`,
`scripts/add_transfers_sheet.py`, `scripts/verify.py`). The workbook is always created from
scratch (`openpyxl.Workbook()`); when `categories.from_template` is set the template workbook is
opened read-only and ONLY its category list (groups column / categories column / row range) is
read — no template sheet is copied into the output (since v0.1.1). Every aggregate is an Excel formula over the
`database` sheet; the only literal numbers are the database rows themselves. Every sheet is RTL
(`sheet_view.rightToLeft = True`). All values in the examples are synthetic.

## Sheet set and order

| Level | Sheets built |
|---|---|
| overview | `database`, `עזר_חודשי`, `הוצאות`, `הכנסות`, `ניתוח נתונים` |
| standard | + `פילוח`, `פירוט לפי קטגוריה`, `פירוט עסקאות`, `הוצאות משתנות לפי בית עסק`, `לסיווג ידני`, `העברות, BIT ו-PAYBOX` (post-build), `רשימות` (hidden) |
| deep | + `קטגוריות מוצעות` |

Final order (`common.SHEET_ORDER`): `הוצאות` (always the first tab), `הכנסות`, `ניתוח נתונים`,
`העברות, BIT ו-PAYBOX` (inserted by `add_transfers_sheet.py` right after `ניתוח נתונים`), `פילוח`,
`פירוט לפי קטגוריה`, `פירוט עסקאות`, `הוצאות משתנות לפי בית עסק`, `קטגוריות מוצעות`, `לסיווג ידני`,
`database`, `עזר_חודשי`, `רשימות`; active sheet = `הוצאות`.

## Before building: what is carried forward

`read_previous()` opens the previous output workbook (read-only) and keeps, from `לסיווג ידני`,
`{id: (user_text, user_cat)}` for every row where either user column is filled — the header row is
found by its first cell `מזהה`. These are re-applied to the new workbook (`preserved_labels` in the
JSON). Nothing else is carried forward (the budget columns of v0.1.0 were removed in v0.1.1). The
user's original template file is never written.

## 1. `database`

All `database.csv` rows, Hebrew headers (`common.DB_HEADERS_HE`), an Excel Table `DB`
(`TableStyleMedium2`), freeze panes `A2`, amounts `#,##0.00`. **Named ranges** over each column,
rows 2..N+1, used by every SUMIFS in the workbook:

| Name | Column | Name | Column | Name | Column |
|---|---|---|---|---|---|
| `DBid` | id | `DBname` | name_clean | `DBmonth` | month |
| `DBsource` | source | `DBtype` | type | `DBmname` | month_name |
| `DBcard` | card | `DBgroup` | group_tz | `DBamt` | amount |
| `DBpay` | pay | `DBcat` | cat_tz | `DBcur` | orig_currency |
| `DBperson` | person | `DBgroup2` | group_new | `DBin` | in_window |
| `DBorig` | original_name | `DBcat2` | cat_new | `DBsummed` | summed |
| `DBdate` | txn_date | `DBcharge` | charge_date | `DBnote` | rule_note |
| `DBdet` | details | `DBlink` | linked_id | `DBtrip` | trip |

Sheet references in formulas are always quoted: `'database'!$Q$2:$Q$501`.

## 2. `עזר_חודשי` (helper)

Row 1 a grey note; row 2 header `סכימה, סוג, קבוצה, קטגוריה, <N month labels>, סה"כ`; freeze `E3`.
One row per (scheme, type, group, category), then total rows; month cell (F9):

```
=SUMIFS(DBamt,DBcat,$D<r>,DBtype,$B<r>,DBmonth,<col>$2,DBsummed,"כן")   (expense / income rows)
=SUMIFS(DBamt,DBcat,$D<r>,DBtype,$B<r>,DBmonth,<col>$2,DBin,"כן")       (non-summed types)
```

Total rows use `DBtype` + `DBsummed` only (no category). Stat sheets reference these cells
(`='עזר_חודשי'!E<r>`), never the database directly, except the total rows (F10).

## 3. Statistic sheets — `הוצאות`, `הכנסות`, `קטגוריות מוצעות`

Column layout (`Layout(n)`, identical for the three sheets):

| Column index | Content |
|---|---|
| 1 | `קבוצה` (merged downwards per group, light blue) |
| 2 | `תת-קטגוריה` (hyperlink to its block in `פירוט לפי קטגוריה` at standard/deep) |
| 3 .. 2+N | one column per month, label `<Hebrew month> <year>`, heat colour scale |
| 3+N | `סה"כ N חודשים` = `SUM` (green scale) |
| 4+N | `ממוצע חודשי` = `AVERAGE` (blue scale; bold red when ≥ `thresholds.highlight_*_avg`) |
| 5+N | `ממוצע ללא אפס` = `IFERROR(AVERAGEIF(rng,"<>0"),0)` (blue scale) |
| 6+N .. 13+N | `מקסימום`, `חודש מקס` (`INDEX($C$2:$<m1>$2,MATCH(max,rng,0))`), `מינימום`, `חודש מין`, `חציון`, `סטיית תקן` (`IFERROR(_xlfn.STDEV.S(rng),0)`), `חודשים עם חיוב` (`COUNTIF(rng,"<>0")`), `הערות` |

Row 1 = title merged across columns 1..5+N plus a legend cell; row 2 = header; freeze `C3`; data
from row 3. Offsets derive from N (`Layout`), never hard-coded — verify reads the header row to
find columns. MAX / MIN cells also carry a `Comment` with the month and value.

`הוצאות`: rows = (group, category) pairs with a non-zero in-window total, groups by total desc,
categories by average desc; a **group subtotal** row (`סה"כ <group>`, orange fill) after each group
whose month cells are `=C5+C6+…` over the member rows; then the total row with the **exact
label** `סה"כ הוצאות שוטפות (נסכם ישירות מ-database)` (yellow, `SUMIFS` on the helper total
row); the next row is the check
`בדיקה: סכום שורות הקטגוריות (ללא שורות סה"כ קבוצה)` = `SUM(totals) − Σ subtotal cells` and
`=IF(ABS(check−total)<1,"✓ תואם לסה""כ","✗ פער: "&TEXT(…))` in the notes column. Below: a grey
non-summed block (`חיסכון, השקעות, העברות פנימיות וחיובי כרטיסי אשראי`) with one row per (type,
category) sorted by total, card-debit rows noted as informational; then, only when reimbursables
exist, their block and an open-balance row. The "אחר / לא מזוהה" row's note points at `לסיווג ידני`.

`הכנסות`: same layout; groups = income groups (per person when the scheme
has them); total row label exactly `סה"כ הכנסות (נסכם ישירות מ-database)`; then
`סה"כ הוצאות שוטפות (מלשונית הוצאות)` (references), `מאזן חודשי (הכנסות − הוצאות שוטפות)` (bold),
`החזרי הוצאות שהתקבלו – לא נסכם` and `חיסכון והשקעות בפועל – לא נסכם` (grey, `SUMIFS` on `DBin`).

`קטגוריות מוצעות` (deep): the secondary scheme on `DBcat2`, same statistics, a total row that must
equal the primary total, plus a "by group" table with `% מסה"כ`.

## 4. `רשימות` (hidden)

One named list per column (`L_tz_g`, `L_tz_c`, `L_new_g`, `L_new_c`, `L_months` (= `כל החודשים` +
labels), `L_pay` (= `הכל` + payment methods), `L_person`, `L_scheme`, `L_level`, `L_allcats`), each a
defined name over `'רשימות'!$<col>$2:$<col>$k`; later the hyperlink anchor arrays `A_keys`,
`A_rows`, `A_first_keys`, `A_first_rows` (`<category>||<merchant>` → first row of its block in
`פירוט עסקאות`).

## 5. `פילוח` (interactive, Excel 365 dynamic arrays)

Fixed layout: title B1, usage line B2, six green input cells `C4:C9` (סכימה, רמה, בחירה, חודש,
אמצעי תשלום, אדם) with list validations — the selection list is
`=INDIRECT(IF($C$4="תזרים",IF($C$5="קבוצה","L_tz_g","L_tz_c"),IF($C$5="קבוצה","L_new_g","L_new_c")))`;
KPIs in `F4:F8` (total, count, average per transaction, monthly average ÷ N when all months are
selected, % of all expenses in those months). The condition is one product of boolean arrays:

```
((DBtype="הוצאה")*(DBsummed="כן")*IF($C$4="תזרים",IF($C$5="קבוצה",DBgroup=$C$6,DBcat=$C$6),
 IF($C$5="קבוצה",DBgroup2=$C$6,DBcat2=$C$6))*(($C$7="כל החודשים")+(DBmname=$C$7))
 *(($C$8="הכל")+(DBpay=$C$8))*(($C$9="הכל")+(DBperson=$C$9)))
```

- KPI cells: `SUMPRODUCT(cond*DBamt)` written as a **single-cell `ArrayFormula`** (F19) — as a
  plain formula implicit intersection returns 0.
- Merchant table (F18) at `B<T0>:F<T0+249>` (`thresholds.dynamic_merchant_rows`): one
  `ArrayFormula` over the fixed range —
  `IFERROR(INDEX(_xlfn.LET(_xlpm.c,cond,_xlpm.fn,_xlfn._xlws.FILTER(DBname,_xlpm.c>0),_xlpm.fa,…,
  _xlpm.u,_xlfn.UNIQUE(_xlpm.fn),_xlpm.m,--(TRANSPOSE(_xlpm.u)=_xlpm.fn),_xlpm.s,MMULT(TRANSPOSE(_xlpm.m),_xlpm.fa),…,
  _xlfn.SORTBY(CHOOSE({1,2,3,4,5},u,s,n,s/n,s/t),s,-1)),_xlfn.SEQUENCE(250),{1,2,3,4,5}),"")`.
- Transaction list at `H<T0>:M<T0+499>` (`dynamic_txn_rows`) — `FILTER` over
  `CHOOSE({1..6},DBdate,DBname,DBamt,DBcat,DBpay,DBdet)` sorted by amount.
- Column G `↗`: `HYPERLINK("#'פירוט עסקאות'!A"&INDEX(A_rows,MATCH($C$6&"||"&B<r>,A_keys,0))&":N"&(…+45),"↗")`
  with a fallback on `A_first_*` when the merchant sits in another category.
- The chart is anchored to the right of the tables so it never covers them.

## 6. `פירוט לפי קטגוריה`, `פירוט עסקאות`, `הוצאות משתנות לפי בית עסק`

- `פירוט לפי קטגוריה`: an index of hyperlinks (4 per row), then one block per (group, category):
  a light-blue header row `<group> / <category>` with `סה"כ:`, a header
  `בית עסק (שם מובן), סה"כ N חודשים, ממוצע חודשי, ממוצע ללא אפס, מס' עסקאות, ממוצע לעסקה,
  % מהקטגוריה, <months>`, merchant rows as `SUMIFS` / `COUNTIFS` on `DBname` + `DBcat`, sorted
  desc; the vacation category is split into **trip blocks** (text-only header, merchant rows with
  an extra `DBtrip` criterion, one `סה"כ <trip>` row, medium top border); every merchant name links
  to its block in `פירוט עסקאות`; a `↩` link returns to `הוצאות`. Block width = max(14, 7+N).
- `פירוט עסקאות`: for every (category, merchant) a header row `<group> / <category> / <merchant>`,
  a `↩ לפירוט הקטגוריה` link, the columns `תאריך עסקה, תאריך חיוב, שם כפי שמופיע בקובץ, סכום ₪,
  אמצעי תשלום, אדם, שם חודש, פרטים, הערת סיווג, מזהה` where **each cell is `='database'!<col><row>`**,
  then a `סה"כ` row = `SUM`. These blocks are the hyperlink targets.
- `הוצאות משתנות לפי בית עסק`: every merchant outside the fixed categories, `#, בית עסק, תת-קטגוריה,
  קבוצה, סה"כ N חודשים, ממוצע חודשי, ממוצע ללא אפס, …`, sorted desc, autofilter, freeze `C4`, bold
  red names when the average ≥ `highlight_expense_avg / 3`, a `סה"כ הוצאות משתנות` row.

**Hyperlink rule**: a link never targets a single cell — it targets the tall range
`A<r>:<width><r+45>` (`thresholds.hyperlink_rows`) so Excel scrolls the block to the **top** of the window.

## 7. `ניתוח נתונים`

Sorted tables written top-down with a running `state["row"]`; each chart records
`minrow = header_row + 2.1·height + 2`, and the next block starts at `max(row, minrow)`, so charts
never overlap tables (charts sit at columns max(13, 8+N) and +11). Blocks: א monthly
expenses/income/balance/savings (bar + line); ב by group (avg and no-zero twin, sorted each by its
own metric, pie); ג by payment method; ג2 per-card pie; ג3 by person; ד income by source with three
pies (average / no-zero / total); income by person per month; top-25 merchants (avg and no-zero);
fixed vs variable; Israel vs abroad; benefits club face vs bank vs discount; י secondary-scheme
groups by month (deep). Per-row averages are `=total/N` and `IFERROR(AVERAGEIF(...,"<>0"),0)`.

## 8. `לסיווג ידני`

Rows 1–2 instructions (merged), row 3 header = all database headers + `נדרש:` + `למילוי משתמש
(טקסט חופשי)` + `קטגוריה (בחירה מהרשימה)`; an Excel Table `ManualTable` (`TableStyleLight9`),
**no freeze panes**; the user columns at the far end, green fill, the category cell validated
against `L_allcats`. Row selection (`manual_mask`): in-window rows of type expense / income /
reimbursable whose category is empty, "אחר / לא מזוהה", or contains `לא מזוהה` / `לא ידוע` /
`לסיווג` / `לא מפורט`, **or** whose `rule_note` contains `לסיווג ידני` / `לבדיקה` / `לאימות` /
`נדרש:` / `לבדוק`; sorted by amount desc. The `נדרש:` column is red and repeats the note's
`נדרש: …` text, else `נדרש: בחירת קטגוריה מהרשימה (לא נמצא כלל סיווג)`, else
`נדרש: אימות הסיווג — <note>`. Previously saved user text / category is re-filled.

## 9. `העברות, BIT ו-PAYBOX` (added by `add_transfers_sheet.py`, openpyxl only)

Dropped and rebuilt right after `ניתוח נתונים`. Title and note merged across the 18 columns:
`מזהה, סוג תנועה, כיוון, תאריך עסקה, חודש, סכום ₪, מוטב / צד שני, שם מובן, קטגוריה (תזרים),
קבוצה, אמצעי תשלום / חשבון, אדם, בחלון, נסכם, הערת סיווג, הערה מהמשתמש / מהצילום, פרטים מהמקור,
מזהה מקושר`. Three sections — 1 bank transfers (generic description patterns: `העברה`, `הו"ק`,
`משיכת שיק`, `הפקדת שיק`, `כספומט`, `קופת גמל`, `ניירות ערך`, `Withdrawal` …, excluding salary
and card-debit rows), 2 P2P (app rows, card rows funding them, withdrawals, duplicates), 3 PayBox —
each sorted by amount desc with an amount colour scale, unknown payees in **red**, and a subtotal
row of **formulas**: `=COUNTA(ids)`, `=SUMIFS(amounts, directions, "יוצא")`, `=SUMIFS(…, "נכנס")`.
Payee text comes from the row itself, the user's label files and the p2p config — no built-in dictionary.

## 10. Formula conventions (failure modes)

- Decide the full row order **before** writing; never `insert_rows` into a written sheet
  (openpyxl does not shift formula references).
- Excel-365 functions need prefixes or every cell shows `#NAME?`: `_xlfn.STDEV.S`, `_xlfn.LET`,
  `_xlfn.UNIQUE`, `_xlfn.SORTBY`, `_xlfn.SEQUENCE`, `_xlfn._xlws.FILTER`; LET variables `_xlpm.<name>`.
- `SUMPRODUCT` / `LET` / `FILTER` results must be `ArrayFormula` objects (single cell for KPIs,
  fixed tall range for tables) — plain formulas fall to implicit intersection (blank or 0).
- Every string literal inside a formula doubles its quotes (`q()`); an unescaped `"` in a category
  name makes Excel refuse the file **silently** while alerts are off — bisect by stripping features.
- Category names must not contain commas (they break the rules CSV and criteria strings).
- Values that look like formulas (`=…` text from statements) are written as text (`data_type = "s"`).
- Long RTL titles are merged across the table width, otherwise the first column clips them.
- Colour scales are full-cell white→colour with dark text; no data bars.

## 11. What `verify.py` checks per sheet

| Scope | Check |
|---|---|
| all sheets | no `#REF!` / `#NAME?` / `#…` in formulas or cached values; the level's sheet set is present |
| `database` | every named range referenced by a formula exists; every summed row has a primary category; potential duplicates (same source/date/amount/name) listed — report only |
| `עזר_חודשי` | every SUMIFS row references a category present in the data |
| `הוצאות` / `הכנסות` | total row found by its **exact** label and referencing helper/database; category rows point at helper rows of the same category; every summed category has a row; group subtotal rows present |
| Excel recalculation (macOS + Excel, `excel_recalc auto\|always`) | temp copy `work/verify_calc.xlsx` recalculated and saved **in place** with `display alerts` off; no formula errors; each total and each category × month cell equals pandas (tol 0.01); the rows-check cell equals the total; transfers-sheet subtotals equal pandas; `פילוח` KPI for the default selection and for a **live** filter (first group / month / person in the data) equals pandas |
| otherwise | the Excel checks are reported `skipped` with the reason — never `pass` |
| cards | card statements vs bank debits in the window, FX explained — report only |

Exit 0 only when no check failed. Card-vs-bank and duplicate reports never fail the run.

## 12. Excel / AppleScript gotchas (macOS)

- `set display alerts to false` (and restore at the end) — otherwise a dialog blocks the run.
- Calculate and **save the copy in place**; a "save as" into another folder pops the sandbox
  "Grant File Access" dialog. The dismissal via System Events is best effort — never depend on
  System Events (assistive access is usually missing).
- Close temp workbooks `saving no` and delete `work/verify_tmp.xlsx` / `work/verify_calc.xlsx`
  before and after; a copy left open makes the next run fail with "did not save the calculated copy".
- `osascript` runs under `subprocess.run(timeout=)` (no `timeout` binary on macOS); poll with loops, not `sleep`.
- Check Excel's `version` ≥ 16 (365) — dynamic arrays and `LET` do not exist in older builds.
- The user's open workbook holds a lock file `~$תזרים.xlsx`; a rebuilt file "does not show the new
  sheets" until they close **without saving** and reopen.
