# Export formats (bank, card, club)

Loads when parsing, when a header assertion fails, when transcribing P2P screenshots, or when
the Bank of Israel rate lookup misbehaves. Every fact here mirrors
`scripts/formats.py` — **that module is the single place to update** when a layout changes, and
the place to add a new institution's layout; the parsers only assert what it declares and fail
loudly on a mismatch. Formats are keyed by their **structure**, never by an institution's name:
match your own file to a section below by its sheet name, header row and header list. Run
`python3 scripts/formats.py <id>` for a summary, `python3 scripts/formats.py` for the full JSON.
Format ids: banks `bank_xlsx_a`, `bank_xls_b`; cards `card_cycle_xlsx`, `card_blocks_xls` /
`card_detail_xlsx`; club `club_docx`; rates `boi`.

The Hebrew strings below are the files' own headers and markers (format knowledge, not
personal data). Every file name, card number and amount in the examples is synthetic.

## Normalized output (all parsers)

Columns `id, source, card, original_name, txn_date, charge_date, amount_ils, orig_currency,
orig_amount, details, source_file, row_ref`. Sign: `amount_ils > 0` = money **out**, `< 0` = money
in (F20). `id = <PREFIX>-<6 hex>` hashed from (basename of `source_file`, `row_ref`) so ids survive
re-parses and file re-ordering. Source files are never modified.

## `bank_xlsx_a` — bank account export, format A (.xlsx, openpyxl; sheet `עובר ושב`, header row 8, newest first)

| Item | Value |
|---|---|
| Sheet | `עובר ושב` |
| Header row / data from | row 8 / row 9 (1-based) |
| Exact header | `תאריך`, `יום ערך`, `תיאור התנועה`, `₪ זכות/חובה`, `₪ יתרה`, `אסמכתה`, `עמלה`, `ערוץ ביצוע` |
| Date encoding | datetime cells |
| Order | newest first |
| Sign in file | + credit / − debit → `amount_ils = −amount` |
| Balance | every row → **F1** chain: `balance[i] − amount_file[i] == balance[i+1]` (tol 0.005); implied opening = last balance − last amount |
| Stop | first fully blank row or end of sheet |
| Quirks | resolve columns by header name, never by index; value date goes to `details` only when it differs from the date; reference / channel / fee / running balance are appended to `details` |

## `bank_xls_b` — bank account export, format B (.xls via xlrd, .xlsx via openpyxl; sheet `Activities`, header row 6, opening-balance row 7, oldest first)

| Item | Value |
|---|---|
| Sheet | `Activities` |
| Header row | row 6 (xlrd index 5) |
| Opening row | row 7: description `יתרת פתיחה`, value in `יתרה` — skipped as a transaction, kept as the opening balance (or `accounts[].opening_balance`) |
| Data from | row 8 |
| Exact header | `""`, `יתרה`, `תאריך ערך`, `זכות`, `חובה`, `תאור`, `אסמכתא`, `סוג פעולה`, `תאריך` |
| Date encoding | `.xls`: Excel serial (`xldate_as_datetime` with the workbook's datemode); `.xlsx`: datetime |
| Order | oldest first |
| Sign in file | separate `זכות` / `חובה` columns → `amount_ils = חובה − זכות` |
| Balance | sparse: filled only on the last row of a date group → **F2**: `running += זכות − חובה` from the opening balance, compared at every filled `יתרה` cell; reconstructed closing must equal the last sheet balance |
| Quirks | xlrd blank cells arrive as `' '`; card debits appear as `<last4> - <issuer>` (e.g. `1234 - כרטיס אשראי`) — configured per card in `bank_debit_pattern` |

## `card_cycle_xlsx` — credit-card cycle file, all cards in one file (.xlsx, openpyxl; sheet `כרטיסי אשראי`, totals in rows 7–8, header row 9, `--` = not yet charged)

| Item | Value |
|---|---|
| File name | `<prefix>_MMYY.xlsx`; **MMYY is the month AFTER the charge**; cycle date = `cycle_day` (default 10) of month MM−1. Never trust the name: confirm against row 7 and the bank debit |
| Sheet | `כרטיסי אשראי` |
| Row 6 | title `פירוט עסקאות - כל הכרטיסים הבנקאיים` |
| Rows 7–8 | cycle header strings: `N עסקאות לחיוב בחודש <month> בכל הכרטיסים` (or `… לחיוב הקודם`); `עסקאות באשראי - סה"כ חיוב:  + ₪<ils> + $<usd> + €<eur>` — recorded; they are the reconciliation targets |
| Header row / data from | row 9 / row 10 |
| Exact header (first 10 cells asserted) | `כרטיס`, `בית עסק`, `תאריך עסקה`, `סכום העסקה`, `מנפיק`, `סוג העסקה`, `פירוט`, `תאריך החיוב`, `סכום החיוב`, `כרטיס הוצג במעמד העסקה?` |
| Optional tail | `מטבע העסקה`, `שער ההמרה`, `תאריך שער`, `עמלת ההמרה`, `מדד בסיס`, `שם המועדון`, `אחוז הנחה`, `סכום הנחה` |
| Stop | first row whose column A starts with `הודעה` (or is empty) |
| Fixed columns (0-based) | card 0, merchant 1, txn_date 2, txn_amount 3, issuer 4, txn_type 5 |
| **Column shift** | from column 6 on, cells shift left when a field is empty: locate the charge date **by pattern** (`dd/mm/yyyy` or `--`) at column ≥ 6; the charge amount is the next numeric cell; `פירוט` is the text between column 6 and the charge date |
| Date encoding | `dd/mm/yyyy` strings |
| Card cell | `<card type> <last4>`, e.g. `ויזה 1234` |
| Transaction types | `ישראל` domestic, `חו"ל` foreign, `זיכוי-חו"ל` foreign credit, `תשלום-ישראל` / `תשלום` installments |
| Installments | `k מ - n` in `פירוט`: amount = the installment (סכום החיוב); full price kept in `orig_amount` and `details` |
| `--` rows | charge date `--` or empty = standing order not yet charged (**F14**: dropped iff the same (card, txn_date, amount) reappears charged in a later file, else assigned the next cycle date and flagged); txn amount `--` = not yet posted (newest file) → row skipped and listed |
| FX rule (**F4**) | a `חו"ל` / `זיכוי-חו"ל` row whose charge date ≠ cycle date is settled from the card's FX account in `fx_settlement.currency`: `amount_ils = charge_amount × BOI_rate(charge_date)`, flagged `המרה משוערת`; original currency from the `הומר ל-` text, from "charged 1:1", else from the country word, else `UNK` |
| Sign | + = charge; credits (זיכוי) negative |
| Duplicates (**F15**) | key (card, merchant, txn_date, charge_date, charge_amount) seen in two different files → drop the later; same key twice in one file → keep |
| Reconciliation (**F3**) | Σ charge amounts with charge_date == cycle date (non-FX) == the bank row matching `bank_debit_pattern` on the cycle date; every non-zero diff is printed and must be explained |
| Overrides | `rules/fx_overrides.csv` (`source_file,row_ref,currency,reason`) for a row typed `ישראל` but settled in FX |

## `card_blocks_xls` / `card_detail_xlsx` — per-card monthly statement (.xls: 3 blocks per card; .xlsx: sheet `פירוט עסקאות`)

File name `<last4>_MMYY.xls|.xlsx`; the charge date is read **from the file**, never from the
name. Default cycle day 2. Byte-identical files and `--ignore <basename>` files are skipped.

**Blocks `.xls`** (`blocks_xls`; xlrd, sheet `Activities`, markers in column B):
- block start `כרטיס:NNNN - <issuer label> חודש החיוב: dd/mm/yyyy` (card = the 4 digits);
- sub-blocks `עסקאות בשקלים חיוב בתאריך dd/mm/yyyy` and `עסקאות במט"ח חיוב בתאריך dd/mm/yyyy`;
- table header row starts with `תאריך עסקה`: ILS `תאריך עסקה, שם  העסק, סכום עסקה, סכום חיוב, פירוט`
  (note the double space in `שם  העסק`); FX `תאריך עסקה, שם  העסק, סכום מקורי, מטבע מקורי, סכום חיוב, מטבע חיוב, פירוט`;
- block total row `סה"כ` (ILS total in column E, FX total in column F); dates are Excel serials;
  blank cells are `' '`;
- **F24** credits: charge < 0 with a positive original amount and memo `זיכוי` → both stored negative;
- self-check: parsed sum per block == the block's `סה"כ` (tol 0.005).

**Detail `.xlsx`** (`detail_xlsx`; openpyxl, sheet `פירוט עסקאות`):
- scan rows 1..15 × columns 1..8 for the title cells: `לחיוב ב-dd.mm` gives the charge day/month,
  a short cell containing `20yy` gives the year, the cell containing 4 digits gives the card;
- header row located dynamically by its first cell `תאריך רכישה`: `תאריך רכישה, שם בית עסק, סכום
  עסקה, מטבע עסקה, סכום חיוב, מטבע חיוב, מס' שובר, פירוט נוסף`; dates `dd.mm.yy`;
- stop at `סה"כ` (e.g. `סה"כ לחיוב החודש בכרטיס בש"ח` in column A, value in B); amounts already
  signed (credits negative, no `זיכוי` label);
- the same merchant may be spelled differently across the two layouts — normalise on `name_clean`
  via rules, not on the raw string.

Currency labels → codes: `דולר ארה"ב` USD, `אירו` EUR, `לירה שטרלינג` GBP, `באט תאילנד` THB,
`ש"ח` / `ש''ח` / `₪` ILS. Reconciliation: Σ block == block `סה"כ` == bank debit `<last4> - <issuer>`
on the charge date. A cycle with a bank debit but no statement file, or a card without a block
(the issuer merged its fee onto the other card), is **reported, never fabricated**;
`reconstruct_missing_cycles` (config per card or `--reconstruct-missing-cycles`) emits a
synthetic row with a note when the user explicitly chooses that.

## `club_docx` — benefits-club export (.docx, python-docx; paragraph-only, two segments)

A paragraph-only document pasted from the club site; two segments, records are fixed runs of
paragraphs after a header sequence; trailing stray paragraphs and a truncated last record are ignored.

| Segment | Header sequence | Record | Shape |
|---|---|---|---|
| Benefits | `תאריך רכישה, שם ההטבה, שם המוצר, כמות, סה"כ, סטטוס` | 9 paragraphs, first = club label (`record_marker`, else any non-empty paragraph followed by a date) | club label, purchase date `dd.mm.yy`, product name (→ `original_name`), benefit name, quantity, total ₪, status, extra, extra |
| Loadable card | `תאריך ושעה, מספר פעולה, מזהה פיתקית, רשת, סניף, סוג פעולה, סכום` | 7 paragraphs, first matches `^\d{2}/\d{2}/\d{2}$` | date `dd/mm/yy`, operation number, voucher id, chain, branch, operation type, amount ₪ (signed) |

Operation types: `טעינה` (positive) → internal transfer, `details` starts `טעינה – לא נסכם`, name
`טעינת כסף - כרטיס נטען`; `חיוב` (negative) → expense at face value, name `<chain> - <branch>`;
benefits → expense at face value. `charge_date` = the settling account's actual debit
(`bank_debit_pattern`, default `חיוב מועדון הטבות`) in the month after the activity month, else
`default_debit_day` (default 2) of that month. **F7**: per debit month
`discount = (Σ loads + Σ benefits) − bank debit`; when |discount| > 0.005 a negative expense row
`הנחת מועדון (זיכוי מחושב)` with `source_file = מחושב` is emitted, so the expense total equals the cash out.

## `boi` — Bank of Israel representative rates (SDMX, volatile)

- Endpoint: `https://edge.boi.org.il/FusionEdgeServer/sdmx/v2/data/dataflow/BOI.STATISTICS/EXR/1.0/RER_<CCY>_ILS?startperiod=<window start>&endperiod=<today>&format=csv`
- Cache: `<fx.rates_dir>/RER_<CCY>_ILS.csv` (default `work/boi_rates/`), fetched once per currency
  when missing and `fx.source == "boi_sdmx"`; `parse_card_cycle.py --offline` never calls the API.
- CSV columns (SDMX header): `SERIES_CODE, FREQ, BASE_CURRENCY, COUNTER_CURRENCY, UNIT_MEASURE,
  DATA_TYPE, DATA_SOURCE, TIME_COLLECT, CONF_STATUS, PUB_WEBSITE, UNIT_MULT, COMMENTS,
  TIME_PERIOD, OBS_VALUE, RELEASE_STATUS`; the lookup uses `TIME_PERIOD` (YYYY-MM-DD) and
  `OBS_VALUE` (ILS per 1 unit of the currency).
- Lookup falls back to the last published day at most **10 days** back (weekends, holidays).
- Failure path: network errors never raise; the problem is reported in the JSON and the
  `fx.manual_rates` CSV (`date,ccy,rate`) is used, rows flagged `שער ידני`; with no rate at all the
  row keeps its foreign amount, `amount_ils` empty, and lands in "לסיווג ידני".
- **Verify the endpoint and column names before relying on them** — the API is outside this
  skill's control. `python3 scripts/fx_rates.py --config …` reports the cache state.

## P2P CSV (BIT / PayBox) — transcribed by Claude from screenshots

Same columns as the normalized output. `source = "<app> <person label>"`, `card = "<app>"`,
`original_name` = counterparty, `details = "<note> | <status> | יוצא/נכנס"`, `source_file` =
screenshot file name (or the CSV name), `row_ref` = running number, outgoing positive, incoming
negative, "Withdrawal to bank" rows kept (classified internal by `bank_withdrawal_pattern` /
the starter rule). Synthetic example row:

```
id,source,card,original_name,txn_date,charge_date,amount_ils,orig_currency,orig_amount,details,source_file,row_ref
BIT-a1b2c3,BIT בן/בת זוג א',BIT,חבר לדוגמה,2025-01-14,2025-01-14,250,ILS,250,מתנה משותפת | הושלם | יוצא,bit_p1_01.png,1
```

Card rows containing `card_marker` (default `BIT`) are paired with app outgoing rows by
`|Δamount| ≤ match_tolerance` (0.01) and `|Δdays| ≤ match_days` (5), nearest date wins, each app
row used once (**F6**).
