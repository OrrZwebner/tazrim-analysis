# Analysis method — rules and formulas

Loads during classification, whenever the reconciliation residual is not 0, and before writing
`work/findings.json`. Every definition here is the one the workbook formulas, `make_summary.py`
and `verify.py` share; a figure that cannot be traced to one of these definitions does not go
into the report. All numbers in the examples are synthetic.

## 1. Month, window and flags (F16)

For a row with bank charge date $d$ and window $[s, e]$:

$$\text{month}(d) = d[0{:}7] \qquad \text{in\_window} = [\,s \le d \le e\,] \qquad
\text{summed} = \text{in\_window} \land \text{type} \in \{\text{הוצאה}, \text{הכנסה}\}$$

- **Month = bank charge month**, not transaction month: card files map one-to-one onto bank debits,
  so no cycle is split; the accepted cost is that a purchase made early in a month is counted in the
  next cycle's month. The transaction date stays in `txn_date` for the transaction lists.
- **Window end** is the last day of the last month. The first day of the following month would add
  an (N+1)-th rent / standing order to an N-month window. Rows outside the window stay in the
  database with `in_window = לא` so the next-cycle file (needed for FX rows) does not pollute totals.
- Flags are recomputed as the **last** step of classification, after every row-appending step.

## 2. Row types and the no-double-count table

`amount` in the database is positive for expenses **and** incomes; `type` carries the direction;
refunds stay negative inside their type (F20). Only `הוצאה` / `הכנסה` are summed.

| The same money appears as … | Counted once as | Other appearance |
|---|---|---|
| card statement rows + the bank's card debit (F13) | the statement rows (`הוצאה`) | bank row → `תשלום כרטיס אשראי`, `summed = לא` |
| an uncharged standing order (`--` charge date) that reappears charged in the next file (F14) | the charged row | uncharged row dropped; if it never reappears it gets the next cycle date and a flag |
| the same statement row in two files (F15) | the earlier file's row | later file's row dropped (same key twice in **one** file is kept — two real purchases) |
| a P2P payment funded by card: app row + card row (F6) | the app row on the payment date | card row → `כפילות`, `linked_id` = app row |
| club load + club purchases + bank debit (F7) | purchases and benefits at face value (`הוצאה`) + one computed negative discount row | load → `העברה פנימית`; bank debit → `תשלום כרטיס אשראי` |
| transfer to the household's own FX account + card FX rows (F4) | the card's FX rows, converted | transfer → `העברה פנימית`; only the bank's exchange fee is an expense |
| cash withdrawal + estimated cash expenses (`estimates[]`, `funded_by_cat`) | the estimate rows | withdrawal rows of that category → `העברה פנימית` |
| reimbursable expense + the later reimbursement | neither is summed | `הוצאה בהחזר` / `החזר הוצאה`, open balance shown (F21) |

## 3. Transfer vs expense (the rule, verbatim)

A bank transfer / cheque is **not** automatically internal: rent standing order, kindergarten
cheques, advisors are expenses; internal only when the money demonstrably stays in the family's
own accounts — pension / gemel deposits, securities purchases, transfers between the household's
own banks, FX conversions, P2P withdrawals to the bank, club loads, cash withdrawals that fund
estimate rows. Inflows from the household's other accounts that are out of scope are internal
transfers, not income, and the methodology section names those accounts.

Direction of a transfer in the transfers sheet (F17):
$$\text{נכנס} \iff (\text{type} = \text{הכנסה}) \oplus (\text{amount} < 0)$$
amounts are shown as absolute values.

## 4. FX settlement and fees (F4, F5)

A card row of type `חו"ל` / `זיכוי-חו"ל` whose charge date $d$ differs from the cycle date is
settled from the card's FX account in currency $c$ (`fx_settlement.currency`):

$$\text{amount\_ils} = \text{charge\_amount}_c \times r_c(d), \qquad
r_c(d) = \text{BOI representative rate on the last published day} \le d \text{ (at most 10 days back)}$$

The row is flagged `המרה משוערת`; per cycle, $\sum$ FX rows must equal the file header's `$` / `€`
totals. The transfer that funds the FX account is internal (§3).

Cross-currency fee when the statement shows `סכום ביניים a CCY … חיוב במט"י b CCY` (a purchase in a
third currency settled in the account currency):

$$\text{fee\_ils} = \frac{b - a}{b}\,\text{amount\_ils}, \qquad
\text{fee\%} = \frac{\sum \text{fee\_ils}}{\sum \text{amount\_ils}}, \qquad
\text{bank fee\%} = \frac{\sum \text{עמלת חליפין}}{\sum \text{FX transfers}}$$

A card linked to an FX account pays **no** conversion fee on purchases in that account's currency;
never state a fee percentage the data does not show.

Toy check: a purchase of 100 USD charged on a day with $r = 3.50$ → 350 ₪. Compute every fee from
the statement's own intermediate amounts; a difference that is only exchange-rate movement is not a fee.

## 5. Statistics over N months (F8, F25)

Let $x_1..x_N$ be a category's monthly totals over the window months; a month without a charge is
$x_i = 0$ (the budget cost of that category per month must reflect empty months).

| Statistic | Excel | Definition |
|---|---|---|
| total | `SUM` | $\sum_i x_i$ |
| average | `AVERAGE` | $\bar{x} = \frac{1}{N}\sum_i x_i$ |
| average without zero | `IFERROR(AVERAGEIF(rng,"<>0"),0)` | $\frac{\sum_{x_i \ne 0} x_i}{|\{i : x_i \ne 0\}|}$, 0 when no month has a value |
| max / min month | `MAX` + `INDEX(months, MATCH(MAX, rng, 0))` (also a cell comment) | $\arg\max_i x_i$ |
| median | `MEDIAN` | — |
| standard deviation | `_xlfn.STDEV.S` inside `IFERROR(…,0)` | $s = \sqrt{\frac{1}{N-1}\sum_i (x_i-\bar{x})^2}$, 0 when $N<2$ |
| months with charge | `COUNTIF(rng,"<>0")` | $|\{i : x_i \ne 0\}|$ |

**Every** average is shown next to its no-zero twin — stat sheets, analysis tables, pivot KPIs,
per-category blocks, dashboard KPIs, and the "no zero" twin of every average chart (sorted by its
own metric). Toy: $x = (250, 0, 500)$ → total 750, average 250, no-zero average 375, max month 3,
$s = \sqrt{(0^2 + (-250)^2 + 250^2)/2} = 250$.

Monthly helper (F9): one row per (scheme, type, group, category):
`=SUMIFS(DBamt, DBcat, $cat, DBtype, $type, DBmonth, month, flag, "כן")` with `flag = DBsummed`
for expenses/incomes and `DBin` for the non-summed types. Total rows (F10) are
`SUMIFS(DBamt, DBtype, "הוצאה", DBmonth, m, DBsummed, "כן")` — never a column sum — and a check
row recomputes $\sum$ category rows $-$ $\sum$ subtotal rows and prints ✓ when within 1 ₪.
Group subtotals (F11) are `+` over the member rows' month cells, statistics over the sums.

Sorting: groups by total descending, sub-categories by average descending; zero-total categories
are dropped. Fixed vs variable (F22): fixed = the configured / scheme-flagged categories,
variable = total − fixed. Israel vs abroad (F23): `orig_currency == ILS` vs not.

## 6. Cash reconciliation identity (F12)

Over the window, the change of the bank balances must be explained by the classified rows. The
identity `make_summary.py` evaluates (`summary.json → recon`), every term in ₪ and defined on
types / flags / config — never on names:

$$
\begin{aligned}
\Delta\text{bank} =\;& \text{income} - \text{expenses} - \text{savings\_bank} - \text{internal\_bank\_net}
 - \text{reimb\_paid\_bank} + \text{reimb\_recv\_bank} \\
&+ \text{fx\_expenses} - \text{p2p\_income} + \text{p2p\_paid\_from\_balance}
 - \text{card\_duplicates\_unlinked} - \text{card\_nonexpense\_rows} \\
&+ \text{club\_balance\_change} + \text{nonstatement\_expenses} - \text{nonbank\_income} + \text{residual}
\end{aligned}
$$

| Term | Meaning (in-window rows) |
|---|---|
| $\Delta$bank | net change of the accounts' balances (from the statements' balance columns when available, else $-\sum$ signed bank rows); `bank_change_source` says which |
| income / expenses | $\sum$ summed `הכנסה` / `הוצאה` rows |
| savings_bank | `חיסכון והשקעות` rows in the bank (money that left but is still the household's) |
| internal_bank_net | signed net of `העברה פנימית` bank rows (FX transfers, own-account moves, club loads, P2P withdrawals …) |
| reimb_paid_bank / reimb_recv_bank | `הוצאה בהחזר` paid from the bank / `החזר הוצאה` received in the bank |
| fx_expenses | card expenses settled from the FX account — they are in `expenses` but did not leave the ILS account, so they are added back |
| p2p_income | incoming P2P counted as income that stayed in the app balance |
| p2p_paid_from_balance | P2P expenses paid from the app balance or funded outside the window |
| card_duplicates_unlinked | `כפילות` card rows without a linked app row |
| card_nonexpense_rows | ILS card rows that are neither expense nor duplicate (e.g. internal) |
| club_balance_change | $\sum$ per program of (face-value net − bank debit) |
| nonstatement_expenses | estimate rows and other rows from no statement (`source` not an account/card/program/app) |
| nonbank_income | income rows not received in a bank account |
| residual | must satisfy $|\text{residual}| < 1$; `--strict` exits 4 otherwise |

The card-statement-vs-bank-debit difference (F3) is deliberately **not** on the explained side —
it surfaces in the residual and is reported per card next to it. After every reclassification
round, re-run the identity; when it drifts, the component whose type definition changed is the
culprit (sign of refunds, P2P funded outside the window, FX rows retyped).

Toy: income 9,000; expenses 7,000 of which 700 FX-settled; savings 1,000; internal net +500 out.
Explained $= 9{,}000 - 7{,}000 - 1{,}000 - 500 + 700 = 1{,}200$; if the bank balances rose by
1,200 the residual is 0.

## 7. Optional features (each gated by config, default off)

- **Reimbursables** (F21): types `הוצאה בהחזר` / `החזר הוצאה`, excluded from all averages; open
  balance per month $= \sum \text{הוצאה בהחזר} - \sum \text{החזר הוצאה}$ (positive = still owed).
- **Estimates** (`estimates[]`): synthetic expense rows `source = "ידני – הערכה"` every
  `every_n_months` on `day`; the funding cash withdrawals (`funded_by_cat`) become internal.
- **Trips** (`trips[]`): a vacation row joins the trip whose country word appears in its name /
  details, whose currency matches, or whose date range contains the charge date; unassigned
  vacation rows get `trip = "לא משויך – לסיווג ידני"` and a `נדרש: מדינה/נסיעה` note. The
  per-category sheet renders one block per trip with its own total row.
- **Benefit programs** (F7): discount row per debit month, computed, `source_file = מחושב`.
- **Missing cycle reconstruction**: off; a synthetic row from the bank debit only when the user chose it.

## 8. Classification order (P10) — why the order matters

1. Generated rules from config (card debits, FX transfer, P2P withdrawal) then `merchant_rules.csv`
   (case-insensitive substring, longest pattern wins, ties in file order); unmatched bank inflows →
   income "אחר / לא מזוהה", everything else → expense "אחר / לא מזוהה", `rule_note = לא נמצא כלל סיווג`.
2. Unmatched foreign-currency rows → vacation category, country from details.
3. Club loads → internal.
4. P2P pairing (F6): $|\Delta \text{amount}| \le 0.01$, $|\Delta \text{days}| \le 5$, nearest date,
   each app row once; unmatched card rows of a member without app data → "<person> – העברה לא מזוהה (BIT/PayBox)".
5. Canonicalise category names to the scheme (spaces / quotes / commas stripped before matching).
6. Computed discount rows.
7. `user_labels.csv` (dropdown category and/or free text → `name_clean = original – text`,
   `rule_note = סיווג ידני של המשתמש: …`), then `user_labels_interpreted.csv` (explicit per-id overrides).
8. Estimates and trips (skipped at overview).
9. Helper columns and flags — last.

At standard/deep, rows still unknown after step 7 carry `נדרש:` in `rule_note` so they land in
"לסיווג ידני"; at overview they are lumped silently.

## 9. Merchant research protocol (deep level)

Batch the unidentified merchant strings; search Hebrew and English, adding the city, "בע"מ", or
"כרטיס אשראי חיוב"; record `what it is / proposed category / confidence high-medium-low / URL` in
`notes/merchant_research.md`; write "לא נמצא" rather than guess; flag prepaid loads and vouchers
(double-count risk); remember a merchant may be registered under its street address — a user hint
plus one search resolves it. Turn each confirmed result into a specific rule placed **before** any
generic pattern it overlaps.
