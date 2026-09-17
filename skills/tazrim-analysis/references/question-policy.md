# Question policy — when to ask, and how

Loads before any AskUserQuestion. The rules come from a standing user instruction: questions must
be detailed and self-contained, and transaction-level questions do not belong in chat at all.
Every example below is synthetic.

## 1. What may be asked in chat (AskUserQuestion)

Methodology and structure only — one round, at most 8 questions, on a first run; a short round
again only when a new structural decision appears (a new transfer class, a new account, a scheme
change). The allowed topics:

| Topic | Typical options (recommended first) |
|---|---|
| window bounds | first day of month A .. last day of month B; extend / shrink |
| month rule | bank charge month (recommended, only supported value); transaction month |
| analysis level | overview / standard / deep |
| template linking | link the template's average column by references, keep listed manual cells / do not link |
| treatment of a **class** of transfers (one question per class) | internal transfer / expense / savings |
| missing or duplicate cycle file | re-download (recommended) / reconstruct from the bank debit (synthetic, flagged) |
| statistics over empty months | count as 0 with "ממוצע ללא אפס" alongside (recommended) / ignore empty months |
| category scheme approval | approve the shown list / approve with the shown additions / edit |
| card → person mapping | confirm the proposed mapping / correct it |
| other accounts out of scope | treat inflows from them as internal (recommended) / as income |

Everything else — "what is this merchant", "is this transfer rent", "which category for this
row" — is **never** asked in chat.

## 2. What goes to the "לסיווג ידני" sheet instead

Any transaction-level doubt: an unknown merchant, an ambiguous transfer, a cheque, a P2P payment
with no note, a vacation row with no country, a row whose classification you inferred but cannot
prove. Mechanism: set the row's `rule_note` to `נדרש: <the exact decision needed>` (through a
rule note, `user_labels_interpreted.csv`, or a category ending in "לא מזוהה"); `build_excel.py`
then places the full database row in the sheet with a red `נדרש:` column and the two user columns.
The user answers in Excel, at their pace, and `apply_user_labels.py` feeds the answers back.

Good `נדרש:` notes name the decision, not the row (the row's own columns are beside it):
`נדרש: למי נכתב השיק ועל מה` · `נדרש: הוצאה או העברה פנימית` · `נדרש: מדינה/נסיעה` ·
`נדרש: בחירת קטגוריה מהרשימה`.

## 3. Form of a chat question

Each question must be fully self-contained **inside the `question` field**:

1. Name the exact source: folder/file name (as saved under `inputs/`), the account or card label,
   the sheet row if known.
2. Give the date(s), the amount(s) in ₪, and the description text exactly as it appears in the
   source (in quotes).
3. Say whether a figure is a **single transaction** or a **sum of several** (and how many).
4. One item per question — one transfer class, one file, one decision. Use several questions in
   one call rather than bundling.
5. Offer only real options, the recommended one first, each with a one-line consequence. Do not add
   an option like "I'll write in free text" — the built-in *Other* is the free-text path; a custom
   option with that label returns no text.
6. Never refer to "the list above", "the table in my message" or your own numbering (1/2/3,
   א/ב/ג) — the user cannot map them. A summary table in chat before the round is fine, but each
   question repeats the detail it needs.
7. When the question is about a file, give the full path relative to the project (or open the file
   for the user when they ask for it).
8. Write the question in the user's language (Hebrew for a Hebrew-speaking household); keep the
   options short.

Record every answer in `notes/questions_and_answers.md` and every resulting decision in
`notes/decisions.md` (date · decision · reason · what it touches).

## 4. Examples (synthetic)

**Good — a transfer class:**

> Question: "In `inputs/bank/bank_a/statement.xlsx` (account עו"ש חשבון א') there are 2 rows
> described exactly as `הו"ק לחשבון אחר` dated 2025-01-10 and 2025-02-10, 1,000 ₪ each (sum of the
> two: 2,000 ₪). Where does this money go? Choose how the analysis should count both rows."
> Options: "Internal transfer to one of our own accounts (recommended if the account is yours) —
> not counted as an expense" · "Expense (e.g. rent, a payment to a person) — counted in the
> expense total" · "Savings/investment — shown in the non-summed block".

**Good — a missing cycle:**

> Question: "The bank shows a debit `חיוב לכרטיס ויזה 1234` of 3,000 ₪ on 2025-02-10 (a single
> transaction), but there is no file `inputs/cards/card_cycle/card_0325.xlsx` for that cycle. How should the
> February cycle be handled?" Options: "I will download the file (recommended; the analysis waits)"
> · "Reconstruct one synthetic row from the bank debit, flagged as reconstructed".

**Bad — bundled, referential, transaction-level:**

> "For items 1–3 in the table above, tell me which are expenses. Also, what is the merchant on
> row 5? Or write your answer in free text (option 3)."

Why it is bad: refers to a table and numbering the user cannot see in the prompt; three decisions
in one question; asks about a single merchant in chat (that row belongs in "לסיווג ידני"); adds a
fake free-text option.

**Bad — vague:**

> "Is the transfer in January internal?"

Why: no file, no account, no amount, no description text, no single-vs-sum statement, no options
with consequences.

## 5. After the round

Do not re-ask a question that was answered; read `notes/decisions.md` first. If an answer
contradicts the data later (e.g. a "transfer" that is clearly a merchant), put the specific rows in
"לסיווג ידני" with a `נדרש:` note that cites the earlier decision — do not reopen the chat question.
