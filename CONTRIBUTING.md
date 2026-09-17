# Contributing

## Ground rules

- **No personal data, ever.** Every example, fixture, test value, merchant, name, card number,
  amount, date and path in this repository is synthetic. A pull request that carries a real
  statement row, a real household member, a real account or card number, or an absolute path
  from someone's machine is rejected. If you are unsure whether a value is personal, leave it out
  and write `SYNTHETIC-NEEDED` in a comment.
- Python 3.8 compatible (no `match`, no walrus-only tricks, `typing` generics), dependencies pinned
  in `requirements.txt`, one JSON object on stdout per script, diagnostics on stderr.
- Hebrew literals only where they name a sheet, header, category, row type or note.
- Contributions are accepted under the repository's [Personal Use License](LICENSE).

## Running the tests

```bash
pip install -r requirements.txt
python3 -m unittest discover -s skills/tazrim-analysis/scripts/tests -v
python3 skills/tazrim-analysis/scripts/make_fixtures.py --out /tmp/tazrim-fixture
python3 skills/tazrim-analysis/scripts/run_pipeline.py --config /tmp/tazrim-fixture/tazrim.config.json --mode A --level deep
```

The fixture generator writes a complete synthetic project (two accounts, cards, a club docx, a
P2P CSV, a template, a BOI rate cache) plus `expected.json` with **hand-derived** expected values;
tests compare parser and classifier output against those literals, never against values computed
from the same code. Add a hand-derived expectation with every new parsing rule. Expectations are
derived independently of the code and **must not be updated to match new output**: if a value
moves, the code is wrong until proven otherwise (re-derive by hand first).

CI (`.github/workflows/test.yml`) runs the same commands on Linux, so the Excel recalculation
checks are reported as `skipped` there — a test must not depend on Excel.

## The 500-line rule

`SKILL.md` is loaded on every invocation of the skill; the files under `references/` load only
when the procedure points at them. Keep `SKILL.md` under 500 lines (the validator warns above
that): new detail — a format quirk, a formula derivation, a workbook column — goes into the matching
reference file, and `SKILL.md` gets one line saying when to read it. Every file under `references/`
must be mentioned in `SKILL.md`, or it will never be loaded.

Frontmatter carries exactly six fields (`name`, `description`, `license`, `compatibility`,
`metadata`, `allowed-tools`); any other field breaks the claude.ai upload. Check with:

```bash
python3 <skill-factory>/skills/verify/scripts/validate_plugin.py .
```

## Adding or fixing an export format

`scripts/formats.py` is the single source of truth for sheet names, header rows, exact header
lists, block markers, date encodings and sign conventions. Formats are keyed by structure
(`bank_xlsx_a`, `card_cycle_xlsx`, ...), never by an institution's name — do not add one. When a
layout changes or a new one is needed:

1. Update the spec in `formats.py` (never patch the parser's assertions around it).
2. Extend `make_fixtures.py` so the synthetic file reproduces the new layout.
3. Update `references/bank-formats.md` to match.
4. Add a hand-derived expectation and run the tests.

## Volatile external authorities — verify before a release

| Authority | What to verify | Where it lives |
|---|---|---|
| Bank of Israel SDMX rates | the endpoint template still answers and the CSV still has `TIME_PERIOD` / `OBS_VALUE` | `formats.py` (`BOI`), `fx_rates.py`, `references/bank-formats.md` |
| Export layouts (`bank_xlsx_a`, `bank_xls_b`, `card_cycle_xlsx`, `card_blocks_xls` / `card_detail_xlsx`, `club_docx`) | header rows and lists against a fresh export | `formats.py`, `references/bank-formats.md` |
| Excel 365 function prefixes for openpyxl (`_xlfn.`, `_xlfn._xlws.`, `_xlpm.`) and `ArrayFormula` | a built workbook opens in the current Excel without `#NAME?` | `build_excel.py`, `references/workbook-spec.md` |
| openpyxl / xlrd / pandas / python-docx / matplotlib pins | the pinned versions install on a clean Python 3.8 | `requirements.txt` |
| Chart.js 4.4.1 on cdnjs | the URL still resolves; `dashboard.chartjs=inline` works offline | `build_dashboard.py` |
| Chrome / Chromium headless flags (`--headless=new`, `--no-pdf-header-footer`) | a PDF is produced by the current browser build | `render_pdf.sh` |
| Excel via AppleScript (`display alerts`, save in place, `close saving no`) | the verify gate runs on the current macOS + Excel | `verify.py`, `references/workbook-spec.md` §13 |

## Cutting a release

1. Bump `version` in `.claude-plugin/plugin.json`, `.claude-plugin/marketplace.json` and the `metadata` block of `SKILL.md`.
2. Run the tests and the validator; grep the tree for anything that looks personal
   (absolute home-directory paths, real names, card numbers) — the privacy scan must pass.
3. Tag `vX.Y.Z`; the release workflow builds `tazrim-analysis.zip` (the skill folder alone, with
   `SKILL.md` at its root) and attaches it to the GitHub release. Never commit build archives.
