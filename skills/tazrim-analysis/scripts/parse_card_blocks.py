#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""parse_card_blocks.py — per-card monthly statements (3-block .xls or detail .xlsx) -> work/normalized/<card-id>.csv (P5, F3/F24).

Purpose : for every `cards[]` entry with issuer "card_blocks_xls" / "card_detail_xlsx": glob `<last4>_MMYY.*` files, parse
          (a) the legacy `.xls` 3-block layout (sheet Activities: `כרטיס:NNNN ... חודש החיוב:
          dd/mm/yyyy` -> `עסקאות בשקלים` / `עסקאות במט"ח` sub-blocks, each with a `תאריך עסקה`
          header row and a `סה"כ` row; the same layout in .xlsx is accepted) and (b) the new
          `.xlsx` layout (sheet `פירוט עסקאות`: header row located by `תאריך רכישה`, charge date
          from `לחיוב ב-dd.mm` + the year cell, card from the NNNN cell, stop at `סה"כ`).
          Credits in the legacy layout carry a positive original amount with a negative charge
          -> both stored negative (F24). Every block's parsed sum is checked against its own
          `סה"כ` (self-check) and reconciled against the settling account's bank debit
          `bank_debit_pattern` on the charge date (F3, reported). Byte-identical files and files
          in `--ignore` are skipped. A cycle with a bank debit but no statement file is REPORTED,
          never invented, unless `reconstruct_missing_cycles` (config per card or the CLI flag)
          is on - then a SYNTHETIC row with a note is emitted.
Inputs  : tazrim.config.json (`cards[]` issuer card_blocks_xls: id, last4, label, owner, files,
          settles_from, bank_debit_pattern, reconstruct_missing_cycles?);
          work/normalized/<settles_from>.csv (from parse_bank.py) for the reconciliation.
Outputs : work/normalized/<card-id>.csv per card; one JSON object on stdout with per-file
          blocks, self-checks, reconciliation and missing cycles.
Exit    : 0 ok; 1 parse failure; 2 config error.
"""
import datetime as _dt
import os
import re
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import (NORMALIZED_COLUMNS, emit, ensure_dirs, fail, log, parse_args,  # noqa: E402
                    project_path, resolve_files, stable_id, write_csv)
from formats import ISSUERS  # noqa: E402
from parse_bank import load_grid, md5, num, s, to_date  # noqa: E402
from parse_card_cycle import bank_debits  # noqa: E402

SPEC = ISSUERS["card_blocks_xls"]
LEG = SPEC["blocks_xls"]
NEW = SPEC["detail_xlsx"]
CCY = SPEC["currency_labels"]
SOURCE_LABEL = "כרטיס אשראי (דף חיוב)"
TOL = 0.005


def currency(v):
    t = s(v)
    if t in CCY:
        return CCY[t]
    if re.match(r"^[A-Z]{3}$", t):
        return t
    raise ValueError("unknown currency label %r" % t)


def ddmmyyyy_in(text):
    m = re.search(LEG["date_pattern_in_markers"], text)
    if not m:
        raise ValueError("no dd/mm/yyyy date in %r" % text)
    return _dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1))).isoformat()


def join_details(*parts):
    return " | ".join(p for p in (s(x).replace("\n", " ") for x in parts) if p)


def is_legacy_layout(rows):
    tc = LEG["text_column"]
    return any(len(r) > tc and s(r[tc]).startswith(LEG["block_card_marker"]) for r in rows[:60])


# ----------------------------------------------------------------------------- legacy blocks
def parse_legacy(rows, datemode, fname, card_of):
    """Block layout -> (records, blocks). records carry last4; blocks = [{charge_date, last4,
    kind, total, parsed_sum, ok}]."""
    tc = LEG["text_column"]
    recs, blocks = [], []
    last4 = charge_date = kind = None
    header = None
    for idx, v in enumerate(rows):
        v = list(v) + [None] * (8 - len(v))
        b = s(v[tc])
        if b.startswith(LEG["block_card_marker"]):
            last4 = re.search(LEG["block_card_pattern"], b).group(1)
            charge_date = ddmmyyyy_in(b)
            header = kind = None
            continue
        if b.startswith(LEG["block_ils_marker"]) or b.startswith(LEG["block_fx_marker"]):
            kind = "ILS" if b.startswith(LEG["block_ils_marker"]) else "FX"
            try:
                bd = ddmmyyyy_in(b)
                if bd != charge_date:
                    raise ValueError("%s r%d: sub-block charge date %s != block %s" % (fname, idx + 1, bd, charge_date))
            except ValueError as e:
                if "no dd/mm/yyyy" not in str(e):
                    raise
            header = None
            continue
        if b == LEG["table_header_marker"]:
            header = [s(x) for x in v]
            expected = LEG["header_ils"] if kind == "ILS" else LEG["header_fx"]
            got = header[tc:tc + len(expected)]
            if got != expected:
                raise ValueError("%s r%d: %s block header mismatch.\n  expected: %s\n  got:      %s" % (
                    fname, idx + 1, kind, expected, got))
            continue
        if b == LEG["total_marker"]:
            tot = num(v[tc + 3]) if kind == "ILS" else num(v[tc + 4])
            blocks.append({"charge_date": charge_date, "last4": last4, "kind": kind, "total": round(tot or 0.0, 2), "row": idx + 1})
            header = None
            continue
        if header is None or last4 is None or b == "":
            continue
        d = to_date(v[tc], datemode)
        if d is None:
            continue
        if kind == "ILS":
            name, orig_amt, chg_amt, memo = v[tc + 1], num(v[tc + 2]), num(v[tc + 3]), v[tc + 4]
            orig_ccy = "ILS"
        else:
            name, orig_amt, orig_ccy, chg_amt, chg_ccy, memo = (v[tc + 1], num(v[tc + 2]), currency(v[tc + 3]),
                                                                num(v[tc + 4]), v[tc + 5], v[tc + 6])
            if currency(chg_ccy) != "ILS":
                raise ValueError("%s r%d: charge currency is not ILS" % (fname, idx + 1))
        if chg_amt is None:
            continue
        chg_amt = round(chg_amt, 2)
        orig_amt = round(orig_amt if orig_amt is not None else chg_amt, 2)
        if chg_amt < 0 < orig_amt:
            orig_amt = -orig_amt                                   # F24
        recs.append({
            "last4": last4, "original_name": s(name), "txn_date": d.isoformat(), "charge_date": charge_date,
            "amount_ils": chg_amt, "orig_currency": orig_ccy, "orig_amount": orig_amt,
            "details": join_details(memo, LEG["block_ils_marker"] if kind == "ILS" else LEG["block_fx_marker"]),
            "source_file": fname, "row_ref": idx + 1, "kind": kind,
        })
    for blk in blocks:
        got = round(sum(r["amount_ils"] for r in recs if r["charge_date"] == blk["charge_date"]
                        and r["last4"] == blk["last4"] and r["kind"] == blk["kind"]), 2)
        blk["parsed_sum"], blk["ok"] = got, abs(got - blk["total"]) < TOL
    return recs, blocks


# ----------------------------------------------------------------------------- new layout
def parse_new(rows, fname, known_last4):
    cells = []
    for r in rows[:15]:
        for c in list(r)[:8]:
            t = s(c)
            if t:
                cells.append(t)
    charge_txt = next((t for t in cells if t.startswith(NEW["charge_cell_prefix"])), None)
    if not charge_txt:
        raise ValueError("%s: no '%s' cell in the first 15 rows" % (fname, NEW["charge_cell_prefix"]))
    year_txt = next((t for t in cells if re.search(NEW["year_cell_pattern"], t) and len(t) < 20), None)
    if not year_txt:
        raise ValueError("%s: no cell with the year (20yy) in the first 15 rows" % fname)
    year = int(re.search(r"(20\d\d)", year_txt).group(1))
    m = re.search(NEW["charge_cell_pattern"], charge_txt)
    charge_date = _dt.date(year, int(m.group(2)), int(m.group(1))).isoformat()
    last4 = None
    for t in cells:
        for m4 in re.findall(NEW["card_cell_pattern"], t):
            if m4 in known_last4:
                last4 = m4
                break
        if last4:
            break
    if last4 is None:
        fm = re.match(r"^(\d{4})_", fname)
        last4 = fm.group(1) if fm else None
    hrow = next((i for i, r in enumerate(rows) if r and s(r[0]) == NEW["header_marker"]), None)
    if hrow is None:
        raise ValueError("%s: header row ('%s') not found" % (fname, NEW["header_marker"]))
    got = [s(x) for x in list(rows[hrow])[:len(NEW["header"])]]
    if got != NEW["header"]:
        raise ValueError("%s: header mismatch.\n  expected: %s\n  got:      %s" % (fname, NEW["header"], got))
    recs, total = [], None
    for idx in range(hrow + 1, len(rows)):
        v = list(rows[idx]) + [None] * 8
        a = s(v[0])
        if a.startswith(NEW["total_marker"]):
            total = num(v[1])
            break
        if not a:
            continue
        d = to_date(a)
        if d is None:
            continue
        recs.append({
            "last4": last4, "original_name": s(v[1]), "txn_date": d.isoformat(), "charge_date": charge_date,
            "amount_ils": round(num(v[4]) or 0.0, 2), "orig_currency": currency(v[5]) if s(v[5]) else "ILS",
            "orig_amount": round(num(v[2]) if num(v[2]) is not None else (num(v[4]) or 0.0), 2),
            "details": join_details(v[7], ("מס' שובר: " + s(v[6])) if s(v[6]) else "", "עסקאות למועד חיוב"),
            "source_file": fname, "row_ref": idx + 1, "kind": "ILS",
        })
    got_sum = round(sum(r["amount_ils"] for r in recs), 2)
    blocks = [{"charge_date": charge_date, "last4": last4, "kind": "ILS", "total": round(total, 2) if total is not None else None,
               "parsed_sum": got_sum, "ok": (total is None) or abs(got_sum - total) < TOL, "row": None}]
    return recs, blocks


# ----------------------------------------------------------------------------- driver
def parse_card_blocks(cfg, cards, ignore, reconstruct_cli):
    by_last4 = {str(c["last4"]): c for c in cards if c.get("last4")}
    files, seen = [], set()
    for c in cards:
        for p in resolve_files(cfg, c.get("files")):
            if p not in seen and os.path.splitext(p)[1].lower() in SPEC["extensions"]:
                seen.add(p)
                files.append(p)
    files.sort(key=os.path.basename)
    if not files:
        raise ValueError("no files matched cards[].files for the card_blocks cards")
    digests, skipped, per_file, all_recs, all_blocks = {}, [], {}, [], []
    for path in files:
        fname = os.path.basename(path)
        if fname in ignore:
            skipped.append("%s: ignored (--ignore)" % fname)
            continue
        d = md5(path)
        if d in digests:
            skipped.append("%s is byte-identical to %s - skipped; re-download the missing cycle" % (fname, digests[d]))
            continue
        digests[d] = fname
        reader = "xlrd" if path.lower().endswith(".xls") else "openpyxl"
        rows, datemode = load_grid(path, LEG["sheet"] if reader == "xlrd" else NEW["sheet"], reader)
        if is_legacy_layout(rows):
            recs, blocks = parse_legacy(rows, datemode, fname, by_last4)
            layout = "legacy"
        else:
            recs, blocks = parse_new(rows, fname, set(by_last4))
            layout = "new"
        for r in recs:
            c = by_last4.get(r["last4"])
            r["card_id"] = c["id"] if c else None
            r["card_label"] = c["label"] if c else "%s %s" % (SOURCE_LABEL, r["last4"])
        per_file[fname] = {"layout": layout, "rows": len(recs), "blocks": blocks}
        all_recs += recs
        all_blocks += [dict(b, file=fname) for b in blocks]
        log("[card_blocks] %s (%s): %d rows, %d blocks, self-check %s" % (
            fname, layout, len(recs), len(blocks), "OK" if all(b["ok"] for b in blocks) else "FAIL"))

    # reconciliation per (charge_date, card) vs bank debit (F3)
    totals = defaultdict(float)
    for b in all_blocks:
        totals[(b["charge_date"], b["last4"])] += b["total"] or b["parsed_sum"]
    reconcile, missing, synthetic = [], [], []
    for c in cards:
        bd = bank_debits(cfg, c["settles_from"], c.get("bank_debit_pattern"))
        keys = {k for k in totals if k[1] == str(c.get("last4"))}
        bank_dates = set(bd) if bd else set()
        for cd in sorted({k[0] for k in keys} | bank_dates):
            ft = round(totals.get((cd, str(c.get("last4"))), 0.0), 2) if (cd, str(c.get("last4"))) in totals else None
            bank = bd.get(cd, (None, []))[0] if bd is not None else None
            entry = {"card": c["label"], "card_id": c["id"], "charge_date": cd, "statement": ft, "bank": bank,
                     "diff": round(ft - bank, 2) if (ft is not None and bank is not None) else None,
                     "status": "ok" if ft is not None and bank is not None else (
                         "bank csv missing" if bd is None else ("no statement file" if ft is None else "no bank debit"))}
            reconcile.append(entry)
            if ft is None and bank is not None:
                do_it = reconstruct_cli or bool(c.get("reconstruct_missing_cycles", False))
                missing.append({"card": c["label"], "charge_date": cd, "bank_debit": bank,
                                "action": "SYNTHETIC row emitted" if do_it else "reported only - download the statement"})
                if do_it:
                    ids = bd[cd][1]
                    all_recs.append({
                        "last4": str(c.get("last4")), "card_id": c["id"], "card_label": c["label"],
                        "original_name": "%s - חיוב ללא פירוט (משוחזר מחיוב הבנק)" % c["label"],
                        "txn_date": cd, "charge_date": cd, "amount_ils": round(bank, 2), "orig_currency": "ILS",
                        "orig_amount": round(bank, 2),
                        "details": "SYNTHETIC | שוחזר מחיוב הבנק %s (%s) - אין קובץ פירוט למחזור זה" % (cd, ",".join(ids)),
                        "source_file": "SYNTHETIC", "row_ref": cd, "kind": "SYNTHETIC",
                    })
                    synthetic.append({"card": c["label"], "charge_date": cd, "amount": bank})
    report = {"files": [os.path.basename(f) for f in files], "skipped": skipped, "per_file": per_file,
              "self_check_ok": all(b["ok"] for b in all_blocks), "blocks": all_blocks,
              "reconcile": reconcile, "missing_cycles": missing, "synthetic_rows": synthetic,
              "unconfigured_cards": sorted({r["last4"] for r in all_recs if not r["card_id"]}),
              "rows": len(all_recs), "total_amount_ils": round(sum(r["amount_ils"] for r in all_recs), 2)}
    return all_recs, report


def to_rows(card_id, recs):
    out = []
    for r in sorted(recs, key=lambda r: (r["charge_date"], r["source_file"], str(r["row_ref"]))):
        out.append({
            "id": stable_id(card_id, r["source_file"], r["row_ref"]), "source": SOURCE_LABEL, "card": r["card_label"],
            "original_name": r["original_name"], "txn_date": r["txn_date"], "charge_date": r["charge_date"],
            "amount_ils": "%.2f" % r["amount_ils"], "orig_currency": r["orig_currency"],
            "orig_amount": "%.2f" % r["orig_amount"], "details": r["details"],
            "source_file": r["source_file"], "row_ref": r["row_ref"],
        })
    return out


def main(argv=None):
    def extra(p):
        p.add_argument("--ignore", nargs="*", default=[], help="file basenames to skip (e.g. a known duplicate)")
        p.add_argument("--reconstruct-missing-cycles", action="store_true",
                       help="emit a SYNTHETIC row from the bank debit for a cycle without a statement file")
    args, cfg = parse_args("parse per-card statements into work/normalized/<card-id>.csv", extra, argv=argv)
    cards = [c for c in cfg["cards"] if c["issuer"] == "card_blocks_xls"]
    if not cards:
        fail("no cards[] with issuer 'card_blocks_xls' configured", "add the cards to tazrim.config.json")
    ensure_dirs(cfg)
    try:
        recs, report = parse_card_blocks(cfg, cards, set(args.ignore), args.reconstruct_missing_cycles)
    except Exception as e:  # noqa: BLE001
        fail("card_blocks: %s" % e, "check the file layout against references/bank-formats.md (formats.py 'card_blocks_xls')")
        return
    out_dir = project_path(cfg, "work.normalized")
    outputs = {}
    for c in cards:
        rows = to_rows(c["id"], [r for r in recs if r["card_id"] == c["id"]])
        path = os.path.join(out_dir, "%s.csv" % c["id"])
        write_csv(path, rows, NORMALIZED_COLUMNS)
        outputs[c["id"]] = {"csv": path, "rows": len(rows), "label": c["label"]}
    stray = [r for r in recs if not r["card_id"]]
    if stray:
        path = os.path.join(out_dir, "card_blocks_unassigned.csv")
        write_csv(path, to_rows("card_blocks_unassigned", stray), NORMALIZED_COLUMNS)
        outputs["card_blocks_unassigned"] = {"csv": path, "rows": len(stray), "note": "card numbers not in cards[]"}
    warnings = list(cfg.warnings) + report["skipped"]
    if not report["self_check_ok"]:
        warnings.append("block self-check failed - see report.blocks")
    for rr in report["reconcile"]:
        if rr["diff"] not in (None, 0.0):
            warnings.append("reconcile %s %s: statement %.2f vs bank %.2f (diff %.2f)" % (
                rr["card"], rr["charge_date"], rr["statement"], rr["bank"], rr["diff"]))
    for mc in report["missing_cycles"]:
        warnings.append("missing statement for %s on %s (bank debit %.2f): %s" % (mc["card"], mc["charge_date"], mc["bank_debit"], mc["action"]))
    for w in warnings:
        log("WARNING: " + w)
    emit({"ok": True, "cards": outputs, "report": report, "warnings": warnings})


if __name__ == "__main__":
    main()
