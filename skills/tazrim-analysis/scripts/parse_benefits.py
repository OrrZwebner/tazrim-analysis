#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""parse_benefits.py — benefits-club docx -> work/normalized/<program-id>.csv (P6, F7).

Purpose : for every `benefit_programs[]` entry read the paragraph-only .docx pasted from the
          club site: segment (a) benefits = 9 paragraphs per record after the header sequence
          `תאריך רכישה, שם ההטבה, שם המוצר, כמות, סה"כ, סטטוס`; segment (b) loadable card = 7
          paragraphs per record after `תאריך ושעה, מספר פעולה, מזהה פיתקית, רשת, סניף, סוג
          פעולה, סכום`. `טעינה` (load) -> internal transfer (details start with the load flag),
          `חיוב` (purchase from the loaded balance) and benefits -> expenses at face value.
          charge_date = the settling account's actual debit date (`bank_debit_pattern`) in the
          month after the activity month, else `default_debit_day` of that month. Per debit month
          the club discount is computed (F7): discount = (loads + benefits at face value) - bank
          debit, emitted as a NEGATIVE expense row "הנחת מועדון (זיכוי מחושב)" (source_file
          "מחושב") so the expense total equals the cash that left the bank.
Inputs  : tazrim.config.json (`benefit_programs[]`: id, issuer, label, owner, files,
          settles_from, bank_debit_pattern, default_debit_day, optional record_marker = the club's
          own label in paragraph 0 of every benefit record; without it a record starts at any
          non-empty paragraph followed by a date); work/normalized/<settles_from>.csv.
Outputs : work/normalized/<program-id>.csv; one JSON object on stdout with counts, totals,
          the monthly reconciliation and the discount rows.
Exit    : 0 ok; 1 parse failure (header sequence not found / python-docx missing); 2 config error.
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
from parse_card_cycle import bank_debits  # noqa: E402

SPEC = ISSUERS["club_docx"]
BEN, CARD = SPEC["benefits_segment"], SPEC["card_segment"]
AMOUNT_RE = re.compile(SPEC["amount_pattern"])
DISCOUNT_NAME = "הנחת מועדון (זיכוי מחושב)"
COMPUTED_FILE = "מחושב"


def parse_amount(text):
    m = AMOUNT_RE.search(str(text).replace("₪", ""))
    if not m:
        raise ValueError("no amount in %r" % text)
    return float(m.group(0).replace(",", ""))


def find_seq(ps, seq, start=0):
    n = len(seq)
    for i in range(start, len(ps) - n + 1):
        if ps[i:i + n] == seq:
            return i + n
    return None


def parse_date(text, fmts):
    for f in fmts:
        try:
            return _dt.datetime.strptime(text.strip(), f).date()
        except ValueError:
            continue
    raise ValueError("bad date %r" % text)


def next_month(y, m):
    return (y, m + 1) if m < 12 else (y + 1, 1)


def charge_date_for(txn, debits, default_day):
    """Activity month m is billed in m+1: the actual debit date when the bank shows one in that
    month, else default_day of m+1."""
    ny, nm = next_month(txn.year, txn.month)
    for d in sorted(debits or {}):
        if d[:7] == "%04d-%02d" % (ny, nm):
            return d
    return _dt.date(ny, nm, default_day).isoformat()


def is_date(text, fmts):
    try:
        parse_date(text, fmts)
        return True
    except ValueError:
        return False


def parse_docx(path, record_marker=None):
    try:
        from docx import Document
    except ImportError:
        raise RuntimeError("python-docx is not installed (pip install python-docx)")
    ps = [p.text.strip() for p in Document(path).paragraphs]
    recs, notes = [], []
    # ---- segment (a): benefits
    i = find_seq(ps, BEN["header_sequence"])
    n_ben = 0
    if i is None:
        notes.append("benefits header sequence not found - no benefit records")
    else:
        ben_fmts = ("%d.%m.%y", "%d/%m/%y", "%d.%m.%Y", "%d/%m/%Y")
        marker = record_marker or BEN["record_start_marker"]

        def starts_record(k):
            if marker:
                return ps[k] == marker
            return bool(ps[k]) and not is_date(ps[k], ben_fmts) and is_date(ps[k + 1], ben_fmts)

        while i + BEN["record_paragraphs"] <= len(ps) and starts_record(i):
            rec = ps[i:i + BEN["record_paragraphs"]]
            try:
                d = parse_date(rec[1], ben_fmts)
                amt = parse_amount(rec[5])
            except ValueError as e:
                notes.append("benefit record at paragraph %d skipped: %s" % (i, e))
                i += BEN["record_paragraphs"]
                continue
            recs.append({"kind": "benefit", "original_name": rec[2], "txn_date": d, "amount": amt,
                         "details": "הטבה: %s | כמות %s | %s" % (rec[3], rec[4], rec[6]), "row_ref": i})
            n_ben += 1
            i += BEN["record_paragraphs"]
    # ---- segment (b): loadable card
    j = find_seq(ps, CARD["header_sequence"])
    n_load = n_pur = 0
    date_re = re.compile(CARD["record_start_pattern"])
    if j is None:
        notes.append("card header sequence not found - no card records")
    else:
        while j + CARD["record_paragraphs"] - 1 < len(ps) and date_re.match(ps[j]):
            rec = ps[j:j + CARD["record_paragraphs"]]
            try:
                d = parse_date(rec[0], ("%d/%m/%y", "%d.%m.%y"))
                raw = parse_amount(rec[6])
            except ValueError as e:
                notes.append("card record at paragraph %d skipped: %s" % (j, e))
                j += CARD["record_paragraphs"]
                continue
            typ = rec[5]
            if typ == "טעינה":
                recs.append({"kind": "load", "original_name": CARD["load_name"], "txn_date": d, "amount": abs(raw),
                             "details": "%s | %s | %s | %s" % (CARD["load_flag"], rec[1], rec[2], typ), "row_ref": j})
                n_load += 1
            elif typ == "חיוב":
                name = " - ".join(x for x in (rec[3], rec[4]) if x)
                recs.append({"kind": "purchase", "original_name": name, "txn_date": d, "amount": abs(raw),
                             "details": "%s | %s | %s" % (rec[1], rec[2], typ), "row_ref": j})
                n_pur += 1
            else:
                notes.append("unknown operation type %r at paragraph %d - skipped" % (typ, j))
            j += CARD["record_paragraphs"]
        tail = [p for p in ps[j:] if p]
        if tail:
            notes.append("%d trailing non-empty paragraphs ignored after the card segment" % len(tail))
    return recs, {"benefits": n_ben, "loads": n_load, "purchases": n_pur}, notes


def parse_program(cfg, prog):
    files = resolve_files(cfg, prog.get("files"))
    if not files:
        raise ValueError("program %s: no files match %r" % (prog["id"], prog.get("files")))
    pattern = prog.get("bank_debit_pattern") or SPEC["bank_debit_pattern_default"]
    debits = bank_debits(cfg, prog["settles_from"], pattern)     # None when the bank csv is absent
    day = int(prog.get("default_debit_day") or 2)
    rows, counts, notes = [], defaultdict(int), []
    for path in files:
        fname = os.path.basename(path)
        recs, cnt, nts = parse_docx(path, prog.get("record_marker"))
        for k, v in cnt.items():
            counts[k] += v
        notes += ["%s: %s" % (fname, n) for n in nts]
        for r in recs:
            cd = charge_date_for(r["txn_date"], debits, day)
            rows.append({
                "id": stable_id(prog["id"], fname, r["row_ref"]), "source": prog["label"], "card": "",
                "original_name": r["original_name"], "txn_date": r["txn_date"].isoformat(), "charge_date": cd,
                "amount_ils": round(r["amount"], 2), "orig_currency": "ILS", "orig_amount": round(r["amount"], 2),
                "details": r["details"], "source_file": fname, "row_ref": r["row_ref"], "kind": r["kind"],
            })
    # ---- F7 per debit month
    by_month = defaultdict(lambda: defaultdict(float))
    for r in rows:
        by_month[r["charge_date"][:7]][r["kind"]] += r["amount_ils"]
    months = set(by_month) | {d[:7] for d in (debits or {})}
    reconcile, discount_rows = [], []
    for ym in sorted(months):
        bank = [(d, a) for d, (a, _ids) in (debits or {}).items() if d[:7] == ym]
        bank_amt = round(sum(a for _d, a in bank), 2)
        bank_ids = [i for d, (_a, ids) in (debits or {}).items() if d[:7] == ym for i in ids]
        loads, ben, pur = (round(by_month[ym][k], 2) for k in ("load", "benefit", "purchase"))
        face = round(loads + ben, 2)
        disc = round(face - bank_amt, 2) if bank else None
        reconcile.append({"debit_month": ym, "bank_debit_date": bank[0][0] if bank else None, "bank": bank_amt if bank else None,
                          "loads": loads, "benefits": ben, "face_value": face, "purchases_from_balance": pur,
                          "discount": disc, "status": "ok" if bank else ("bank csv missing" if debits is None else "no bank debit")})
        if disc is not None and abs(disc) > 0.005:
            cd = bank[0][0]
            discount_rows.append({
                "id": stable_id(prog["id"], COMPUTED_FILE, "D" + ym), "source": prog["label"], "card": "",
                "original_name": DISCOUNT_NAME, "txn_date": cd, "charge_date": cd, "amount_ils": round(-disc, 2),
                "orig_currency": "ILS", "orig_amount": round(disc, 2),
                "details": "שורה מחושבת (F7): טעינות %.2f + הטבות %.2f בערך נקוב פחות חיוב בבנק %.2f | מזהה חיוב בנק: %s" % (
                    loads, ben, bank_amt, ",".join(bank_ids)),
                "source_file": COMPUTED_FILE, "row_ref": "D" + ym, "kind": "discount",
            })
    all_rows = sorted(rows + discount_rows, key=lambda r: (r["txn_date"], str(r["row_ref"])))
    summary = {
        "label": prog["label"], "files": [os.path.basename(f) for f in files], "counts": dict(counts),
        "totals": {k: round(sum(r["amount_ils"] for r in rows if r["kind"] == k), 2) for k in ("benefit", "load", "purchase")},
        "reconcile": reconcile, "discount_rows": [{"month": r["row_ref"][1:], "amount_ils": r["amount_ils"], "id": r["id"]} for r in discount_rows],
        "bank_debits": debits, "bank_debit_pattern": pattern, "notes": notes, "rows": len(all_rows),
    }
    return all_rows, summary


def main(argv=None):
    args, cfg = parse_args("parse the benefits-club docx into work/normalized/<program-id>.csv", argv=argv)
    progs = [b for b in cfg["benefit_programs"] if b["issuer"] == "club_docx"]
    if not progs:
        fail("no benefit_programs[] with issuer 'club_docx' configured", "add the program to tazrim.config.json")
    ensure_dirs(cfg)
    out_dir = project_path(cfg, "work.normalized")
    result, warnings = {}, list(cfg.warnings)
    for prog in progs:
        try:
            rows, summary = parse_program(cfg, prog)
        except Exception as e:  # noqa: BLE001
            fail("program %s: %s" % (prog["id"], e), "check the docx against references/bank-formats.md (formats.py 'club_docx')")
            return
        path = os.path.join(out_dir, "%s.csv" % prog["id"])
        write_csv(path, rows, NORMALIZED_COLUMNS)
        summary["csv"] = path
        result[prog["id"]] = summary
        if summary["bank_debits"] is None:
            warnings.append("%s: bank csv for %s not found - run parse_bank.py first; charge dates defaulted, no discount rows" % (prog["id"], prog["settles_from"]))
        warnings += ["%s: %s" % (prog["id"], n) for n in summary["notes"]]
        log("[%s] benefits=%d loads=%d purchases=%d discount rows=%d" % (
            prog["id"], summary["counts"].get("benefits", 0), summary["counts"].get("loads", 0),
            summary["counts"].get("purchases", 0), len(summary["discount_rows"])))
    for w in warnings:
        log("WARNING: " + w)
    emit({"ok": True, "programs": result, "warnings": warnings})


if __name__ == "__main__":
    main()
