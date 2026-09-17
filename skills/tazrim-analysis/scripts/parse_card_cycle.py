#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""parse_card_cycle.py — card-cycle monthly statements (one file per billing cycle, all cards) -> work/normalized/<card-id>.csv (P4, F3/F4/F14/F15).

Purpose : for every `cards[]` entry with issuer "card_cycle_xlsx": glob `<prefix>_MMYY.xlsx` files, parse the
          `כרטיסי אשראי` sheet (header row 9 asserted; rows 7-8 cycle header strings recorded;
          data until a row starting "הודעה"), locate the charge date BY PATTERN (columns shift),
          skip not-yet-posted rows (`--` amount), keep installments at the installment amount,
          settle FX rows (charge date != cycle date, type חו"ל) from the card's FX account at the
          Bank of Israel representative rate of the charge date (F4; cache / API / manual, never
          failing on network errors), carry uncharged standing orders to the next cycle unless they
          reappear charged (F14), drop byte-identical files and cross-file duplicate rows (F15),
          and reconcile every file x card against the settling account's bank debit (F3, reported).
          Cycle date = `cycle_day` of month MM-1 (the file month is the month AFTER the charge).
Inputs  : tazrim.config.json (`cards[]` with issuer card_cycle_xlsx: id, last4, label, owner, files,
          settles_from, bank_debit_pattern, cycle_day, fx_settlement{currency,
          transfer_pattern?}); `fx.*`; optional `rules/fx_overrides.csv`
          (source_file,row_ref,currency,reason); work/normalized/<settles_from>.csv (from
          parse_bank.py) for the reconciliation - skipped with a warning when absent.
Outputs : work/normalized/<card-id>.csv per card (rows of unconfigured cards go to
          work/normalized/cal_unassigned.csv); one JSON object on stdout with per-file header
          strings, parsed/kept counts, reconciliation table, FX rows, carry-overs, duplicates.
Exit    : 0 ok; 1 parse failure (header mismatch / no files); 2 config error.
"""
import datetime as _dt
import os
import re
import sys
from collections import Counter, OrderedDict, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import (NORMALIZED_COLUMNS, emit, ensure_dirs, fail, log, parse_args,  # noqa: E402
                    project_path, read_csv, resolve_files, stable_id, write_csv)
from formats import ISSUERS  # noqa: E402
from fx_rates import COUNTRY_CCY, RateBook, country_in  # noqa: E402
from parse_bank import load_grid, md5, num, s, to_date  # noqa: E402

SPEC = ISSUERS["card_cycle_xlsx"]
DATE_RE = re.compile(SPEC["charge_date_pattern"])
INSTALL_RE = re.compile(SPEC["installment_pattern"])
CONV_RE = re.compile(SPEC["fx_conversion_pattern"])
FILE_RE = re.compile(r"^.*?_?(\d{2})(\d{2})\.xlsx$", re.IGNORECASE)   # <prefix>_MMYY.xlsx (any prefix)
SOURCE_LABEL = "כרטיס אשראי (קובץ מחזור)"


# ----------------------------------------------------------------------------- helpers
def cycle_date_for(fname, cycle_day):
    """<prefix>_MMYY -> ISO date of `cycle_day` in month MM-1 (year rolls back for MM=01)."""
    m = FILE_RE.match(os.path.basename(fname))
    if not m:
        raise ValueError("%s: file name must be <prefix>_MMYY.xlsx" % fname)
    mm, yy = int(m.group(1)), int(m.group(2))
    month = mm - 1 if mm > 1 else 12
    year = 2000 + yy if mm > 1 else 2000 + yy - 1
    return _dt.date(year, month, cycle_day).isoformat()


def next_cycle(iso_cycle, cycle_day):
    d = _dt.date.fromisoformat(iso_cycle)
    y, m = (d.year + 1, 1) if d.month == 12 else (d.year, d.month + 1)
    return _dt.date(y, m, cycle_day).isoformat()


def read_file(path):
    """(header7, header8, [(sheet_row, values)]) with the header row asserted."""
    rows, _ = load_grid(path, SPEC["sheet"], "openpyxl")
    hi = SPEC["header_row"] - 1
    if len(rows) <= hi:
        raise ValueError("%s: no header row %d" % (os.path.basename(path), SPEC["header_row"]))
    got = [s(v) for v in rows[hi][:len(SPEC["header"])]]
    if got != SPEC["header"]:
        raise ValueError("%s: header mismatch.\n  expected: %s\n  got:      %s" % (
            os.path.basename(path), SPEC["header"], got))
    h7 = s(rows[SPEC["cycle_header_rows"][0] - 1][0]) if len(rows) >= 7 else ""
    h8 = s(rows[SPEC["cycle_header_rows"][1] - 1][0]) if len(rows) >= 8 else ""
    data = []
    for idx in range(SPEC["data_from_row"] - 1, len(rows)):
        vals = list(rows[idx])
        first = vals[0] if vals else None
        if first is None or s(first) == "" or s(first).startswith(SPEC["stop_marker"]):
            break
        data.append((idx + 1, vals))
    return h7, h8, data


def locate_charge(vals):
    """Charge-date cell found by pattern (dd/mm/yyyy or '--') at col >= 6; amount = next numeric
    cell. Returns (charge_date_iso_or_None, charge_amount_or_None, pirut_text, uncharged_flag)."""
    idx = None
    for k in range(6, len(vals)):
        v = vals[k]
        if isinstance(v, str) and (DATE_RE.match(v.strip()) or v.strip() == SPEC["uncharged_marker"]):
            idx = k
            break
    if idx is None:
        pirut = vals[6] if len(vals) > 6 and isinstance(vals[6], str) else ""
        return None, None, (pirut or "").strip(), True
    pirut = " ".join(s(x) for x in vals[6:idx] if not (x is None or s(x) == ""))
    cell = vals[idx].strip()
    cd = to_date(cell).isoformat() if DATE_RE.match(cell) else None
    amt = num(vals[idx + 1]) if idx + 1 < len(vals) else None
    return cd, amt, pirut, cd is None


def load_fx_overrides(cfg):
    path = project_path(cfg, "rules.fx_overrides")
    out = {}
    if os.path.isfile(path):
        for r in read_csv(path, skip_comments=True):
            try:
                out[(r["source_file"].strip(), int(float(r["row_ref"])))] = (r["currency"].strip().upper(), r.get("reason", ""))
            except (KeyError, ValueError):
                log("WARNING: fx_overrides row ignored (need source_file,row_ref,currency,reason): %r" % r)
    return out


def bank_debits(cfg, account_id, pattern):
    """{iso_date: (amount, [ids])} of the settling account's rows whose description contains
    `pattern`, from work/normalized/<account_id>.csv; None when that CSV does not exist."""
    path = os.path.join(project_path(cfg, "work.normalized"), "%s.csv" % account_id)
    if not os.path.isfile(path) or not pattern:
        return None
    out = {}
    for r in read_csv(path):
        if pattern in r["original_name"]:
            amt, ids = out.get(r["charge_date"], (0.0, []))
            out[r["charge_date"]] = (round(amt + float(r["amount_ils"]), 2), ids + [r["id"]])
    return out


# ----------------------------------------------------------------------------- core
def parse_card_cycle(cfg, cards, book):
    """Parse all files of the given cycle-format cards. Returns (records_by_card_id, report)."""
    by_last4 = {}
    for c in cards:
        if c.get("last4"):
            by_last4[str(c["last4"])] = c
    files, seen_paths = [], set()
    for c in cards:
        for p in resolve_files(cfg, c.get("files")):
            if p not in seen_paths and FILE_RE.match(os.path.basename(p)):
                seen_paths.add(p)
                files.append(p)
    files.sort(key=lambda p: (cycle_date_for(p, 1), os.path.basename(p)))
    if not files:
        raise ValueError("no <prefix>_MMYY.xlsx files matched cards[].files")

    overrides = load_fx_overrides(cfg)
    digests, skipped_files, headers = {}, [], OrderedDict()
    records, unposted = [], []
    cycle_days = {c["id"]: c.get("cycle_day", SPEC["cycle_day_default"]) for c in cards}
    default_day = cards[0].get("cycle_day", SPEC["cycle_day_default"])

    for fidx, path in enumerate(files):
        fname = os.path.basename(path)
        d = md5(path)
        if d in digests:
            skipped_files.append("%s is byte-identical to %s - skipped (F15); re-download the missing cycle" % (fname, digests[d]))
            continue
        digests[d] = fname
        h7, h8, data = read_file(path)
        headers[fname] = {"row7": h7, "row8": h8, "rows": len(data)}
        for sheet_row, vals in data:
            card_cell = s(vals[0])
            m4 = re.search(r"(\d{4})\s*$", card_cell)
            card = by_last4.get(m4.group(1)) if m4 else None
            card_id = card["id"] if card else None
            cycle_day = cycle_days.get(card_id, default_day)
            cycle_date = cycle_date_for(fname, cycle_day)
            merchant = s(vals[1])
            txn_d = to_date(vals[2])
            if txn_d is None:
                unposted.append("%s r%d: bad transaction date %r" % (fname, sheet_row, s(vals[2])))
                continue
            if s(vals[3]) in (SPEC["unposted_amount_marker"], ""):
                unposted.append("%s r%d: %s - amount '%s' (not yet posted) - skipped" % (fname, sheet_row, merchant, s(vals[3])))
                continue
            txn_amt = num(vals[3])
            if txn_amt is None:
                unposted.append("%s r%d: %s - non-numeric amount %r - skipped" % (fname, sheet_row, merchant, s(vals[3])))
                continue
            txn_type = s(vals[5])
            cd, charge_amt, pirut, uncharged = locate_charge(vals)
            details = [txn_type] if txn_type else []
            if pirut:
                details.append(pirut)
            im = INSTALL_RE.match(pirut)
            if txn_type.startswith("תשלום") or im:
                details.append("תשלום %s מ-%s מתוך %g" % (im.group(1), im.group(2), txn_amt) if im
                               else "תשלום מתוך %g" % txn_amt)
            if any(k in pirut or k in merchant for k in SPEC["standing_order_markers"]) and pirut != "הוראת קבע":
                details.append("הוראת קבע")
            if uncharged or charge_amt is None:
                charge_amt = txn_amt

            # --- FX settlement decision (F4) ------------------------------------------
            fx_ccy, fx_reason = None, None
            key = (fname, sheet_row)
            if key in overrides:
                fx_ccy, fx_reason = overrides[key]
                details.append("מטבע חיוב לפי fx_overrides.csv: %s" % (fx_reason or fx_ccy))
            elif txn_type in ('חו"ל', 'זיכוי-חו"ל') and cd is not None and cd != cycle_date:
                fxs = (card or {}).get("fx_settlement") or {}
                fx_ccy = fxs.get("currency")
                if not fx_ccy:
                    details.append("שורת חו\"ל עם תאריך חיוב שונה ממחזור החיוב אך לכרטיס אין fx_settlement - נספר בש\"ח")
            orig_ccy, orig_amt, amount_ils, rate_info = "ILS", txn_amt, charge_amt, None
            if fx_ccy:
                cm = CONV_RE.search(pirut)
                country = country_in(pirut)
                if abs(charge_amt - txn_amt) < 0.005:
                    orig_ccy = fx_ccy                                  # charged 1:1 -> already in account ccy
                elif country:
                    orig_ccy = COUNTRY_CCY[country]
                    details.append("מטבע מקור משוער לפי מדינה")
                else:
                    orig_ccy = "UNK"
                    details.append("מטבע מקור לא זוהה")
                if cm:
                    details.append("סכום ביניים %s" % cm.group(1))
                hit = book.lookup(fx_ccy, cd)
                details.append('חיוב בחשבון מט"י: %.2f %s' % (charge_amt, fx_ccy))
                if hit:
                    rate, rate_day, src = hit
                    amount_ils = round(charge_amt * rate, 2)
                    details.append("המרה משוערת לפי שער %s %s %.4f ל-%s" % (
                        "יציג בנק ישראל" if src == "boi" else "ידני", fx_ccy, rate, rate_day))
                    if src == "manual":
                        details.append("שער ידני")
                    rate_info = {"rate": rate, "rate_day": rate_day, "source": src}
                else:
                    amount_ils = charge_amt
                    details.append("שער לא זמין - הסכום לא הומר (נדרש: fx.manual_rates)")
                    rate_info = {"rate": None, "rate_day": None, "source": "none"}
            records.append({
                "card_id": card_id, "card_label": card["label"] if card else card_cell, "card_cell": card_cell,
                "owner": card.get("owner") if card else None,
                "original_name": merchant, "txn_date": txn_d.isoformat(), "charge_date": cd,
                "amount_ils": round(amount_ils, 2), "orig_currency": orig_ccy, "orig_amount": round(orig_amt, 2),
                "details": details, "source_file": fname, "row_ref": sheet_row, "txn_type": txn_type,
                "cycle_date": cycle_date, "cycle_day": cycle_day, "uncharged": uncharged, "fx_ccy": fx_ccy,
                "charge_amt": round(charge_amt, 2), "file_idx": fidx, "rate": rate_info,
            })

    # --- F14: uncharged standing orders --------------------------------------------------
    charged = defaultdict(list)
    for r in records:
        if not r["uncharged"]:
            charged[(r["card_cell"], r["txn_date"], round(r["orig_amount"], 2))].append(r)
    dropped_unch, kept_unch = [], []
    for r in records:
        if not r["uncharged"]:
            continue
        hits = [h for h in charged[(r["card_cell"], r["txn_date"], round(r["orig_amount"], 2))] if h["file_idx"] > r["file_idx"]]
        if hits:
            h = hits[0]
            r["drop"] = "uncharged row re-appears charged in %s r%d (%s)" % (h["source_file"], h["row_ref"], h["charge_date"])
            dropped_unch.append(r)
        else:
            nxt = next_cycle(r["cycle_date"], r["cycle_day"])
            r["charge_date"] = nxt
            r["details"].append("תאריך חיוב חסר בקובץ; משוער למחזור החיוב הבא (%s)" % nxt)
            kept_unch.append(r)

    # --- F15: exact cross-file duplicates --------------------------------------------------
    seen, dropped_dups, same_file = {}, [], 0
    for r in records:
        if r.get("drop"):
            continue
        k = (r["card_cell"], r["original_name"], r["txn_date"], r["charge_date"], r["charge_amt"])
        if k in seen and seen[k]["source_file"] != r["source_file"]:
            r["drop"] = "exact duplicate of %s r%d" % (seen[k]["source_file"], seen[k]["row_ref"])
            dropped_dups.append(r)
        elif k in seen:
            same_file += 1
        else:
            seen[k] = r

    final = [r for r in records if not r.get("drop")]

    # --- F3: reconciliation per file x card ---------------------------------------------------
    debits = {}
    for c in cards:
        debits[c["id"]] = bank_debits(cfg, c["settles_from"], c.get("bank_debit_pattern"))
    reconcile = []
    for fname in headers:
        recs = [r for r in records if r["source_file"] == fname]
        for c in cards:
            cr = [r for r in recs if r["card_id"] == c["id"]]
            if not cr:
                continue
            cyc = cr[0]["cycle_date"]
            ils = round(sum(r["charge_amt"] for r in cr if not r["uncharged"] and r["charge_date"] == cyc and not r["fx_ccy"]), 2)
            fx = round(sum(r["charge_amt"] for r in cr if r["fx_ccy"]), 2)
            fx_ils = round(sum(r["amount_ils"] for r in cr if r["fx_ccy"]), 2)
            bd = debits.get(c["id"])
            bank = bd.get(cyc, (None, []))[0] if bd is not None else None
            reconcile.append({
                "file": fname, "cycle": cyc, "card": c["label"], "card_id": c["id"], "ils_sum": ils,
                "bank": bank, "diff": round(ils - bank, 2) if bank is not None else None,
                "bank_status": "ok" if bank is not None else ("bank csv missing" if bd is None else "no debit on cycle date"),
                "fx_sum": fx, "fx_ccy": (c.get("fx_settlement") or {}).get("currency"), "fx_sum_ils": fx_ils,
            })

    by_card = defaultdict(list)
    for r in final:
        by_card[r["card_id"] or "cal_unassigned"].append(r)
    report = {
        "files": [os.path.basename(f) for f in files], "headers": headers, "skipped_files": skipped_files,
        "parsed": len(records), "kept": len(final), "per_file": {
            f: {"parsed": sum(1 for r in records if r["source_file"] == f),
                "kept": sum(1 for r in final if r["source_file"] == f)} for f in headers},
        "unposted_skipped": unposted,
        "uncharged_dropped": len(dropped_unch), "uncharged_kept": len(kept_unch),
        "uncharged_details": ["%s r%d %s %s %.2f -> %s" % (r["source_file"], r["row_ref"], r["card_label"], r["original_name"], r["orig_amount"], r.get("drop") or "kept, charge_date=%s" % r["charge_date"]) for r in dropped_unch + kept_unch],
        "cross_file_duplicates": len(dropped_dups), "same_file_same_key_kept": same_file,
        "duplicate_details": ["%s r%d %s %.2f -> %s" % (r["source_file"], r["row_ref"], r["original_name"], r["charge_amt"], r["drop"]) for r in dropped_dups],
        "reconcile": reconcile,
        "fx_rows": [{"file": r["source_file"], "row_ref": r["row_ref"], "card": r["card_label"], "merchant": r["original_name"],
                     "charge_date": r["charge_date"], "charge_amount": r["charge_amt"], "fx_ccy": r["fx_ccy"],
                     "orig_currency": r["orig_currency"], "amount_ils": r["amount_ils"], "rate": r["rate"]} for r in final if r["fx_ccy"]],
        "unconfigured_cards": sorted({r["card_cell"] for r in final if not r["card_id"]}),
        "type_counts": dict(Counter(r["txn_type"] for r in final)),
        "total_amount_ils": round(sum(r["amount_ils"] for r in final), 2),
    }
    return by_card, report


def to_rows(card_id, recs):
    out = []
    for r in sorted(recs, key=lambda r: (r["file_idx"], r["row_ref"])):
        out.append({
            "id": stable_id(card_id, r["source_file"], r["row_ref"]), "source": SOURCE_LABEL,
            "card": r["card_label"], "original_name": r["original_name"], "txn_date": r["txn_date"],
            "charge_date": r["charge_date"], "amount_ils": "%.2f" % r["amount_ils"],
            "orig_currency": r["orig_currency"], "orig_amount": "%.2f" % r["orig_amount"],
            "details": " | ".join(x for x in r["details"] if x), "source_file": r["source_file"], "row_ref": r["row_ref"],
        })
    return out


def main(argv=None):
    def extra(p):
        p.add_argument("--offline", action="store_true", help="never call the BOI API (cache / manual rates only)")
    args, cfg = parse_args("parse card-cycle statements into work/normalized/<card-id>.csv", extra, argv=argv)
    cards = [c for c in cfg["cards"] if c["issuer"] == "card_cycle_xlsx"]
    if not cards:
        fail("no cards[] with issuer 'card_cycle_xlsx' configured", "add the cards to tazrim.config.json")
    ensure_dirs(cfg)
    book = RateBook(cfg, allow_network=not args.offline)
    try:
        by_card, report = parse_card_cycle(cfg, cards, book)
    except Exception as e:  # noqa: BLE001
        fail("card_cycle: %s" % e, "check the file layout against references/bank-formats.md (formats.py 'card_cycle_xlsx')")
        return
    out_dir = project_path(cfg, "work.normalized")
    outputs = {}
    for c in cards:
        rows = to_rows(c["id"], by_card.get(c["id"], []))
        path = os.path.join(out_dir, "%s.csv" % c["id"])
        write_csv(path, rows, NORMALIZED_COLUMNS)
        outputs[c["id"]] = {"csv": path, "rows": len(rows), "label": c["label"]}
    if by_card.get("cal_unassigned"):
        path = os.path.join(out_dir, "cal_unassigned.csv")
        write_csv(path, to_rows("cal_unassigned", by_card["cal_unassigned"]), NORMALIZED_COLUMNS)
        outputs["cal_unassigned"] = {"csv": path, "rows": len(by_card["cal_unassigned"]),
                                     "note": "rows of card numbers not in cards[] - add them to the config"}
    warnings = list(cfg.warnings) + report["skipped_files"] + book.problems
    for rr in report["reconcile"]:
        if rr["diff"] not in (None, 0.0):
            warnings.append("reconcile %s %s %s: statement %.2f vs bank %.2f (diff %.2f) - explain (FX row typed ישראל? see fx_overrides.csv)" % (
                rr["file"], rr["cycle"], rr["card"], rr["ils_sum"], rr["bank"], rr["diff"]))
    for w in warnings:
        log("WARNING: " + w)
    log("card_cycle: parsed %d, kept %d, uncharged dropped %d kept %d, cross-file dups %d, fx rows %d" % (
        report["parsed"], report["kept"], report["uncharged_dropped"], report["uncharged_kept"],
        report["cross_file_duplicates"], len(report["fx_rows"])))
    emit({"ok": True, "cards": outputs, "report": report, "fx_problems": book.problems, "warnings": warnings})


if __name__ == "__main__":
    main()
