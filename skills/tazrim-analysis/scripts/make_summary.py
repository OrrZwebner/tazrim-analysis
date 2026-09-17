#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_summary.py — independent pandas recomputation of every report figure (P14 + F12).

Purpose : recompute, from work/database.csv alone and with the same definitions as the workbook
          formulas (F8/F9/F16/F22/F23/F25), everything the figures, the report and the hand-over
          need, and the CASH RECONCILIATION identity (F12) with its residual — reported, never hidden.
Inputs  : tazrim.config.json (window, people, accounts/cards/benefit programs/p2p, fixed
          categories, thresholds); work/database.csv (exactly common.DB_COLUMNS);
          optional work/normalized/*.csv bank files whose `balance` column (or a `יתרה: <n>` token
          in `details`) gives the bank balance change over the window.
Outputs : work/summary.json (all figures); ONE JSON digest on stdout:
          {"ok": true, "summary": <path>, "n_months": N, "n_rows": n, "totals": {...},
           "recon": {"bank_change", "explained", "residual", "residual_ok", "bank_change_source"}}
Exit    : 0 ok; 1 database missing / unreadable; 2 config error (via common.parse_args);
          4 with --strict when |residual| >= 1.

Definitions (all on types/flags/config, never on merchant or person names):
  summed rows    = in_window == כן and type in (הוצאה, הכנסה)          (F16)
  month series   = one value per window month, month without a row = 0  (F8)
  avg            = total / N ; avg_nz = total / months with a non-zero value (F25)
  stdev          = sample standard deviation over the N monthly values (STDEV.S), 0 when N < 2
  sign           = `amount` > 0 for expenses and incomes (type carries direction); for the
                   non-summed types the normalized sign is kept: > 0 money out, < 0 money in.
  source kind    = a row belongs to an account/card/benefit program/p2p entry when its `source`,
                   `card` or `pay` equals the entry id or label (or starts with the label);
                   rows matching no entry are "other" (estimates, computed rows).

Cash reconciliation identity (F12), every term in ₪ over the window:
  bank_change = income − expenses − savings_bank − internal_bank_net − reimb_paid_bank
                + reimb_recv_bank + fx_expenses − p2p_income + p2p_paid_from_balance
                − card_duplicates_unlinked − card_nonexpense_rows + Σ(club_face_net − club_bank)
                + nonstatement_expenses − nonbank_income + residual
  The card-statement-vs-bank-debit difference (F3) is deliberately NOT part of the explained
  side, so it shows up in the residual; it is reported per card next to it.
Python 3.8 compatible. Needs pandas.
"""
import datetime as _dt
import glob
import math
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import (  # noqa: E402
    DB_COLUMNS, NO, SHARED, TYPE_CARD_DEBIT, TYPE_DUPLICATE, TYPE_EXPENSE, TYPE_INCOME,
    TYPE_INTERNAL, TYPE_REIMBURSABLE, TYPE_REIMBURSEMENT, TYPE_SAVINGS, UNKNOWN_CAT, YES,
    emit, fail, log, month_label, months, owner_label, parse_args, project_path, write_json)

try:
    import pandas as pd
except ImportError:  # pragma: no cover
    pd = None

MONEY_IN_TYPES = (TYPE_INCOME, TYPE_REIMBURSEMENT)
ILS = ("ILS", "", "₪", "NIS")
CHECK_MARKERS = ("לבדיקה", "לסיווג ידני", "נדרש:")


# ----------------------------------------------------------------------------- helpers
def _r2(x):
    try:
        return round(float(x), 2)
    except (TypeError, ValueError):
        return 0.0


def _series(df, M):
    """Monthly sums over the window months (missing month = 0)."""
    if df.empty:
        return [0.0] * len(M)
    s = df.groupby("month")["amount"].sum()
    return [_r2(s.get(m, 0.0)) for m in M]


def stats(vals, M):
    """F8 statistics over the N monthly values."""
    n = len(vals)
    total = sum(vals)
    nz = [v for v in vals if abs(v) > 0.005]
    mean = total / n if n else 0.0
    if n >= 2:
        sd = math.sqrt(sum((v - mean) ** 2 for v in vals) / (n - 1))
    else:
        sd = 0.0
    srt = sorted(vals)
    if n == 0:
        med = 0.0
    elif n % 2:
        med = srt[n // 2]
    else:
        med = (srt[n // 2 - 1] + srt[n // 2]) / 2.0
    imax = max(range(n), key=lambda i: vals[i]) if n else None
    imin = min(range(n), key=lambda i: vals[i]) if n else None
    return {
        "total": _r2(total), "avg": _r2(mean), "avg_nz": _r2(sum(nz) / len(nz)) if nz else 0.0,
        "n_nonzero": len(nz), "median": _r2(med), "stdev": _r2(sd),
        "max": _r2(vals[imax]) if n else 0.0, "max_month": M[imax] if n else None,
        "min": _r2(vals[imin]) if n else 0.0, "min_month": M[imin] if n else None,
        "by_month": [_r2(v) for v in vals],
    }


def _matches(row_vals, entry):
    """True when any of the row's identifying strings names this config entry."""
    keys = [str(entry.get("id") or ""), str(entry.get("label") or "")]
    keys = [k for k in keys if k]
    for v in row_vals:
        v = str(v or "")
        if not v:
            continue
        for k in keys:
            if v == k or v.startswith(k + " ") or v.startswith(k + " ("):
                return True
    return False


def source_index(cfg, db):
    """Return (kind Series, entry-id Series) for every database row."""
    kinds, ids = [], []
    sections = [("account", cfg.get("accounts", [])), ("card", cfg.get("cards", [])),
                ("benefit", cfg.get("benefit_programs", [])), ("p2p", cfg.get("p2p", []))]
    for src, card, pay in zip(db["source"], db["card"], db["pay"]):
        kind, eid = "other", ""
        for k, entries in sections:
            for e in entries:
                if _matches((src, card, pay), e) or (k == "card" and e.get("last4") and
                                                     str(e["last4"]) in str(card or "")):
                    kind, eid = k, e["id"]
                    break
            if eid:
                break
        kinds.append(kind)
        ids.append(eid)
    return pd.Series(kinds, index=db.index), pd.Series(ids, index=db.index)


def cash_out(df):
    """Signed cash flow of rows from the payer's point of view: + = money out."""
    if df.empty:
        return 0.0
    sign = df["type"].isin(MONEY_IN_TYPES).map(lambda x: -1.0 if x else 1.0)
    return float((df["amount"] * sign).sum())


def _reconstruct_balances(sub, start, end, cfg_opening):
    """F1/F2 balance reconstruction over ONE ordering of an account's normalized rows.

    `sub` carries _d (date), _amt (amount_ils, > 0 = money out), _bal (balance AFTER the row or
    NaN). The balance before the first row is anchored on the first balanced row
    (before = after + out) and walked back over the earlier rows; then the running balance is
    walked forward, compared with every balance checkpoint (tolerance 0.005, re-anchored on each),
    and read off just before the first in-window row (pre) and after the last in-window row (post).
    Returns None when the account has no balance checkpoint (and no configured opening)."""
    dates = sub["_d"].tolist()
    amts = [float(a) for a in sub["_amt"].tolist()]
    bals = [None if (isinstance(b, float) and math.isnan(b)) else float(b) for b in sub["_bal"].tolist()]
    anchor = next((i for i, b in enumerate(bals) if b is not None), None)
    if anchor is None:
        if cfg_opening is None:
            return None
        before0 = float(cfg_opening)
    else:
        before0 = bals[anchor] + amts[anchor] + sum(amts[:anchor])
    running, pre, post, breaks, n_in = before0, None, None, 0, 0
    for i, d in enumerate(dates):
        if d >= start and pre is None:
            pre = running
        running -= amts[i]
        if bals[i] is not None:
            if abs(running - bals[i]) > 0.005:
                breaks += 1
            running = bals[i]
        if start <= d <= end:
            post, n_in = running, n_in + 1
    if pre is None or post is None or not n_in:
        return None
    if cfg_opening is not None and anchor is not None and abs(float(cfg_opening) - before0) > 0.005:
        breaks += 1  # configured opening disagrees with the statement's own chain
    return {"balance_before": _r2(pre), "balance_end": _r2(post), "change": _r2(post - pre),
            "chain_breaks": breaks, "opening_file": _r2(before0), "rows_in_window": n_in}


def bank_change_from_balances(cfg, start, end):
    """Bank balance change over the window from the normalized bank CSVs (balance column or
    'יתרה: n' in details). Returns (total or None, per-account dict, note).

    Rows are sorted by date, but the order INSIDE a date depends on the statement direction
    (newest-first files carry the day's closing balance on the lowest row_ref). Both within-date
    orders (row_ref ascending / descending) are reconstructed with F1/F2 and the one under which
    the statement's own balance chain has fewer breaks is used; ties keep the ascending order."""
    norm_dir = project_path(cfg, "work.normalized")
    files = sorted(glob.glob(os.path.join(norm_dir, "*.csv")))
    if not files:
        return None, {}, "no normalized CSVs"
    accounts = cfg.get("accounts", [])
    per = {}
    for f in files:
        try:
            df = pd.read_csv(f, encoding="utf-8-sig", dtype=str, keep_default_na=False)
        except Exception as e:  # noqa: BLE001
            log("WARNING: cannot read %s: %s" % (f, e))
            continue
        if df.empty or "source" not in df.columns:
            continue
        for a in accounts:
            if a["id"] in per:
                continue
            sub = df[[_matches((s, "", ""), a) for s in df["source"]]].copy()
            if sub.empty:
                continue
            if "balance" in sub.columns:
                sub["_bal"] = pd.to_numeric(sub["balance"], errors="coerce")
            else:
                sub["_bal"] = sub["details"].map(_balance_from_details) if "details" in sub.columns else float("nan")
            sub["_amt"] = pd.to_numeric(sub.get("amount_ils", 0), errors="coerce").fillna(0.0)
            sub["_d"] = sub["txn_date"].astype(str).str[:10]
            sub["_file"] = sub["source_file"].astype(str) if "source_file" in sub.columns else ""
            sub["_ref"] = pd.to_numeric(sub.get("row_ref", 0), errors="coerce").fillna(0)
            cfg_open = a.get("opening_balance")
            best = None
            for order, asc in (("row_ref_asc", True), ("row_ref_desc", False)):
                s_ = sub.sort_values(["_d", "_file", "_ref"], ascending=[True, True, asc], kind="mergesort")
                res = _reconstruct_balances(s_, start, end, cfg_open)
                if res is None:
                    continue
                res["order"] = order
                if best is None or res["chain_breaks"] < best["chain_breaks"]:
                    best = res
            if best is None:
                continue
            if best["chain_breaks"]:
                log("WARNING: %s: %d balance-chain break(s) while deriving the window balances" % (a["id"], best["chain_breaks"]))
            best.update({"label": a.get("label", a["id"]), "file": os.path.basename(f)})
            per[a["id"]] = best
    if not per:
        return None, {}, "no balance column / יתרה token found in normalized bank CSVs"
    return _r2(sum(v["change"] for v in per.values())), per, "balances from normalized bank CSVs (F1/F2 reconstruction)"


_BAL_RE = re.compile(r"יתרה[:\s]+(-?[\d,]+(?:\.\d+)?)")


def _balance_from_details(text):
    m = _BAL_RE.search(str(text or ""))
    if not m:
        return float("nan")
    try:
        return float(m.group(1).replace(",", ""))
    except ValueError:
        return float("nan")


# ----------------------------------------------------------------------------- main build
def build_summary(cfg, db_path=None):
    """Compute the summary dict from work/database.csv. Raises on a missing/invalid database."""
    if pd is None:
        raise RuntimeError("pandas is required (pip install -r requirements.txt)")
    db_path = db_path or project_path(cfg, "work.database")
    if not os.path.isfile(db_path):
        raise FileNotFoundError(db_path)
    db = pd.read_csv(db_path, encoding="utf-8-sig", dtype=str, keep_default_na=False)
    missing = [c for c in DB_COLUMNS if c not in db.columns]
    if missing:
        raise ValueError("database.csv is missing columns: %s" % ", ".join(missing))
    db["amount"] = pd.to_numeric(db["amount"], errors="coerce").fillna(0.0)
    for c in ("details", "rule_note", "linked_id", "trip", "card", "pay", "person", "orig_currency"):
        db[c] = db[c].fillna("").astype(str)
    db["_in"] = db["in_window"] == YES
    db["_sum"] = db["summed"] == YES
    db["_kind"], db["_eid"] = source_index(cfg, db)
    db["_ils"] = db["orig_currency"].str.strip().str.upper().isin([x.upper() for x in ILS])

    M = months(cfg)
    N = len(M)
    MN = [month_label(m) for m in M]
    start, end = cfg["window"]["start"], cfg["window"]["end"]
    W = db[db["_in"]]
    S_ = db[db["_sum"]]
    E = S_[S_["type"] == TYPE_EXPENSE]
    I = S_[S_["type"] == TYPE_INCOME]
    thr = cfg.get("thresholds", {})
    top_n = int(thr.get("top_merchants", 25) or 25)

    S = {
        "generated": _dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "level": cfg.level, "household": cfg["household"].get("label", ""),
        "window": {"start": start, "end": end, "label": "%s..%s" % (start, end)},
        "months": M, "month_names": MN, "n_months": N,
        "n_rows_db": int(len(db)), "n_rows_in_window": int(len(W)),
        "n_expense_rows": int(len(E)), "n_income_rows": int(len(I)),
        "rows_by_type": {k: int(v) for k, v in db["type"].value_counts().items()},
        "rows_by_type_in_window": {k: int(v) for k, v in W["type"].value_counts().items()},
    }

    # --- totals and monthly series
    S["expenses_by_month"] = _series(E, M)
    S["income_by_month"] = _series(I, M)
    S["balance_by_month"] = [_r2(i - e) for i, e in zip(S["income_by_month"], S["expenses_by_month"])]
    sav = W[W["type"] == TYPE_SAVINGS]
    intern = W[W["type"] == TYPE_INTERNAL]
    S["savings_by_month"] = _series(sav, M)
    S["internal_by_month"] = _series(intern, M)
    S["total_expenses"] = _r2(E["amount"].sum())
    S["total_income"] = _r2(I["amount"].sum())
    S["avg_expenses"] = _r2(S["total_expenses"] / N)
    S["avg_income"] = _r2(S["total_income"] / N)
    S["balance_total"] = _r2(S["total_income"] - S["total_expenses"])
    S["avg_balance"] = _r2(S["balance_total"] / N)
    S["savings_rate"] = _r2(S["balance_total"] / S["total_income"]) if S["total_income"] else 0.0
    S["savings_total"] = _r2(sav["amount"].sum())
    S["savings_by_cat"] = {k: _r2(v) for k, v in sav.groupby("cat_tz")["amount"].sum().sort_values(ascending=False).items()}
    S["expense_stats"] = stats(S["expenses_by_month"], M)
    S["income_stats"] = stats(S["income_by_month"], M)

    # --- groups / categories (primary and secondary scheme)
    def by_group(df, gcol, ccol):
        groups = {}
        for g, sub in df.groupby(gcol):
            groups[g] = stats(_series(sub, M), M)
            groups[g]["n_rows"] = int(len(sub))
            groups[g]["share"] = _r2(sub["amount"].sum() / df["amount"].sum()) if df["amount"].sum() else 0.0
        cats = {}
        for (g, c), sub in df.groupby([gcol, ccol]):
            st = stats(_series(sub, M), M)
            st.update({"group": g, "cat": c, "n_rows": int(len(sub)),
                       "share": _r2(sub["amount"].sum() / df["amount"].sum()) if df["amount"].sum() else 0.0})
            cats["%s / %s" % (g, c)] = st
        groups = dict(sorted(groups.items(), key=lambda kv: -kv[1]["total"]))
        cats = dict(sorted(cats.items(), key=lambda kv: (-groups.get(kv[1]["group"], {}).get("total", 0), -kv[1]["total"])))
        return groups, cats

    S["expenses_by_group"], S["expenses_by_cat"] = by_group(E, "group_tz", "cat_tz")
    S["income_by_group"], S["income_by_cat"] = by_group(I, "group_tz", "cat_tz")
    has_new = bool((E["cat_new"].str.strip() != "").any()) and not (E["cat_new"] == E["cat_tz"]).all()
    S["has_secondary_scheme"] = has_new
    if has_new:
        S["expenses_new_by_group"], S["expenses_new_by_cat"] = by_group(E, "group_new", "cat_new")
    S["income_by_source"] = {c: st for c, st in S["income_by_cat"].items()}

    # --- fixed vs variable (F22), Israel vs abroad (F23)
    fixed = [str(x) for x in cfg.get("fixed_categories", [])]
    S["fixed_categories"] = fixed
    S["fixed_by_month"] = _series(E[E["cat_tz"].isin(fixed)], M)
    S["variable_by_month"] = [_r2(t - f) for t, f in zip(S["expenses_by_month"], S["fixed_by_month"])]
    S["fixed_total"] = _r2(sum(S["fixed_by_month"]))
    S["variable_total"] = _r2(sum(S["variable_by_month"]))
    S["israel_by_month"] = _series(E[E["_ils"]], M)
    S["abroad_by_month"] = _series(E[~E["_ils"]], M)
    S["israel_total"] = _r2(E[E["_ils"]]["amount"].sum())
    S["abroad_total"] = _r2(E[~E["_ils"]]["amount"].sum())
    S["abroad_by_currency"] = {k: {"total": _r2(v["amount"].sum()), "n": int(len(v))}
                               for k, v in E[~E["_ils"]].groupby("orig_currency")}

    # --- by payment method / card / person
    S["expenses_by_pay"] = {k: dict(stats(_series(v, M), M), n_rows=int(len(v)))
                            for k, v in sorted(E.groupby("pay"), key=lambda kv: -kv[1]["amount"].sum())}
    cards_out = {}
    for e in cfg.get("cards", []) + cfg.get("benefit_programs", []):
        sub = E[E["_eid"] == e["id"]]
        if len(sub):
            cards_out[e.get("label", e["id"])] = dict(stats(_series(sub, M), M), n_rows=int(len(sub)),
                                                       owner=owner_label(cfg, e.get("owner", SHARED)))
    S["expenses_by_card"] = dict(sorted(cards_out.items(), key=lambda kv: -kv[1]["total"]))
    people = [owner_label(cfg, p["id"]) for p in cfg["household"]["people"]] + [owner_label(cfg, SHARED)]
    S["people"] = people

    def by_person(df):
        out = {}
        present = list(df["person"].unique())
        for p in people + [x for x in present if x not in people]:
            sub = df[df["person"] == p]
            if len(sub) or p in people:
                out[p] = dict(stats(_series(sub, M), M), n_rows=int(len(sub)))
        return out
    S["expenses_by_person"] = by_person(E)
    S["income_by_person"] = by_person(I)

    # --- merchants
    merch = []
    for name, sub in E.groupby("name_clean"):
        st = stats(_series(sub, M), M)
        st.update({"name": name, "count": int(len(sub)), "cat": sub["cat_tz"].mode().iloc[0] if len(sub) else "",
                   "group": sub["group_tz"].mode().iloc[0] if len(sub) else "",
                   "avg_txn": _r2(sub["amount"].sum() / len(sub)) if len(sub) else 0.0})
        merch.append(st)
    S["top_merchants"] = sorted(merch, key=lambda m: -m["total"])[:top_n]
    S["top_merchants_by_avg"] = sorted(merch, key=lambda m: -m["avg"])[:top_n]
    S["top_merchants_by_avg_nz"] = sorted(merch, key=lambda m: -m["avg_nz"])[:top_n]
    S["n_merchants"] = len(merch)
    S["variable_merchants"] = sorted([m for m in merch if m["cat"] not in fixed], key=lambda m: -m["total"])
    # recurring merchants: a charge in every window month with low variation (generic heuristic)
    rec = []
    for m in merch:
        if N >= 2 and m["n_nonzero"] == N and m["avg"] > 0 and (m["stdev"] / m["avg"]) < 0.10:
            rec.append(m)
    S["recurring_merchants"] = sorted(rec, key=lambda m: -m["total"])
    dbl = E.groupby(["name_clean", "txn_date", "amount"]).size()
    S["possible_double_charges"] = [{"name": k[0], "date": k[1], "amount": _r2(k[2]), "n": int(v)}
                                    for k, v in dbl[dbl >= 2].items() if k[2] > 0]
    S["refunds"] = [{"id": r.id, "name": r.name_clean, "date": r.txn_date, "amount": _r2(r.amount), "cat": r.cat_tz}
                    for r in E[E["amount"] < 0].sort_values("amount").itertuples()]

    # --- unknown / check items
    unk_mask = (S_["cat_tz"] == UNKNOWN_CAT) | S_["rule_note"].str.contains("|".join(CHECK_MARKERS), regex=True)
    unk = S_[unk_mask]
    S["unknown_items"] = [{"id": r.id, "date": r.txn_date, "name": r.original_name, "amount": _r2(r.amount),
                           "source": r.source, "type": r.type, "cat": r.cat_tz, "note": r.rule_note}
                          for r in unk.sort_values("amount", ascending=False).itertuples()]
    S["unknown_total"] = _r2(E[E["cat_tz"] == UNKNOWN_CAT]["amount"].sum())
    S["n_unknown"] = int(len(unk))

    # --- trips (from the trip column; labels from config when configured)
    trips = {}
    for t, sub in E[E["trip"].str.strip() != ""].groupby("trip"):
        trips[t] = {"total": _r2(sub["amount"].sum()), "n": int(len(sub)),
                    "items": [{"name": r.name_clean, "date": r.txn_date, "amount": _r2(r.amount)}
                              for r in sub.sort_values("amount", ascending=False).head(15).itertuples()]}
    S["trips"] = dict(sorted(trips.items(), key=lambda kv: -kv[1]["total"]))
    S["trips_configured"] = [t.get("label") for t in cfg.get("trips", [])]

    # --- reimbursables (F21)
    rp = W[W["type"] == TYPE_REIMBURSABLE]
    rr = W[W["type"] == TYPE_REIMBURSEMENT]
    S["reimbursables"] = {
        "present": bool(len(rp) or len(rr)),
        "paid_total": _r2(rp["amount"].sum()), "received_total": _r2(rr["amount"].sum()),
        "open_balance": _r2(rp["amount"].sum() - rr["amount"].sum()),
        "paid_by_month": _series(rp, M), "received_by_month": _series(rr, M),
        "items": [{"id": r.id, "date": r.txn_date, "name": r.name_clean, "type": r.type, "cat": r.cat_tz,
                   "amount": _r2(r.amount), "note": r.rule_note}
                  for r in pd.concat([rp, rr]).sort_values("txn_date").itertuples()],
    }

    # --- card debits (F3 / F13) per configured card and benefit program
    deb = W[W["type"] == TYPE_CARD_DEBIT]
    bank_rows = W[W["_kind"] == "account"]
    card_table = []
    matched_idx = set()
    for e in cfg.get("cards", []) + cfg.get("benefit_programs", []):
        pat = e.get("bank_debit_pattern")
        d = deb[deb["original_name"].str.contains(re.escape(pat), regex=True)] if pat else deb.iloc[0:0]
        matched_idx.update(d.index.tolist())
        rows_e = W[W["_eid"] == e["id"]]
        is_club = e in cfg.get("benefit_programs", [])
        stmt = cash_out(rows_e[rows_e["_ils"] & (rows_e["type"] != TYPE_INTERNAL)]) if is_club else cash_out(rows_e[rows_e["_ils"]])
        card_table.append({"id": e["id"], "label": e.get("label", e["id"]), "kind": "benefit" if is_club else "card",
                           "bank_debits_by_month": _series(d, M), "bank_debits": _r2(d["amount"].sum()),
                           "statement_ils_by_month": _series(rows_e[rows_e["_ils"] & (rows_e["type"] != TYPE_INTERNAL)] if is_club else rows_e[rows_e["_ils"]], M),
                           "statement_ils": _r2(stmt), "fx_expenses": _r2(rows_e[(~rows_e["_ils"]) & (rows_e["type"] == TYPE_EXPENSE) & rows_e["_sum"]]["amount"].sum()),
                           "diff": _r2(d["amount"].sum() - stmt) if not is_club else None,
                           "pattern": pat or ""})
    unmatched = deb[~deb.index.isin(matched_idx)]
    S["card_debits"] = {"table": card_table, "total": _r2(deb["amount"].sum()),
                        "unmatched_total": _r2(unmatched["amount"].sum()),
                        "unmatched": [{"id": r.id, "name": r.original_name, "date": r.txn_date, "amount": _r2(r.amount)}
                                      for r in unmatched.itertuples()]}

    # --- P2P (F6)
    p2p_rows = W[W["_kind"] == "p2p"]
    dup_cards = W[(W["type"] == TYPE_DUPLICATE) & (W["_kind"] == "card") & W["_ils"]]
    p2p_exp = p2p_rows[(p2p_rows["type"] == TYPE_EXPENSE) & p2p_rows["_sum"]]
    p2p_inc = p2p_rows[(p2p_rows["type"] == TYPE_INCOME) & p2p_rows["_sum"]]
    linked_ok = dup_cards["linked_id"].isin(set(p2p_exp["id"]))
    p2p_funded = float(dup_cards[linked_ok]["amount"].sum())
    dup_unlinked = float(dup_cards[~linked_ok]["amount"].sum())
    markers = [p.get("card_marker") for p in cfg.get("p2p", []) if p.get("card_marker")]
    unmatched_p2p_cards = W.iloc[0:0]
    if markers:
        mk = W["original_name"].str.contains("|".join(re.escape(m) for m in markers), case=False, regex=True)
        unmatched_p2p_cards = W[mk & (W["_kind"] == "card") & (W["type"] == TYPE_EXPENSE) & W["_sum"]]
    S["p2p"] = {
        "configured": bool(cfg.get("p2p")), "present": bool(len(p2p_rows) or len(dup_cards)),
        "income": _r2(p2p_inc["amount"].sum()), "expenses": _r2(p2p_exp["amount"].sum()),
        "expenses_by_month": _series(p2p_exp, M), "income_by_month": _series(p2p_inc, M),
        "matched_pairs": int(len(dup_cards[linked_ok])), "funded_by_card_in_window": _r2(p2p_funded),
        "paid_from_balance": _r2(p2p_exp["amount"].sum() - p2p_funded),
        "card_duplicates_unlinked": _r2(dup_unlinked),
        "unmatched_card_rows": {"n": int(len(unmatched_p2p_cards)), "total": _r2(unmatched_p2p_cards["amount"].sum()),
                                "by_person": {k: _r2(v) for k, v in unmatched_p2p_cards.groupby("person")["amount"].sum().items()}},
        "income_by_cat": {k: _r2(v) for k, v in p2p_inc.groupby("cat_tz")["amount"].sum().items()},
    }

    # --- benefit programs (F7)
    clubs = []
    for e in cfg.get("benefit_programs", []):
        rows_e = W[W["_eid"] == e["id"]]
        exp_e = rows_e[(rows_e["type"] == TYPE_EXPENSE) & rows_e["_sum"]]
        loads = rows_e[rows_e["type"] == TYPE_INTERNAL]
        disc = exp_e[(exp_e["amount"] < 0) | (exp_e["source_file"] == "מחושב")]
        pat = e.get("bank_debit_pattern")
        d = deb[deb["original_name"].str.contains(re.escape(pat), regex=True)] if pat else deb.iloc[0:0]
        face_net = float(exp_e["amount"].sum())
        clubs.append({"id": e["id"], "label": e.get("label", e["id"]),
                      "purchases_face": _r2(exp_e[exp_e["amount"] > 0]["amount"].sum()),
                      "discount_total": _r2(-disc["amount"].sum()), "discount_by_month": [_r2(-v) for v in _series(disc, M)],
                      "loads": _r2(loads["amount"].sum()), "face_net": _r2(face_net),
                      "bank_debits": _r2(d["amount"].sum()), "bank_debits_by_month": _series(d, M),
                      "balance_change": _r2(face_net - float(d["amount"].sum())),
                      "discount_rate": _r2(-disc["amount"].sum() / exp_e[exp_e["amount"] > 0]["amount"].sum())
                      if float(exp_e[exp_e["amount"] > 0]["amount"].sum()) else 0.0})
    S["benefit_programs"] = clubs

    # --- FX (F4) — card expenses settled in a foreign currency
    fx_exp = E[(E["_kind"] == "card") & (~E["_ils"])]
    S["fx"] = {"configured": any(c.get("fx_settlement") for c in cfg.get("cards", [])),
               "expenses": _r2(fx_exp["amount"].sum()), "n_rows": int(len(fx_exp)),
               "by_currency": {k: _r2(v) for k, v in fx_exp.groupby("orig_currency")["amount"].sum().items()},
               "by_card": {k: _r2(v) for k, v in fx_exp.groupby("pay")["amount"].sum().items()},
               "estimated_conversions": int(fx_exp["details"].str.contains("המרה משוערת").sum())}

    # --- estimates / other-source rows
    other_exp = E[E["_kind"] == "other"]
    other_inc = I[~I["_kind"].isin(["account", "p2p"])]
    S["nonstatement_expenses"] = {"total": _r2(other_exp["amount"].sum()), "n": int(len(other_exp)),
                                  "by_source": {k: _r2(v) for k, v in other_exp.groupby("source")["amount"].sum().items()},
                                  "configured_estimates": len(cfg.get("estimates", []))}
    S["nonbank_income"] = {"total": _r2(other_inc["amount"].sum()), "n": int(len(other_inc)),
                           "by_source": {k: _r2(v) for k, v in other_inc.groupby("source")["amount"].sum().items()}}

    # --- internal transfers in the bank, by category (signed, + = out)
    int_bank = bank_rows[bank_rows["type"] == TYPE_INTERNAL]
    S["internal_bank"] = {"net": _r2(cash_out(int_bank)), "out": _r2(int_bank[int_bank["amount"] > 0]["amount"].sum()),
                          "in": _r2(-int_bank[int_bank["amount"] < 0]["amount"].sum()),
                          "by_cat": {k: _r2(v) for k, v in int_bank.groupby("cat_tz")["amount"].sum().sort_values().items()},
                          "by_month": _series(int_bank, M)}

    # --- CASH RECONCILIATION (F12)
    savings_bank = float(bank_rows[bank_rows["type"] == TYPE_SAVINGS]["amount"].sum())
    reimb_paid_bank = float(bank_rows[bank_rows["type"] == TYPE_REIMBURSABLE]["amount"].sum())
    reimb_recv_bank = float(bank_rows[bank_rows["type"] == TYPE_REIMBURSEMENT]["amount"].sum())
    card_rows_ils = W[(W["_kind"] == "card") & W["_ils"]]
    card_nonexp = card_rows_ils[~card_rows_ils["type"].isin([TYPE_EXPENSE, TYPE_DUPLICATE])]
    card_nonexp_out = cash_out(card_nonexp)
    bank_change_rows = cash_out(bank_rows) * -1.0  # cash_out is + for money out; change = money in
    bal_change, per_acct, bal_note = bank_change_from_balances(cfg, start, end)
    comps = [
        ("income", "+", S["total_income"], True),
        ("expenses", "-", S["total_expenses"], True),
        ("savings_bank", "-", _r2(savings_bank), True),
        ("internal_bank_net", "-", S["internal_bank"]["net"], True),
        ("reimb_paid_bank", "-", _r2(reimb_paid_bank), S["reimbursables"]["present"]),
        ("reimb_recv_bank", "+", _r2(reimb_recv_bank), S["reimbursables"]["present"]),
        ("fx_expenses", "+", S["fx"]["expenses"], S["fx"]["configured"] or S["fx"]["expenses"] != 0),
        ("p2p_income", "-", S["p2p"]["income"], S["p2p"]["configured"] or S["p2p"]["present"]),
        ("p2p_paid_from_balance", "+", S["p2p"]["paid_from_balance"], S["p2p"]["configured"] or S["p2p"]["present"]),
        ("card_duplicates_unlinked", "-", S["p2p"]["card_duplicates_unlinked"], S["p2p"]["card_duplicates_unlinked"] != 0),
        ("card_nonexpense_rows", "-", _r2(card_nonexp_out), card_nonexp_out != 0),
        ("club_balance_change", "+", _r2(sum(c["balance_change"] for c in clubs)), bool(clubs)),
        ("nonstatement_expenses", "+", S["nonstatement_expenses"]["total"], S["nonstatement_expenses"]["n"] > 0),
        ("nonbank_income", "-", S["nonbank_income"]["total"], S["nonbank_income"]["n"] > 0),
    ]
    explained = 0.0
    for _k, sign, val, _present in comps:
        explained += val if sign == "+" else -val
    explained = _r2(explained)
    bank_change = bal_change if bal_change is not None else _r2(bank_change_rows)
    residual = _r2(bank_change - explained)
    S["recon"] = {
        "bank_change": bank_change,
        "bank_change_source": "balances" if bal_change is not None else "rows",
        "bank_change_from_rows": _r2(bank_change_rows),
        "bank_change_from_balances": bal_change, "balances_note": bal_note, "per_account": per_acct,
        "components": [{"key": k, "sign": s, "value": v, "present": bool(p)} for k, s, v, p in comps],
        "explained": explained, "residual": residual, "residual_ok": abs(residual) < 1.0,
        "card_vs_bank_diff": _r2(sum(c["diff"] or 0.0 for c in card_table if c["kind"] == "card")),
        "card_debits_unmatched": S["card_debits"]["unmatched_total"],
        "formula": "bank_change = income - expenses - savings_bank - internal_bank_net - reimb_paid_bank + reimb_recv_bank"
                   " + fx_expenses - p2p_income + p2p_paid_from_balance - card_duplicates_unlinked - card_nonexpense_rows"
                   " + club_balance_change + nonstatement_expenses - nonbank_income + residual",
    }
    return S


def main(argv=None):
    def extra(p):
        p.add_argument("--database", default=None, help="override work/database.csv path")
        p.add_argument("--out", default=None, help="override work/summary.json path")
        p.add_argument("--strict", action="store_true", help="exit 4 when |residual| >= 1")
    args, cfg = parse_args("independent recomputation of the report figures + cash reconciliation (P14/F12)", extra, argv)
    try:
        S = build_summary(cfg, args.database)
    except FileNotFoundError as e:
        fail("database not found: %s" % e, "run classify.py first (work/database.csv)")
    except (ValueError, RuntimeError) as e:
        fail(str(e), "database.csv must carry exactly common.DB_COLUMNS")
    out = args.out or project_path(cfg, "work.summary")
    write_json(out, S)
    rc = S["recon"]
    log("summary: %d rows, N=%d months, expenses %.2f income %.2f | recon %s: bank_change %.2f explained %.2f residual %.2f"
        % (S["n_rows_db"], S["n_months"], S["total_expenses"], S["total_income"], rc["bank_change_source"],
           rc["bank_change"], rc["explained"], rc["residual"]))
    if not rc["residual_ok"]:
        log("WARNING: reconciliation residual %.2f (card-vs-bank diff %.2f, unmatched card debits %.2f)"
            % (rc["residual"], rc["card_vs_bank_diff"], rc["card_debits_unmatched"]))
    digest = {"ok": True, "summary": out, "n_months": S["n_months"], "n_rows": S["n_rows_db"],
              "totals": {"expenses": S["total_expenses"], "income": S["total_income"],
                         "avg_expenses": S["avg_expenses"], "avg_income": S["avg_income"],
                         "savings": S["savings_total"], "unknown": S["unknown_total"]},
              "recon": {k: rc[k] for k in ("bank_change", "bank_change_source", "explained", "residual", "residual_ok", "card_vs_bank_diff")}}
    if args.strict and not rc["residual_ok"]:
        digest["ok"] = False
        digest["error"] = "reconciliation residual %.2f >= 1" % rc["residual"]
        digest["hint"] = "re-run the identity after every rule change; each component is defined on types/flags (F12)"
        emit(digest)
        sys.exit(4)
    emit(digest)


if __name__ == "__main__":
    main()
