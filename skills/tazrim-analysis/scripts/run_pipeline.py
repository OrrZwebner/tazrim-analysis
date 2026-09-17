#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""run_pipeline.py — run the tazrim-analysis stages in order, one JSON summary.

Purpose : orchestrate the sibling scripts per mode:
            A (first run / full)  : parse_bank, parse_card_cycle, parse_card_blocks, parse_benefits, classify,
                                    build_excel, add_transfers_sheet, verify, make_summary,
                                    make_figures, make_report_html, build_dashboard
            B (re-label)          : from classify
            C (add month / files) : from the parsers (same list as A)
          `--from <step>` / `--to <step>` narrow the range; `--dry-run` lists the plan only.
          Parsers whose config section is empty are skipped (no accounts -> parse_bank; no cycle-format
          cards -> parse_card_cycle; no per-card-statement cards -> parse_card_blocks; no benefit programs ->
          parse_benefits). Steps whose script does not exist yet are reported as "missing" and
          the run continues; a failing step stops the run.
Inputs  : tazrim.config.json (+ --level); each step is a subprocess with the same --config /
          --project-dir / --level and must print one JSON object.
Outputs : one JSON object on stdout {ok, mode, level, steps: [{name, status: ok|fail|missing|
          skipped|planned, seconds, result}], stopped_at, total_seconds}.
Exit    : 0 when no step failed; 1 when a step failed; 2 on a config error.

Python 3.8 compatible; stdlib only.
"""
import json
import os
import subprocess
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import emit, fail, log, parse_args  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
STEPS = ["parse_bank", "parse_card_cycle", "parse_card_blocks", "parse_benefits", "classify", "build_excel",
         "add_transfers_sheet", "verify", "make_summary", "make_figures", "make_report_html", "build_dashboard"]
MODE_START = {"A": "parse_bank", "B": "classify", "C": "parse_bank"}
CYCLE_ISSUERS, BLOCKS_ISSUERS = ("card_cycle_xlsx",), ("card_blocks_xls", "card_detail_xlsx")


def skip_reason(name, cfg):
    """Reason to skip a parser whose config section is empty; None otherwise."""
    cards = cfg.get("cards") or []
    if name == "parse_bank" and not cfg.get("accounts"):
        return "no accounts[] in config"
    if name == "parse_card_cycle" and not any(c.get("issuer") in CYCLE_ISSUERS for c in cards):
        return "no card_cycle_xlsx cards in config"
    if name == "parse_card_blocks" and not any(c.get("issuer") in BLOCKS_ISSUERS for c in cards):
        return "no card_blocks_xls / card_detail_xlsx cards in config"
    if name == "parse_benefits" and not cfg.get("benefit_programs"):
        return "no benefit_programs[] in config"
    return None


def run_step(name, cfg, level, timeout=1800):
    script = os.path.join(HERE, name + ".py")
    if not os.path.isfile(script):
        return {"name": name, "status": "missing", "seconds": 0.0, "result": {"error": "script not found: %s" % script}}
    cmd = [sys.executable, script, "--config", cfg.path, "--project-dir", cfg.project_dir, "--level", level]
    t0 = time.time()
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        return {"name": name, "status": "fail", "seconds": round(time.time() - t0, 1), "result": {"error": "timeout after %ds" % timeout}}
    last = [ln for ln in r.stdout.strip().splitlines() if ln.strip()]
    try:
        result = json.loads(last[-1]) if last else {}
    except ValueError:
        result = {"stdout_tail": r.stdout[-500:]}
    if r.stderr.strip():
        log("---- %s stderr (tail):\n%s" % (name, r.stderr.strip()[-1500:]))
    ok = r.returncode == 0 and result.get("ok", True) is not False
    if not ok and "error" not in result:
        result["error"] = "exit code %d" % r.returncode
    return {"name": name, "status": "ok" if ok else "fail", "seconds": round(time.time() - t0, 1), "result": result}


def plan(mode, start, stop):
    first = STEPS.index(start or MODE_START[mode])
    last = STEPS.index(stop) if stop else len(STEPS) - 1
    if last < first:
        fail("--to %r comes before --from %r" % (stop, start or MODE_START[mode]), "order: %s" % " -> ".join(STEPS))
    return STEPS[first:last + 1]


def main(argv=None):
    def extra(p):
        p.add_argument("--mode", choices=["A", "B", "C"], default="A", help="A full run, B from classify, C from the parsers")
        p.add_argument("--from", dest="start", choices=STEPS, default=None, help="first step")
        p.add_argument("--to", dest="stop", choices=STEPS, default=None, help="last step")
        p.add_argument("--dry-run", action="store_true", help="print the plan without running")
        p.add_argument("--continue-on-fail", action="store_true", help="do not stop at the first failing step")
    args, cfg = parse_args(__doc__.splitlines()[0], extra, argv)
    steps = plan(args.mode, args.start, args.stop)
    out, stopped, t0 = [], None, time.time()
    for name in steps:
        reason = skip_reason(name, cfg)
        if reason:
            out.append({"name": name, "status": "skipped", "seconds": 0.0, "result": {"reason": reason}})
            log("== %s: skipped (%s)" % (name, reason))
            continue
        if args.dry_run:
            exists = os.path.isfile(os.path.join(HERE, name + ".py"))
            out.append({"name": name, "status": "planned" if exists else "missing", "seconds": 0.0, "result": {}})
            continue
        st = run_step(name, cfg, cfg.level)
        out.append(st)
        log("== %s: %s (%.1fs)%s" % (name, st["status"], st["seconds"],
                                     (" — " + str(st["result"].get("error", ""))) if st["status"] in ("fail", "missing") else ""))
        if st["status"] == "fail" and not args.continue_on_fail:
            stopped = name
            break
    ok = not any(s["status"] == "fail" for s in out)
    emit({"ok": ok, "mode": args.mode, "level": cfg.level, "dry_run": args.dry_run, "steps": out,
          "stopped_at": stopped, "total_seconds": round(time.time() - t0, 1)})
    if not ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
