#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fx_rates.py — Bank of Israel representative rates with cache, API and manual fallback (F4).

Purpose : give parse_card_cycle.py / classify.py one rate lookup: cached SDMX CSV under `fx.rates_dir`
          -> BOI SDMX API (only when the cache is missing and fx.source == "boi_sdmx") ->
          `fx.manual_rates` CSV (date,ccy,rate) -> None. Never raises on network errors; every
          problem is recorded in `RateBook.problems` so the caller can report it.
          Also holds the generic country-word -> currency map used to infer an original currency.
Inputs  : cfg (common.Config) for fx.rates_dir / fx.source / fx.manual_rates and the window start.
Outputs : none on its own. `python3 fx_rates.py --config ...` prints a JSON status of the cache.
Exit    : 0; 2 on config error.
Python 3.8 compatible; stdlib only.
"""
import csv
import datetime as _dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import emit, log, parse_args, project_path  # noqa: E402
from formats import BOI  # noqa: E402

#: Generic country word (as it appears in issuer "פירוט" cells) -> ISO currency. Extend freely.
COUNTRY_CCY = {
    "ארצות הברית": "USD", "ארה\"ב": "USD", "קנדה": "CAD", "בריטניה": "GBP", "אנגליה": "GBP",
    "יוון": "EUR", "איטליה": "EUR", "ספרד": "EUR", "צרפת": "EUR", "גרמניה": "EUR", "הולנד": "EUR",
    "אוסטריה": "EUR", "פורטוגל": "EUR", "בלגיה": "EUR", "לוקסמבורג": "EUR", "אירלנד": "EUR",
    "קפריסין": "EUR", "פינלנד": "EUR", "סלובניה": "EUR", "קרואטיה": "EUR", "מלטה": "EUR",
    "שוויץ": "CHF", "שוודיה": "SEK", "נורבגיה": "NOK", "דנמרק": "DKK", "פולין": "PLN", "צ'כיה": "CZK",
    "הונגריה": "HUF", "טורקיה": "TRY", "תאילנד": "THB", "יפן": "JPY", "סין": "CNY", "הודו": "INR",
    "סינגפור": "SGD", "אוסטרליה": "AUD", "ניו זילנד": "NZD", "מקסיקו": "MXN", "ברזיל": "BRL",
    "דרום אפריקה": "ZAR", "איחוד האמירויות": "AED", "ירדן": "JOD", "מצרים": "EGP", "גאורגיה": "GEL",
    "וייטנאם": "VND", "אינדונזיה": "IDR", "נפאל": "NPR", "ישראל": "ILS",
}

_LOOKBACK = BOI.get("lookback_days", 10)


def country_in(text):
    # type: (str) -> str
    """First known country word found in `text` (longest first), else ''."""
    t = str(text or "")
    for name in sorted(COUNTRY_CCY, key=len, reverse=True):
        if name in t:
            return name
    return ""


class RateBook(object):
    """Lazy per-currency rate tables. `lookup(ccy, iso_date)` -> (rate, day, source) or None."""

    def __init__(self, cfg, allow_network=True):
        self.cfg = cfg
        self.rates_dir = project_path(cfg, "fx.rates_dir")
        self.source = (cfg.get("fx") or {}).get("source", "boi_sdmx")
        self.manual_path = (cfg.get("fx") or {}).get("manual_rates")
        self.allow_network = allow_network and self.source == "boi_sdmx"
        self.tables = {}      # ccy -> {iso_date: rate}
        self.manual = None    # ccy -> {iso_date: rate}
        self.problems = []    # human-readable, reported by the caller
        self.fetched = []     # currencies fetched from the API in this run

    # -- cache / api ---------------------------------------------------------------
    def _cache_path(self, ccy):
        return os.path.join(self.rates_dir, BOI["cache_file_template"].format(ccy=ccy))

    def _read_sdmx(self, path):
        out = {}
        with open(path, encoding="utf-8-sig", newline="") as fh:
            for r in csv.DictReader(fh):
                d, v = r.get(BOI["date_column"]), r.get(BOI["rate_column"])
                if d and v not in (None, ""):
                    try:
                        out[str(d)[:10]] = float(v)
                    except ValueError:
                        continue
        return out

    def _fetch(self, ccy):
        """Download the SDMX CSV into the cache. Returns True on success; never raises."""
        import urllib.request
        start = self.cfg["window"]["start"]
        end = _dt.date.today().isoformat()
        url = BOI["endpoint_template"].format(ccy=ccy, start=start, end=end)
        path = self._cache_path(ccy)
        try:
            os.makedirs(self.rates_dir, exist_ok=True)
            with urllib.request.urlopen(url, timeout=20) as resp:
                data = resp.read()
            if BOI["date_column"].encode("utf-8") not in data:
                self.problems.append("BOI %s: unexpected response (no %s column); endpoint may have changed"
                                     % (ccy, BOI["date_column"]))
                return False
            with open(path, "wb") as fh:
                fh.write(data)
            self.fetched.append(ccy)
            log("fetched BOI rates for %s -> %s" % (ccy, path))
            return True
        except Exception as e:  # noqa: BLE001 — network errors must never fail the parse
            self.problems.append("BOI %s: fetch failed (%s); using cache/manual rates" % (ccy, e))
            return False

    def _load_manual(self):
        if self.manual is not None:
            return
        self.manual = {}
        if not self.manual_path:
            return
        path = project_path(self.cfg, self.manual_path)
        if not os.path.isfile(path):
            self.problems.append("fx.manual_rates file not found: %s" % path)
            return
        with open(path, encoding="utf-8-sig", newline="") as fh:
            for r in csv.DictReader(fh):
                try:
                    self.manual.setdefault(str(r["ccy"]).strip().upper(), {})[str(r["date"]).strip()[:10]] = float(r["rate"])
                except (KeyError, ValueError):
                    self.problems.append("manual rates: bad row %r (need date,ccy,rate)" % r)

    def table(self, ccy):
        # type: (str) -> dict
        if ccy in self.tables:
            return self.tables[ccy]
        path = self._cache_path(ccy)
        if not os.path.isfile(path) and self.allow_network:
            self._fetch(ccy)
        t = {}
        if os.path.isfile(path):
            try:
                t = self._read_sdmx(path)
            except Exception as e:  # noqa: BLE001
                self.problems.append("rate cache %s unreadable: %s" % (path, e))
        if not t:
            self.problems.append("no BOI rates available for %s" % ccy)
        self.tables[ccy] = t
        return t

    # -- lookup --------------------------------------------------------------------
    @staticmethod
    def _back(table, iso_date, lookback):
        try:
            d = _dt.date.fromisoformat(str(iso_date)[:10])
        except ValueError:
            return None
        for _ in range(lookback + 1):
            k = d.isoformat()
            if k in table:
                return table[k], k
            d -= _dt.timedelta(days=1)
        return None

    def lookup(self, ccy, iso_date):
        # type: (str, str) -> object
        """(rate, rate_day, source) with source 'boi' | 'manual', or None when no rate is known
        within `lookback_days` before `iso_date`."""
        hit = self._back(self.table(ccy), iso_date, _LOOKBACK)
        if hit:
            return hit[0], hit[1], "boi"
        self._load_manual()
        hit = self._back(self.manual.get(ccy, {}), iso_date, _LOOKBACK)
        if hit:
            return hit[0], hit[1], "manual"
        self.problems.append("no rate for %s on %s (cache, API and manual all missing)" % (ccy, iso_date))
        return None


def main(argv=None):
    args, cfg = parse_args("report the state of the BOI rate cache", argv=argv)
    book = RateBook(cfg, allow_network=False)
    ccys = sorted({(c.get("fx_settlement") or {}).get("currency") for c in cfg["cards"]} - {None})
    status = {}
    for ccy in ccys:
        t = book.table(ccy)
        status[ccy] = {"cached_days": len(t), "first": min(t) if t else None, "last": max(t) if t else None}
    emit({"ok": True, "rates_dir": book.rates_dir, "currencies": status, "problems": book.problems})


if __name__ == "__main__":
    main()
