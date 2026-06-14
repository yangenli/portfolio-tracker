#!/usr/bin/env python
"""
Portfolio Project — performance tracker.

Reads every student portfolio file in ./submissions, pulls prices from Yahoo
Finance, computes each student's cumulative return and Sharpe ratio over the
tracking window, compares them to the S&P 500, and writes ./docs/results.json
for the static web page to load.

Run it again any time (e.g. every few days) to refresh the numbers:

    "C:/Users/yange/anaconda3/envs/simple/python.exe" compute_portfolios.py

Author note: prices are dividend/split adjusted (total return). The S&P 500
benchmark default is ^GSPC. All assumptions are collected in the CONFIG block
below so they are easy to change.
"""

import torch  # noqa: F401  (import first per machine convention; harmless if unused)

import os
import re
import csv
import glob
import json
import math
import datetime as dt

import numpy as np
import pandas as pd
import openpyxl
import yfinance as yf

# ──────────────────────────────────────────────────────────────────────────
# CONFIG  — edit these and re-run
# ──────────────────────────────────────────────────────────────────────────
HERE          = os.path.dirname(os.path.abspath(__file__))
SUBMISSIONS   = os.path.join(HERE, "submissions")
OUT_JSON      = os.path.join(HERE, "docs", "results.json")

TRACKING_START = "2025-06-22"   # portfolios "lock in" on this date
TRACKING_END   = "2025-08-06"   # end of the 8-week tracking window

BENCHMARK_TICKER = "^GSPC"      # what we call "the S&P 500" on the page
BENCHMARK_NAME   = "S&P 500"

RISK_FREE_ANNUAL = 0.0          # used in the Sharpe ratio; 0 = simple classroom version
TRADING_DAYS     = 252          # annualization factor for Sharpe

# Set False before publishing publicly (e.g. GitHub Pages): student names are
# stripped from results.json and the page falls back to Student IDs. The raw
# submission files (whose names are in the filenames) are git-ignored anyway.
PUBLISH_NAMES = False

# Obvious ticker typo fixes. Every correction is logged loudly so you can veto
# any of them. Leave the dict empty if you'd rather not auto-correct anything.
TICKER_CORRECTIONS = {
    "APPL":  "AAPL",   # Apple
    "PLNTR": "PLTR",   # Palantir
}

# If a student didn't put their 8-digit ID in the filename you can pin one here:
#   "<substring of filename>": "<student id>"
STUDENT_ID_OVERRIDES = {}

# ──────────────────────────────────────────────────────────────────────────
# Parsing student files
# ──────────────────────────────────────────────────────────────────────────

def _rows_from_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _rows_from_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return [list(r) for r in csv.reader(f)]


def _norm(s):
    return str(s).strip().lower() if s is not None else ""


def find_columns(rows):
    """Locate the header row and the ticker / shares column indices."""
    for i, row in enumerate(rows):
        cells = [_norm(c) for c in row]
        ticker_col = next((j for j, c in enumerate(cells) if "ticker" in c), None)
        if ticker_col is None:
            continue
        # Candidate "shares" columns: header mentions 'share' but not 'price'.
        share_cands = [j for j, c in enumerate(cells) if "share" in c and "price" not in c]
        if share_cands:
            # Prefer an exact "shares"/"share" header, else the first candidate.
            exact = [j for j in share_cands if _norm(rows[i][j]) in ("shares", "share")]
            shares_col = exact[0] if exact else share_cands[0]
        else:
            # Fallback: assume the column right after ticker holds the shares.
            shares_col = ticker_col + 1
        return i, ticker_col, shares_col
    return None, None, None


def normalize_ticker(raw):
    t = str(raw).strip().upper().replace(" ", "")
    t = t.replace(".", "-")            # BRK.B -> BRK-B  (Yahoo convention)
    return t


def parse_portfolio(path):
    """Return (holdings, notes) where holdings = {ticker: shares}."""
    rows = _rows_from_csv(path) if path.lower().endswith(".csv") else _rows_from_xlsx(path)
    header_i, tcol, scol = find_columns(rows)
    holdings, notes = {}, []
    if header_i is None:
        return holdings, ["could not find a 'Ticker' header"]

    for row in rows[header_i + 1:]:
        if tcol >= len(row):
            continue
        traw = row[tcol]
        if traw is None or str(traw).strip() == "":
            continue
        tnorm = normalize_ticker(traw)
        if not re.search(r"[A-Z]", tnorm):          # not ticker-like
            continue
        if "TOTAL" in tnorm:                         # a summary row
            continue
        # shares value
        sval = row[scol] if scol < len(row) else None
        try:
            shares = float(str(sval).replace(",", "").replace("$", "").strip())
        except (TypeError, ValueError):
            notes.append(f"skipped {tnorm}: shares not numeric ({sval!r})")
            continue
        if shares <= 0:
            continue
        # apply typo corrections
        if tnorm in TICKER_CORRECTIONS:
            fixed = TICKER_CORRECTIONS[tnorm]
            notes.append(f"corrected ticker {tnorm} -> {fixed}")
            tnorm = fixed
        holdings[tnorm] = holdings.get(tnorm, 0.0) + shares
    return holdings, notes


ID_RE = re.compile(r"0\d{7}")          # student IDs here are 8 digits starting with 0

def student_identity(filename):
    base = os.path.splitext(os.path.basename(filename))[0]
    for key, sid in STUDENT_ID_OVERRIDES.items():
        if key in filename:
            return sid, base.split("_")[0]
    m = ID_RE.search(base)
    name_slug = base.split("_")[0]
    sid = m.group(0) if m else name_slug
    return sid, name_slug


# ──────────────────────────────────────────────────────────────────────────
# Prices & metrics
# ──────────────────────────────────────────────────────────────────────────

def download_prices(tickers, start, end):
    """Return a DataFrame of adjusted close prices, columns = tickers."""
    # fetch a small buffer before start so the first tracking day is covered
    buf_start = (pd.Timestamp(start) - pd.Timedelta(days=7)).strftime("%Y-%m-%d")
    buf_end   = (pd.Timestamp(end) + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
    data = yf.download(sorted(tickers), start=buf_start, end=buf_end,
                       auto_adjust=True, progress=False, group_by="column")
    if isinstance(data.columns, pd.MultiIndex):
        close = data["Close"].copy()
    else:  # single ticker
        close = data[["Close"]].copy()
        close.columns = list(tickers)
    return close


def sharpe_ratio(daily_returns):
    r = np.asarray(daily_returns, dtype=float)
    r = r[~np.isnan(r)]
    if len(r) < 2 or r.std(ddof=1) == 0:
        return None
    rf_daily = RISK_FREE_ANNUAL / TRADING_DAYS
    excess = r - rf_daily
    return float(excess.mean() / r.std(ddof=1) * math.sqrt(TRADING_DAYS))


def series_metrics(value_series):
    """value_series: pd.Series indexed by date. Returns dict of metrics + curve."""
    v = value_series.dropna()
    daily = v.pct_change().dropna()
    cum_return = float(v.iloc[-1] / v.iloc[0] - 1)
    curve = [round(float(x), 6) for x in (v / v.iloc[0] - 1.0).values]  # fractional
    return {
        "cumulative_return": round(cum_return, 6),
        "sharpe": (round(sharpe_ratio(daily), 4) if sharpe_ratio(daily) is not None else None),
        "curve": curve,
        "start_value": round(float(v.iloc[0]), 2),
        "last_value": round(float(v.iloc[-1]), 2),
    }


# ──────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────

def main():
    files = sorted(glob.glob(os.path.join(SUBMISSIONS, "*.xlsx")) +
                   glob.glob(os.path.join(SUBMISSIONS, "*.csv")))
    print(f"Found {len(files)} submission files.\n")

    students = []     # list of dicts: id, name, file, holdings, notes
    all_tickers = set()
    for f in files:
        holdings, notes = parse_portfolio(f)
        sid, name = student_identity(f)
        students.append({"id": sid, "name": name, "file": os.path.basename(f),
                         "holdings": holdings, "notes": notes})
        all_tickers.update(holdings.keys())
        if notes:
            for n in notes:
                print(f"  [{name}] {n}")

    print(f"\n{len(all_tickers)} unique tickers across all portfolios.")

    # ---- prices ----
    close = download_prices(all_tickers | {BENCHMARK_TICKER}, TRACKING_START, TRACKING_END)

    # canonical trading-day calendar = benchmark days within the window
    bench = close[BENCHMARK_TICKER].dropna()
    bench = bench[(bench.index >= pd.Timestamp(TRACKING_START)) &
                  (bench.index <= pd.Timestamp(TRACKING_END))]
    calendar = bench.index
    if len(calendar) == 0:
        raise SystemExit("No benchmark trading days in the window — check the dates.")

    # align every price column to the calendar, fill small gaps
    prices = close.reindex(calendar).ffill().bfill()

    # which tickers actually have data?
    valid_tickers = {t for t in all_tickers if t in prices.columns and prices[t].notna().any()}
    bad_tickers = all_tickers - valid_tickers
    if bad_tickers:
        print(f"\nTickers with NO Yahoo data (excluded): {sorted(bad_tickers)}")

    dates = [d.strftime("%Y-%m-%d") for d in calendar]

    # ---- benchmark metrics ----
    bench_series = prices[BENCHMARK_TICKER]
    bm = series_metrics(bench_series)
    benchmark_out = {
        "name": BENCHMARK_NAME, "ticker": BENCHMARK_TICKER,
        "cumulative_return": bm["cumulative_return"], "sharpe": bm["sharpe"],
        "curve": bm["curve"],
    }
    print(f"\n{BENCHMARK_NAME} ({BENCHMARK_TICKER}): "
          f"return {bm['cumulative_return']*100:.2f}%  sharpe {bm['sharpe']}")

    # ---- per student ----
    BUDGET = 100_000.0
    out_students = {}
    internal_names = {}     # id -> real name, for the console report only
    duplicates = {}
    for s in students:
        valid = {t: sh for t, sh in s["holdings"].items() if t in valid_tickers}
        dropped = sorted(set(s["holdings"]) - set(valid))
        if not valid:
            print(f"  !! {s['name']} ({s['id']}): no valid tickers, skipping")
            continue

        # portfolio value series = sum(shares * price)
        value = sum(prices[t] * sh for t, sh in valid.items())
        m = series_metrics(value)
        start_total = m["start_value"]

        beats_return = m["cumulative_return"] > benchmark_out["cumulative_return"]
        beats_sharpe = (m["sharpe"] is not None and benchmark_out["sharpe"] is not None
                        and m["sharpe"] > benchmark_out["sharpe"])
        if beats_return and beats_sharpe:
            points = 5
        elif beats_return or beats_sharpe:
            points = 3
        else:
            points = None  # "to be confirmed by instructor"

        holdings_detail = []
        max_pos = 0.0
        for t, sh in sorted(valid.items()):
            p0 = float(prices[t].iloc[0])
            pos_val = p0 * sh
            max_pos = max(max_pos, pos_val)
            holdings_detail.append({
                "ticker": t, "shares": round(sh, 4),
                "start_price": round(p0, 4),
                "last_price": round(float(prices[t].iloc[-1]), 4),
                "start_value": round(pos_val, 2),
                "weight": round(pos_val / start_total, 4),
            })

        # ---- data-quality flags ----
        flags = []
        if dropped:
            flags.append(f"could not price: {', '.join(dropped)}")
        if max_pos > BUDGET:               # a single position bigger than the whole budget
            flags.append("a single position exceeds the $100k budget (bad share count or price)")
        if start_total > 3 * BUDGET or start_total < 0.4 * BUDGET:
            flags.append(f"start value ${start_total:,.0f} is far from $100k")
        # "corrupt" = weights are meaningless, keep OUT of the ranked leaderboard
        corrupt = (max_pos > BUDGET) or (start_total > 3 * BUDGET)

        internal_names[s["id"]] = s["name"]
        record = {
            "student_id": s["id"], "name": (s["name"] if PUBLISH_NAMES else None),
            "file": (s["file"] if PUBLISH_NAMES else None),
            "start_value": m["start_value"], "last_value": m["last_value"],
            "cumulative_return": m["cumulative_return"], "sharpe": m["sharpe"],
            "beats_return": beats_return, "beats_sharpe": beats_sharpe,
            "points": points,
            "curve": m["curve"],
            "holdings": holdings_detail,
            "dropped_tickers": dropped,
            "flags": flags,
            "needs_review": corrupt,
            "notes": s["notes"],
        }
        if s["id"] in out_students:
            duplicates.setdefault(s["id"], []).append(s["file"])
        out_students[s["id"]] = record

    if duplicates:
        print(f"\nDuplicate student IDs (last one kept): {duplicates}")

    # ---- leaderboard: rank trustworthy portfolios; list flagged ones separately ----
    rankable = [r for r in out_students.values() if not r["needs_review"]]
    ranking = sorted(rankable, key=lambda r: r["cumulative_return"], reverse=True)
    leaderboard = [{
        "rank": i + 1, "student_id": r["student_id"], "name": r["name"],
        "cumulative_return": r["cumulative_return"], "sharpe": r["sharpe"],
        "beats_return": r["beats_return"], "beats_sharpe": r["beats_sharpe"],
        "flags": r["flags"],
    } for i, r in enumerate(ranking)]
    needs_review = [{
        "student_id": r["student_id"], "name": r["name"],
        "cumulative_return": r["cumulative_return"], "flags": r["flags"],
    } for r in out_students.values() if r["needs_review"]]

    result = {
        "generated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "tracking_start": TRACKING_START,
        "tracking_end": TRACKING_END,
        "last_price_date": dates[-1],
        "dates": dates,
        "benchmark": benchmark_out,
        "leaderboard": leaderboard,
        "needs_review": needs_review,
        "students": out_students,
        "methodology": {
            "prices": "Yahoo Finance adjusted close (dividends & splits included)",
            "cumulative_return": "portfolio value on last day / value on first day - 1",
            "sharpe": f"mean(daily excess return)/std * sqrt({TRADING_DAYS}), "
                      f"risk-free {RISK_FREE_ANNUAL:.1%} annual",
            "benchmark": f"{BENCHMARK_NAME} ({BENCHMARK_TICKER})",
        },
    }

    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=1)
    print(f"\nWrote {OUT_JSON}  ({len(out_students)} students)")

    # tidy console leaderboard
    print("\nRank  Return   Sharpe  ID         Name")
    for r in leaderboard:
        badge = "  *" + "; ".join(r["flags"]) if r["flags"] else ""
        print(f"{r['rank']:>3}  {r['cumulative_return']*100:>7.2f}%  "
              f"{(r['sharpe'] if r['sharpe'] is not None else 0):>6.2f}  "
              f"{r['student_id']:<10} {internal_names.get(r['student_id'],'')}{badge}")
    if needs_review:
        print("\nNEEDS REVIEW (excluded from ranking):")
        for r in needs_review:
            print(f"  {r['student_id']:<10} {internal_names.get(r['student_id'],''):<16} "
                  f"{r['cumulative_return']*100:>8.2f}%  -- {'; '.join(r['flags'])}")
    if not PUBLISH_NAMES:
        print("\n(results.json is ANONYMIZED — names omitted. Set PUBLISH_NAMES=True to include them.)")


if __name__ == "__main__":
    main()
