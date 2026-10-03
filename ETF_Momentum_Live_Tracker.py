#!/usr/bin/env python3
# =============================================================================
# ETF Momentum Rotation — LIVE HOLDINGS TRACKER
# Single self-contained script. Paste into a GitHub repo and run: python ETF_Momentum_Live_Tracker.py
# It reads YOUR current holdings (etf_holdings.csv) and, using today's prices,
# tells you which ETFs are STILL ON HOLD and which have triggered a SELL, per the
# strategy's exact exit rules:
#     EXIT 1  Hard stop   : price <= entry * (1 - 15%)
#     EXIT 2  Trail stop  : price <= peak-since-entry * (1 - 20%)
#     EXIT 3  Rotation    : ETF dropped out of the top-6 by momentum rank
#     else                : HOLD
# It also lists the current top-6 buy candidates (above their 200-DMA) to rotate
# freed slots into. Outputs one Excel workbook + one self-contained dark HTML report.
# (This is the LIVE tracker — not the backtest and not the fresh-start screener.)
# OPTIONAL: when run with SEND_TT_SIGNALS=1 and a TRADETRON_TOKEN (GitHub secret), it also
# sends today's ranks + regime to the Tradetron strategy 'Nifty Gold and ETF Momentum 1L'.
# Without those two settings it behaves exactly as before and sends nothing.
# =============================================================================

import subprocess, sys

def _ensure(pkg, import_as=None):
    try:
        __import__(import_as or pkg)
    except ImportError:
        print(f"Installing {pkg}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "--quiet"])

for _p, _i in [("yfinance", "yfinance"), ("pandas", "pandas"), ("numpy", "numpy"),
               ("openpyxl", "openpyxl"), ("plotly", "plotly")]:
    _ensure(_p, _i)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import warnings
warnings.filterwarnings("ignore")

import os, math, json, time, webbrowser
import urllib.request, urllib.error
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.offline as pyo

# =============================================================================
# USER INPUTS
# =============================================================================
HOLDINGS_FILE   = "etf_holdings.csv"   # your current holdings (edit this file, not the code)
INITIAL_CAPITAL = 1_000_000            # used only to suggest per-slot sizing for new buys
AS_OF           = ""                   # "" = latest; or "YYYY-MM-DD" to check status as of a past date

# =============================================================================
# STRATEGY PARAMETERS (match the backtest exactly)
# =============================================================================
N_HOLD      = 6                        # buy only from the top 6 by rank
HOLD_RANK   = 10                       # but keep a holding until its rank falls past 10 (lower churn)
LB          = [21, 63, 126, 252]
WEIGHTS     = [0.15, 0.40, 0.30, 0.15]
DMA_PERIOD  = 200
MIN_HISTORY = 262
SL_PCT      = 0.15      # hard stop: sell if price <= entry * 0.85
TRAIL_PCT   = 0.20      # trail stop: sell if price <= peak * 0.80
BENCH       = "^CRSLDX"
BENCH_LABEL = "Nifty 500"

# Market-regime MASTER switch: buy/hold ETFs only when this index is above its
# DMA; otherwise sell every ETF and sit 100% in cash (no buys, no gold).
REGIME_INDEX = "^NSEI"     # Nifty 50
REGIME_LABEL = "Nifty 50"
REGIME_DMA   = 200

# Peak-since-entry basis for the trailing stop:
#   "daily"   = highest daily close since your entry (a true live trailing high) [default]
#   "monthly" = highest month-start close since entry (matches the monthly backtest exactly)
PEAK_BASIS  = "daily"

C54 = [
    ("NIFTYBEES", "Broad_Equity"), ("JUNIORBEES", "Broad_Equity"), ("MID150BEES", "Broad_Equity"),
    ("MIDSELIETF", "Broad_Equity"), ("MONIFTY500", "Broad_Equity"), ("HDFCSML250", "Broad_Equity"),
    ("MOM100", "Broad_Equity"),
    ("BANKBEES", "Banking_Finance"), ("PSUBNKBEES", "Banking_Finance"), ("PVTBANIETF", "Banking_Finance"),
    ("FINIETF", "Banking_Finance"),
    ("ITBEES", "Sector"), ("PHARMABEES", "Sector"), ("AUTOBEES", "Sector"), ("CONSUMBEES", "Sector"),
    ("CONSUMER", "Sector"), ("FMCGIETF", "Sector"), ("CHEMICAL", "Sector"), ("METALIETF", "Sector"),
    ("OILIETF", "Sector"), ("COMMOIETF", "Sector"), ("INFRAIETF", "Sector"), ("CPSEETF", "Sector"),
    ("ICICIB22", "Sector"), ("MOCAPITAL", "Sector"), ("MODEFENCE", "Sector"), ("MOREALTY", "Sector"),
    ("MOTOUR", "Sector"), ("GROWWPOWER", "Sector"), ("GROWWRAIL", "Sector"), ("GROWWHOSPI", "Sector"),
    ("DIVOPPBEES", "Thematic"), ("EVINDIA", "Thematic"), ("INTERNET", "Thematic"), ("MNC", "Thematic"),
    ("SELECTIPO", "Thematic"), ("TOP10ADD", "Thematic"),
    ("LOWVOLIETF", "Factor"), ("NV20IETF", "Factor"), ("QUAL30IETF", "Factor"), ("NIFTYQLITY", "Factor"),
    ("HDFCGROWTH", "Factor"), ("MOM50", "Factor"), ("MOMENTUM50", "Factor"),
    ("MOVALUE", "Factor"),
    ("GOLDBEES", "Gold_Silver"), ("SILVERBEES", "Gold_Silver"),
    ("LTGILTBEES", "Bonds"), ("EBBETF0430", "Bonds"), ("GILT5YBEES", "Bonds"),
]
CATEGORY = {s: c for s, c in C54}

OUTPUT_DIR = Path(__file__).parent / "ETF_Momentum_Tracker_Output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
EXCEL_PATH = OUTPUT_DIR / "ETF_Momentum_Tracker_Results.xlsx"
HTML_PATH  = OUTPUT_DIR / "ETF_Momentum_Tracker_Report.html"


# =============================================================================
# DATA
# =============================================================================
def load_prices():
    if os.getenv("SCREENER_MOCK") == "1":
        return _mock_prices()
    import yfinance as yf
    syms = [s for s, _ in C54]
    ytk = [s + ".NS" for s in syms] + [BENCH, REGIME_INDEX]
    print(f"Downloading data for {len(syms)} ETFs + benchmark...")
    raw = yf.download(ytk, period="max", interval="1d", auto_adjust=True,
                      group_by="ticker", threads=True, progress=False)

    def close_of(tkr):
        try:
            if isinstance(raw.columns, pd.MultiIndex):
                if tkr in raw.columns.get_level_values(0):
                    s = raw[tkr]["Close"].dropna()
                    return s if not s.empty else None
                return None
            return raw["Close"].dropna()
        except Exception:
            return None

    prices, dropped = {}, 0
    for s in syms:
        c = close_of(s + ".NS")
        if c is None or len(c) < MIN_HISTORY:
            dropped += 1
            continue
        c.index = pd.to_datetime(c.index).tz_localize(None)
        prices[s] = c.sort_index()
    bench = close_of(BENCH)
    if bench is not None:
        bench.index = pd.to_datetime(bench.index).tz_localize(None)
        bench = bench.sort_index()
    regime = close_of(REGIME_INDEX)
    if regime is not None:
        regime.index = pd.to_datetime(regime.index).tz_localize(None)
        regime = regime.sort_index()
    print(f"{len(prices)} ETFs loaded, {dropped} dropped (insufficient history)")
    return prices, bench, regime


def _mock_prices():
    rng = np.random.default_rng(11)
    idx = pd.bdate_range(end=datetime.now().date(), periods=800)
    prices = {}
    for i, (s, _) in enumerate(C54):
        drift = 0.0003 + (i % 9) * 0.00007
        prices[s] = pd.Series(100 * np.exp(np.cumsum(rng.normal(drift, 0.013, len(idx)))), index=idx)
    bench = pd.Series(100 * np.exp(np.cumsum(rng.normal(0.0004, 0.009, len(idx)))), index=idx)
    regime = pd.Series(100 * np.exp(np.cumsum(rng.normal(0.0004, 0.010, len(idx)))), index=idx)
    print(f"[MOCK] generated {len(prices)} ETFs + benchmark + regime index")
    return prices, bench, regime


# =============================================================================
# RANKS / SCORES  (point-in-time)
# =============================================================================
def _ret(series, lb):
    if len(series) <= lb:
        return None
    past = float(series.iloc[-lb])
    return None if past <= 0 else (float(series.iloc[-1]) / past - 1.0) * 100.0


def compute_ranks(prices, asof):
    rows = []
    for s, series in prices.items():
        ser = series.loc[:asof] if asof is not None else series
        if len(ser) < MIN_HISTORY:
            continue
        r = [_ret(ser, lb) for lb in LB]
        if any(v is None for v in r):
            continue
        rows.append({
            "symbol": s, "category": CATEGORY.get(s, ""),
            "score": sum(w * v for w, v in zip(WEIGHTS, r)),
            "ret_1m": r[0], "ret_3m": r[1], "ret_6m": r[2], "ret_12m": r[3],
            "price": float(ser.iloc[-1]), "dma200": float(ser.iloc[-DMA_PERIOD:].mean()),
        })
    df = pd.DataFrame(rows).sort_values("score", ascending=False).reset_index(drop=True)
    df["rank"] = np.arange(1, len(df) + 1)
    df["above_dma"] = df["price"] > df["dma200"]
    df["top6"] = df["rank"] <= N_HOLD
    return df


# =============================================================================
# HOLDINGS
# =============================================================================
def read_holdings():
    if not os.path.exists(HOLDINGS_FILE):
        return pd.DataFrame(columns=["symbol", "entry_date", "entry_price", "qty"])
    df = pd.read_csv(HOLDINGS_FILE, comment="#", skip_blank_lines=True)
    if df.empty:
        return df
    df["symbol"] = df["symbol"].astype(str).str.strip().str.upper()
    df["entry_date"] = pd.to_datetime(df["entry_date"], errors="coerce")
    df["entry_price"] = pd.to_numeric(df["entry_price"], errors="coerce")
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce")
    return df.dropna(subset=["symbol", "entry_price", "qty"])


def _peak_since(series, entry_date, asof):
    ser = series.loc[:asof] if asof is not None else series
    if entry_date is not None and not pd.isna(entry_date):
        ser = ser.loc[ser.index >= entry_date]
    if ser.empty:
        return None
    if PEAK_BASIS == "monthly":
        m = ser.groupby(ser.index.to_period("M")).first()   # month-start-ish close
        return float(m.max())
    return float(ser.max())


def evaluate(holdings, prices, ranks, asof, regime_on=True):
    rank_of = {r["symbol"]: r for _, r in ranks.iterrows()}
    out = []
    for _, h in holdings.iterrows():
        sym = h["symbol"]
        series = prices.get(sym)
        if series is None:
            out.append({**h.to_dict(), "status": "NO DATA", "reason": "no price for symbol",
                        "current": None, "peak": None, "pnl": None, "rank": None,
                        "sl_level": None, "trail_level": None, "sl_cushion": None,
                        "trail_cushion": None, "value": None})
            continue
        ser = series.loc[:asof] if asof is not None else series
        cur = float(ser.iloc[-1])
        peak = _peak_since(series, h["entry_date"], asof) or cur
        entry = float(h["entry_price"])
        sl_level = entry * (1 - SL_PCT)
        trail_level = peak * (1 - TRAIL_PCT)
        rinfo = rank_of.get(sym)
        rk = int(rinfo["rank"]) if rinfo is not None else None
        if regime_on is False:
            status, reason = "SELL", f"Regime OFF — {REGIME_LABEL} below its {REGIME_DMA}-DMA (go to cash)"
        elif cur <= sl_level:
            status, reason = "SELL", f"Hard SL — down {(cur/entry-1)*100:.1f}% from entry (limit -{int(SL_PCT*100)}%)"
        elif cur <= trail_level:
            status, reason = "SELL", f"Trail stop — down {(cur/peak-1)*100:.1f}% from peak (limit -{int(TRAIL_PCT*100)}%)"
        elif rk is None or rk > HOLD_RANK:
            status, reason = "SELL", (f"Rotation — rank {rk} (fallen past {HOLD_RANK})" if rk
                                      else f"Rotation — no longer ranked in top {HOLD_RANK}")
        else:
            status, reason = "HOLD", f"rank {rk} (<= {HOLD_RANK}), stops intact"
        out.append({
            "symbol": sym, "category": CATEGORY.get(sym, ""),
            "entry_date": (h["entry_date"].date().isoformat() if not pd.isna(h["entry_date"]) else "-"),
            "entry_price": entry, "qty": int(h["qty"]), "current": cur, "peak": peak,
            "pnl": (cur / entry - 1) * 100, "value": cur * int(h["qty"]),
            "rank": rk, "in_top6": (rk is not None and rk <= N_HOLD),
            "sl_level": sl_level, "trail_level": trail_level,
            "sl_cushion": (cur / sl_level - 1) * 100, "trail_cushion": (cur / trail_level - 1) * 100,
            "status": status, "reason": reason,
        })
    return pd.DataFrame(out)


# =============================================================================
# EXCEL
# =============================================================================
def build_excel(hold_df, ranks, buys, meta, path):
    from openpyxl.styles import Font, PatternFill, Alignment
    green = PatternFill("solid", fgColor="C6EFCE"); red = PatternFill("solid", fgColor="FFC7CE")
    hdr = Font(bold=True, color="FFFFFF"); hdrfill = PatternFill("solid", fgColor="1F3B57")

    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        _reg = ("ON — invested" if meta.get("regime_on") is True
                else "OFF — all cash" if meta.get("regime_on") is False else "unknown")
        summ = pd.DataFrame({
            "Metric": ["As of", f"Regime ({REGIME_LABEL} vs {REGIME_DMA}-DMA)", "Holdings",
                       "Still on HOLD", "Flagged to SELL", "Portfolio value (₹)", "Total P&L %",
                       "Buy candidates (top-6, 3M>0, above DMA, not held)", "Generated"],
            "Value": [meta["as_of"], _reg, meta["n_hold_total"], meta["n_hold"], meta["n_sell"],
                      round(meta["value"]), (round(meta["pnl"], 2) if meta["pnl"] is not None else "-"),
                      meta["n_buy"], meta["generated"]],
        })
        summ.to_excel(xl, sheet_name="Summary", index=False)

        hc = ["symbol", "category", "status", "reason", "entry_date", "entry_price", "qty",
              "current", "pnl", "peak", "sl_level", "trail_level", "sl_cushion", "trail_cushion",
              "rank", "value"]
        hd = (hold_df[hc].copy() if len(hold_df) else pd.DataFrame(columns=hc))
        for c in ["entry_price", "current", "pnl", "peak", "sl_level", "trail_level",
                  "sl_cushion", "trail_cushion", "value"]:
            if c in hd:
                hd[c] = pd.to_numeric(hd[c], errors="coerce").round(2)
        hd.to_excel(xl, sheet_name="Holdings_Status", index=False)

        bc = buys[["rank", "symbol", "category", "score", "price", "dma200"]].copy() if len(buys) else \
            pd.DataFrame(columns=["rank", "symbol", "category", "score", "price", "dma200"])
        if len(bc):
            bc["suggested_qty"] = (meta["slot_size"] / bc["price"]).apply(math.floor)
            for c in ["score", "price", "dma200"]:
                bc[c] = bc[c].round(2)
        bc.to_excel(xl, sheet_name="Buy_Candidates", index=False)

        rc = ["rank", "symbol", "category", "score", "ret_1m", "ret_3m", "ret_6m", "ret_12m",
              "price", "dma200", "above_dma", "top6"]
        rr = ranks[rc].copy()
        for c in ["score", "ret_1m", "ret_3m", "ret_6m", "ret_12m", "price", "dma200"]:
            rr[c] = rr[c].round(2)
        rr.to_excel(xl, sheet_name="Rankings", index=False)

        for sh in xl.book.worksheets:
            for cell in sh[1]:
                cell.font = hdr; cell.fill = hdrfill; cell.alignment = Alignment(horizontal="center")
            for col in sh.columns:
                w = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                sh.column_dimensions[col[0].column_letter].width = min(max(w + 2, 10), 42)
        ws = xl.book["Holdings_Status"]
        if len(hd):
            for r in range(2, ws.max_row + 1):
                st = ws.cell(r, hc.index("status") + 1).value
                fill = green if st == "HOLD" else (red if st == "SELL" else None)
                if fill:
                    for c in range(1, len(hc) + 1):
                        ws.cell(r, c).fill = fill


# =============================================================================
# HTML
# =============================================================================
DARK = dict(paper_bgcolor="#0D1B2A", plot_bgcolor="#0D1B2A",
            font=dict(color="#E9EEF7"), margin=dict(l=40, r=20, t=40, b=40))
_first = [True]


def _div(fig):
    if _first[0]:
        _first[0] = False
        return pyo.plot(fig, include_plotlyjs=True, output_type="div")
    return pyo.plot(fig, include_plotlyjs=False, output_type="div")


def build_html(hold_df, ranks, buys, meta, path):
    acc, pos, neg, mut = "#F5A524", "#37D3A6", "#E5595E", "#8B98AD"

    # charts: holdings P&L, and universe DMA gate
    charts = ""
    if len(hold_df) and hold_df["current"].notna().any():
        hv = hold_df[hold_df["current"].notna()]
        f1 = go.Figure(go.Bar(
            x=hv["symbol"], y=hv["pnl"].round(1),
            marker_color=[pos if s == "HOLD" else neg for s in hv["status"]]))
        f1.update_layout(title="Your holdings — P&L % (green = hold, red = sell)", **DARK)
        charts += _div(f1)
    ab = int(ranks["above_dma"].sum()); be = len(ranks) - ab
    f2 = go.Figure(go.Pie(labels=["Above 200-DMA", "Below 200-DMA"], values=[ab, be], hole=0.45,
                          marker=dict(colors=[pos, neg])))
    f2.update_layout(title=f"200-DMA gate across the {len(ranks)} ETFs", **DARK)
    charts += _div(f2)

    pnl_txt = f'{meta["pnl"]:+.1f}%' if meta["pnl"] is not None else "-"
    kpis = [("As of", meta["as_of"]), ("Holdings", meta["n_hold_total"]),
            ("Still on HOLD", meta["n_hold"]), ("Flagged SELL", meta["n_sell"]),
            ("Portfolio ₹", f'{meta["value"]:,.0f}'), ("Total P&L", pnl_txt)]
    kpi_html = "".join(f"<div class='kpi'><div class='kv'>{v}</div><div class='kl'>{k}</div></div>"
                       for k, v in kpis)

    def hcard(r, kind):
        cls = "hold" if kind == "HOLD" else "sell"
        pnlc = "pos" if (r["pnl"] or 0) >= 0 else "neg"
        return (f"<div class='hc {cls}'><div class='hs'>{r['symbol']} "
                f"<span class='hcat'>{r['category']}</span></div>"
                f"<div class='hp {pnlc}'>{r['pnl']:+.1f}%</div>"
                f"<div class='hm'>{r['reason']}</div>"
                f"<div class='hm2'>now ₹{r['current']:,.2f} · entry ₹{r['entry_price']:,.2f} · "
                f"SL ₹{r['sl_level']:,.2f} ({r['sl_cushion']:+.1f}%) · "
                f"trail ₹{r['trail_level']:,.2f} ({r['trail_cushion']:+.1f}%)</div></div>")

    holds = hold_df[hold_df["status"] == "HOLD"] if len(hold_df) else hold_df
    sells = hold_df[hold_df["status"] == "SELL"] if len(hold_df) else hold_df
    nodata = hold_df[hold_df["status"] == "NO DATA"] if len(hold_df) else hold_df

    if len(hold_df) == 0:
        hold_block = ("<div class='empty'>No holdings in <code>etf_holdings.csv</code> yet — you're "
                      "starting fresh. Buy from the candidates below and add each to the file.</div>")
        sell_block = ""
    else:
        hold_block = ("<div class='cards'>" + "".join(hcard(r, "HOLD") for _, r in holds.iterrows()) + "</div>") \
            if len(holds) else "<div class='empty'>None of your holdings are still on hold.</div>"
        sell_block = ("<div class='cards'>" + "".join(hcard(r, "SELL") for _, r in sells.iterrows()) + "</div>") \
            if len(sells) else "<div class='empty'>Nothing flagged to sell — all holdings intact.</div>"
        if len(nodata):
            sell_block += ("<div class='empty'>No price data for: "
                           + ", ".join(nodata["symbol"]) + " — check the symbol.</div>")

    if len(buys):
        buy_block = "<div class='cards'>" + "".join(
            f"<div class='hc buy'><div class='hs'>{r.symbol} <span class='hcat'>{r.category}</span></div>"
            f"<div class='hm'>rank {int(r.rank)} · score {r.score:.1f} · "
            f"{(r.price/r.dma200-1)*100:+.1f}% above 200-DMA</div>"
            f"<div class='hm2'>~{math.floor(meta['slot_size']/r.price):,} units at ₹{r.price:,.2f} "
            f"(≈₹{meta['slot_size']:,.0f}/slot)</div></div>"
            for r in buys.itertuples(index=False)) + "</div>"
    else:
        buy_block = ("<div class='empty'>No buy candidates — none of the top-6 not held pass the filters "
                     "(positive 3-month return and above their 200-DMA), so those slots stay in cash.</div>")

    # full ranking table
    def cell(v, pct=False):
        if pct:
            return f"<td class='num {'pos' if v>=0 else 'neg'}'>{v:+.1f}%</td>"
        return f"<td class='num'>{v:,.2f}</td>"
    held_syms = set(hold_df["symbol"]) if len(hold_df) else set()
    trows = ""
    for r in ranks.itertuples(index=False):
        held = r.symbol in held_syms
        tags = ""
        if r.top6: tags += "<span class='tag top'>TOP6</span>"
        if held: tags += "<span class='tag held'>HELD</span>"
        dma = "<span class='ok'>▲</span>" if r.above_dma else "<span class='no'>▼</span>"
        trows += (f"<tr class='{'heldrow' if held else ('toprow' if r.top6 else '')}'>"
                  f"<td class='num'>{r.rank}</td><td class='sym'>{r.symbol}</td>"
                  f"<td class='muted'>{r.category}</td><td class='num'>{r.score:.1f}</td>"
                  + cell(r.ret_1m, True) + cell(r.ret_3m, True) + cell(r.ret_6m, True) + cell(r.ret_12m, True)
                  + cell(r.price) + f"<td class='num'>{dma}</td><td>{tags}</td></tr>")

    css = """
    :root{--bg:#0D1B2A;--panel:#12263A;--line:#24384f;--fg:#E9EEF7;--fg2:#8B98AD;--acc:#F5A524;--pos:#37D3A6;--neg:#E5595E}
    *{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font-family:Inter,system-ui,Arial,sans-serif}
    .wrap{max-width:1180px;margin:0 auto;padding:26px 20px 70px}
    h1{font-size:20px;margin:0 0 2px}.sub{color:var(--fg2);font-size:13px;margin-bottom:16px}
    h2{font-size:13px;text-transform:uppercase;letter-spacing:.06em;color:var(--fg2);margin:26px 0 10px}
    .kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:12px}
    .kpi{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:14px}
    .kpi .kv{font-size:20px;font-weight:600}.kpi .kl{color:var(--fg2);font-size:11px;text-transform:uppercase;letter-spacing:.05em;margin-top:4px}
    .cards{display:flex;gap:10px;flex-wrap:wrap}
    .hc{border-radius:12px;padding:12px 14px;min-width:230px;flex:1 1 230px;max-width:320px;border:1px solid var(--line);background:var(--panel)}
    .hc.hold{border-color:rgba(55,211,166,.5);background:linear-gradient(160deg,rgba(55,211,166,.12),rgba(55,211,166,.02))}
    .hc.sell{border-color:rgba(229,89,94,.5);background:linear-gradient(160deg,rgba(229,89,94,.12),rgba(229,89,94,.02))}
    .hc.buy{border-color:rgba(245,165,36,.5);background:linear-gradient(160deg,rgba(245,165,36,.12),rgba(245,165,36,.03))}
    .hc .hs{font-weight:700;font-size:15px}.hcat{font-weight:400;color:var(--fg2);font-size:11px}
    .hc .hp{font-family:ui-monospace,monospace;font-size:18px;margin:2px 0 4px}
    .hc .hm{font-size:12px;color:var(--fg)}.hc .hm2{font-size:11px;color:var(--fg2);margin-top:6px;line-height:1.5}
    .empty{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:16px;color:var(--fg2)}
    .tw{overflow-x:auto;border:1px solid var(--line);border-radius:12px;margin-top:6px}
    table{border-collapse:collapse;width:100%;min-width:880px;background:var(--panel)}
    th{position:sticky;top:0;background:#183049;color:var(--fg2);font-size:10.5px;text-transform:uppercase;letter-spacing:.04em;padding:9px 10px;text-align:right;border-bottom:1px solid var(--line);white-space:nowrap;cursor:pointer}
    th:nth-child(2),th:nth-child(3){text-align:left}
    td{padding:8px 10px;border-bottom:1px solid rgba(36,56,79,.6);text-align:right;font-size:12.5px;white-space:nowrap}
    td.sym{text-align:left;font-weight:600}td:nth-child(3){text-align:left}.muted{color:var(--fg2)}
    .num{font-variant-numeric:tabular-nums}.pos{color:var(--pos)}.neg{color:var(--neg)}
    .heldrow td{background:rgba(245,165,36,.08)}.toprow td{background:rgba(58,80,107,.16)}
    .ok{color:var(--pos)}.no{color:var(--neg)}
    .tag{font-size:9.5px;font-weight:700;padding:2px 6px;border-radius:4px;margin-left:3px}
    .tag.top{background:#3A506B;color:#E9EEF7}.tag.held{background:var(--acc);color:#0D1B2A}
    .rules{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:14px 16px;color:var(--fg2);font-size:12.5px;line-height:1.7}
    .note{color:var(--fg2);font-size:12px;margin-top:14px;line-height:1.55}
    .chart{background:var(--panel);border:1px solid var(--line);border-radius:12px;margin:12px 0;padding:6px}
    .regime{border-radius:12px;padding:13px 16px;margin:14px 0 6px;font-size:13px;border:1px solid var(--line);background:var(--panel);line-height:1.5}
    .regime .sm{display:block;font-family:ui-monospace,monospace;font-size:11px;color:var(--fg2);margin-top:5px}
    .regime.on{border-color:rgba(55,211,166,.5);background:linear-gradient(160deg,rgba(55,211,166,.12),rgba(55,211,166,.02))}.regime.on b{color:var(--pos)}
    .regime.off{border-color:rgba(229,89,94,.55);background:linear-gradient(160deg,rgba(229,89,94,.16),rgba(229,89,94,.03))}.regime.off b{color:var(--neg)}
    .regime.unk b{color:var(--acc)}
    """

    header = (f"<h1>ETF Momentum Rotation — Live Holdings Tracker</h1>"
              f"<div class='sub'>As of {meta['as_of']} · {meta['n_hold_total']} holdings · "
              f"{meta['n_hold']} still on hold · {meta['n_sell']} flagged to sell · benchmark {BENCH_LABEL}</div>")

    rules = (f"<div class='rules'><b>Regime gate (checked first):</b> hold or buy ETFs only when {REGIME_LABEL} "
             f"is above its {REGIME_DMA}-day average — otherwise every ETF is sold and you sit 100% in cash. "
             f"When invested, a holding is <b>kept</b> until one of these exits fires (priority order): "
             f"<b>Hard SL</b> (real-time) — price ≤ entry × (1−{int(SL_PCT*100)}%); <b>Trail stop</b> "
             f"(real-time) — price ≤ peak-since-entry × (1−{int(TRAIL_PCT*100)}%); <b>Rotation</b> — its rank "
             f"falls past {HOLD_RANK} (rank &gt; {HOLD_RANK}). Momentum score = 0.15·1M + 0.40·3M + 0.30·6M + "
             f"0.15·12M returns. <b>New buys</b> come from the top {N_HOLD} not currently held that are above "
             f"their {DMA_PERIOD}-day average <b>and have a positive 3-month return</b> (bought at month-end; "
             f"kept down to rank {HOLD_RANK}). Peak-since-entry basis: <b>{PEAK_BASIS}</b>.</div>")

    note = ("<div class='note'><b>Live monitor.</b> Status uses today's prices. The <b>hard and trailing stops "
            "are real-time</b> — act on a stop-driven SELL the day it appears; <b>new buys and rotation</b> "
            "(rank-based) happen at month-start. Prices are Yahoo Finance auto-adjusted closes; peak-since-entry "
            "is computed from price history (switch PEAK_BASIS for a different trailing basis). Not investment advice.</div>")

    ro = meta.get("regime_on")
    rc, rd = meta.get("regime_cur"), meta.get("regime_dma")
    rma = (f"{REGIME_LABEL} {rc:,.0f} vs {REGIME_DMA}-DMA {rd:,.0f}"
           if rc is not None and rd is not None else "")
    if ro is True:
        regime_banner = (f"<div class='regime on'><b>Regime ON — invested</b> · {REGIME_LABEL} is above its "
                         f"{REGIME_DMA}-day average, so the strategy is active (hold / buy the top {N_HOLD})."
                         f"<span class='sm'>{rma}</span></div>")
    elif ro is False:
        regime_banner = (f"<div class='regime off'><b>Regime OFF — all cash</b> · {REGIME_LABEL} is at/below its "
                         f"{REGIME_DMA}-day average. Every ETF is flagged SELL and you sit 100% in cash until it "
                         f"climbs back above.<span class='sm'>{rma}</span></div>")
    else:
        regime_banner = (f"<div class='regime unk'><b>Regime unknown</b> · not enough {REGIME_LABEL} "
                         f"history to compute the {REGIME_DMA}-DMA.</div>")

    tbl = (f"<div class='tw'><table id='t'><thead><tr>"
           f"<th>Rank</th><th>Symbol</th><th>Category</th><th>Score</th><th>1M</th><th>3M</th><th>6M</th>"
           f"<th>12M</th><th>Price</th><th>DMA</th><th>Flags</th></tr></thead><tbody>{trows}</tbody></table></div>")

    sort_js = """
    <script>
    document.querySelectorAll('#t th').forEach(function(th,i){th.addEventListener('click',function(){
      var tb=document.querySelector('#t tbody');var rows=[].slice.call(tb.querySelectorAll('tr'));
      var asc=th._asc=!th._asc;
      rows.sort(function(a,b){var x=a.children[i].innerText.replace(/[^0-9.\\-]/g,''),y=b.children[i].innerText.replace(/[^0-9.\\-]/g,'');
        var nx=parseFloat(x),ny=parseFloat(y);if(!isNaN(nx)&&!isNaN(ny))return asc?nx-ny:ny-nx;
        return asc?a.children[i].innerText.localeCompare(b.children[i].innerText):b.children[i].innerText.localeCompare(a.children[i].innerText);});
      rows.forEach(function(r){tb.appendChild(r);});});});
    </script>"""

    html = ("<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'>"
            "<meta name='viewport' content='width=device-width, initial-scale=1'>"
            "<title>ETF Momentum Live Tracker</title>"
            "<link href='https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap' rel='stylesheet'>"
            f"<style>{css}</style></head><body><div class='wrap'>"
            f"{header}{regime_banner}<div class='kpis'>{kpi_html}</div>"
            f"<h2>Still on hold</h2>{hold_block}"
            f"<h2>Flagged to sell</h2>{sell_block}"
            f"<h2>Buy candidates (top-6, 3M&gt;0, above 200-DMA, not held)</h2>{buy_block}"
            f"<h2>Charts</h2><div class='chart'>{charts}</div>"
            f"<h2>Full ranking ({len(ranks)} ETFs)</h2>{tbl}"
            f"<h2>Rules</h2>{rules}{note}</div>{sort_js}</body></html>")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)


# =============================================================================
# TRADETRON SIGNAL SENDER  (optional — only runs when SEND_TT_SIGNALS=1)
# =============================================================================
# Sends, for every ETF, one "rank code" to the Tradetron strategy, then the regime,
# and LAST a month stamp (YYYYMM). Tradetron acts only when the stamp changes, i.e.
# on the first trading day of a new month that the data arrives, at 3:00-3:20 pm.
#   rank code 1-99    = rank, and allowed to be bought (3M > 0 AND above its 200-DMA)
#   rank code 101-199 = rank + 100, NOT allowed to be bought (still used for the rank-10 exit)
#   rank code 0       = no data for this ETF today -> Tradetron takes no action on it
#   regime 1 = Nifty 50 above its 200-DMA, 2 = below (sell all, stay in cash)
# The token is read from the TRADETRON_TOKEN secret. Never type the token in this file.
TT_WEBHOOK_URL   = "https://api.tradetron.tech/api?"   # the trailing ? is required
TT_PAIRS_PER_CALL = 3        # variables per call
TT_PAUSE_SECS     = 2.5      # Tradetron limits: 3 calls/second, 30/minute, 250/hour
TT_MIN_SCORED     = 40       # safety: do not send if fewer ETFs than this could be ranked
TT_SESSION        = (920, 1525)   # only send while the market is open (IST, HHMM)


def _tt_post(token, pairs):
    """Send [(variable, value), ...] in one call. Returns True on success."""
    body = {"auth-token": token}
    for i, (k, v) in enumerate(pairs):
        sfx = "" if i == 0 else str(i)
        body["key" + sfx] = str(k)
        body["value" + sfx] = str(v)
    data = json.dumps(body).encode()
    for attempt in range(1, 4):
        try:
            req = urllib.request.Request(
                TT_WEBHOOK_URL, data=data, method="POST",
                headers={"Content-Type": "application/json", "User-Agent": "tt-webhook/1.0"})
            with urllib.request.urlopen(req, timeout=20) as resp:
                text = resp.read().decode(errors="replace")
                if resp.status == 200 and '"success":false' not in text.replace(" ", "").lower():
                    return True
                print(f"      Tradetron replied {resp.status}: {text[:200]}")
        except urllib.error.HTTPError as e:
            print(f"      Tradetron HTTP error {e.code} (attempt {attempt}/3)")
        except Exception as e:
            print(f"      Could not reach Tradetron: {type(e).__name__} (attempt {attempt}/3)")
        time.sleep(6 * attempt)
    return False


def build_tt_pairs(ranks, regime_on):
    """Rank codes for all ETFs in the universe + the regime flag."""
    by_sym = {r["symbol"]: r for _, r in ranks.iterrows()}
    pairs = []
    for sym, _cat in C54:
        r = by_sym.get(sym)
        if r is None:
            code = 0
        else:
            rk = int(r["rank"])
            can_buy = bool(r["above_dma"]) and float(r["ret_3m"]) > 0
            code = rk if can_buy else 100 + rk
        pairs.append((f"{sym}_r", code))
    pairs.append(("regime", 1 if regime_on else 2))
    return pairs


def send_tradetron_signals(prices, ranks, regime_on):
    """Returns True (sent), False (tried and failed) or None (skipped on purpose)."""
    if os.getenv("SEND_TT_SIGNALS") != "1":
        return None
    print("[TT] Tradetron signals → ", end="")
    dry = os.getenv("TT_DRY_RUN") == "1"
    force = os.getenv("TT_FORCE") == "1"
    token = os.getenv("TRADETRON_TOKEN", "").strip()
    if not token and not dry:
        print("FAILED — the TRADETRON_TOKEN secret is missing")
        return False
    if AS_OF:
        print("skipped — AS_OF is set (a past date), signals are sent only for today")
        return None
    now = datetime.now(ZoneInfo("Asia/Kolkata"))
    hhmm = now.hour * 100 + now.minute
    last_date = max(p.index[-1] for p in prices.values() if len(p)).date()
    if not force:
        if last_date != now.date():
            print(f"skipped — latest price is {last_date}, not today ({now.date()}): "
                  f"market holiday or prices not updated yet")
            return None
        if not (TT_SESSION[0] <= hhmm <= TT_SESSION[1]):
            print(f"skipped — market is closed (time now {now:%H:%M} IST)")
            return None
    if regime_on is None:
        print("FAILED — could not work out the Nifty 50 regime (no index data)")
        return False
    if len(ranks) < TT_MIN_SCORED:
        print(f"FAILED — only {len(ranks)} ETFs could be ranked (need {TT_MIN_SCORED}); data problem")
        return False

    pairs = build_tt_pairs(ranks, regime_on)
    stamp = now.year * 100 + now.month
    n_buyable = sum(1 for k, v in pairs if k.endswith("_r") and 1 <= v <= N_HOLD)
    print(f"{len(pairs) - 1} rank codes, regime {'ON' if regime_on else 'OFF'}, "
          f"stamp {stamp}, top-{N_HOLD} buyable now: {n_buyable}")
    if dry:
        for k, v in pairs:
            print(f"      {k} = {v}")
        print(f"      reb = {stamp}   (dry run — nothing sent)")
        return None
    calls = [pairs[i:i + TT_PAIRS_PER_CALL] for i in range(0, len(pairs), TT_PAIRS_PER_CALL)]
    for n, chunk in enumerate(calls, 1):
        if not _tt_post(token, chunk):
            print(f"      FAILED at call {n}/{len(calls)} — month stamp NOT sent, "
                  f"so Tradetron will not act on a half-delivered ranking")
            return False
        time.sleep(TT_PAUSE_SECS)
    # The stamp goes last and alone: Tradetron only acts once every rank has arrived.
    if not _tt_post(token, [("reb", stamp)]):
        print("      FAILED sending the month stamp — Tradetron will not act today")
        return False
    print(f"      sent ✓ ({len(calls) + 1} calls)")
    return True


# =============================================================================
# MAIN
# =============================================================================
def main():
    print("ETF Momentum Rotation — Live Holdings Tracker")
    print("-" * 44)
    asof = pd.Timestamp(AS_OF) if AS_OF else None
    print("[1] Loading prices    → ", end="")
    prices, bench, regime = load_prices()
    print(f"[2] Ranking universe  → scoring {len(C54)} ETFs, 200-DMA gate")
    ranks = compute_ranks(prices, asof)

    # Market-regime master switch: Nifty 50 vs its 200-DMA
    reg = regime.loc[:asof] if (asof is not None and regime is not None) else regime
    regime_on, r_cur, r_dma = None, None, None
    if reg is not None and len(reg) >= REGIME_DMA:
        r_cur = float(reg.iloc[-1]); r_dma = float(reg.iloc[-REGIME_DMA:].mean())
        regime_on = r_cur > r_dma
    print(f"[3] Regime check      → {REGIME_LABEL} vs {REGIME_DMA}-DMA: "
          f"{'ON (invested)' if regime_on else ('OFF (all cash)' if regime_on is False else 'unknown')}")

    tt_result = send_tradetron_signals(prices, ranks, regime_on)

    print("[4] Reading holdings  → ", end="")
    holdings = read_holdings()
    print(f"{len(holdings)} holding(s) in {HOLDINGS_FILE}")
    print("[5] Checking exits    → regime / SL / trail / rotation")
    hold_df = evaluate(holdings, prices, ranks, asof, regime_on) if len(holdings) else \
        pd.DataFrame(columns=["symbol", "status", "current", "pnl"])
    held_syms = set(hold_df["symbol"]) if len(hold_df) else set()
    if regime_on is False:
        buys = ranks.iloc[0:0]   # regime OFF -> no buying, sit in cash
    else:
        buys = ranks[(ranks["top6"]) & (ranks["above_dma"]) & (ranks["ret_3m"] > 0)
                     & (~ranks["symbol"].isin(held_syms))]

    last_date = max([(p.loc[:asof] if asof is not None else p).index[-1] for p in prices.values() if len(p)])
    n_hold = int((hold_df["status"] == "HOLD").sum()) if len(hold_df) else 0
    n_sell = int((hold_df["status"] == "SELL").sum()) if len(hold_df) else 0
    val = float(hold_df["value"].dropna().sum()) if len(hold_df) and "value" in hold_df else 0.0
    if len(hold_df) and "value" in hold_df and hold_df["value"].notna().any():
        cost = float((hold_df["entry_price"] * hold_df["qty"]).sum())
        pnl = (val / cost - 1) * 100 if cost else None
    else:
        pnl = None
    meta = {
        "as_of": last_date.strftime("%d %b %Y"), "n_hold_total": len(holdings),
        "n_hold": n_hold, "n_sell": n_sell, "n_buy": len(buys), "value": val, "pnl": pnl,
        "slot_size": INITIAL_CAPITAL / N_HOLD, "generated": datetime.now().strftime("%d %b %Y %H:%M"),
        "regime_on": regime_on, "regime_cur": r_cur, "regime_dma": r_dma,
    }
    print(f"    still on HOLD={n_hold} | flagged SELL={n_sell} | buy candidates={len(buys)}")
    print("[6] Excel saved       → ", end="")
    build_excel(hold_df, ranks, buys, meta, EXCEL_PATH); print(EXCEL_PATH.name, "✓")
    print("[7] HTML saved        → ", end="")
    build_html(hold_df, ranks, buys, meta, HTML_PATH); print(HTML_PATH.name, "✓")
    print("[8] Opening browser   → ", end="")
    try:
        webbrowser.open(HTML_PATH.as_uri()); print("report launched ✓")
    except Exception:
        print("no browser (headless) — open the HTML from the output folder")
    print(f"\nSaved: {EXCEL_PATH}\nSaved: {HTML_PATH}")
    if tt_result is False:
        print("\nTradetron signals were NOT delivered — see the [TT] lines above.")
        sys.exit(1)


if __name__ == "__main__":
    main()
