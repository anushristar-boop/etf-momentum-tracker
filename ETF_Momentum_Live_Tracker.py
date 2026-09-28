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

import os, math, webbrowser
from datetime import datetime
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
N_HOLD      = 6
LB          = [21, 63, 126, 252]
WEIGHTS     = [0.15, 0.40, 0.30, 0.15]
DMA_PERIOD  = 200
MIN_HISTORY = 262
SL_PCT      = 0.15      # hard stop: sell if price <= entry * 0.85
TRAIL_PCT   = 0.20      # trail stop: sell if price <= peak * 0.80
BENCH       = "^CRSLDX"
BENCH_LABEL = "Nifty 500"

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
    ytk = [s + ".NS" for s in syms] + [BENCH]
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
    print(f"{len(prices)} ETFs loaded, {dropped} dropped (insufficient history)")
    return prices, bench


def _mock_prices():
    rng = np.random.default_rng(11)
    idx = pd.bdate_range(end=datetime.now().date(), periods=800)
    prices = {}
    for i, (s, _) in enumerate(C54):
        drift = 0.0003 + (i % 9) * 0.00007
        prices[s] = pd.Series(100 * np.exp(np.cumsum(rng.normal(drift, 0.013, len(idx)))), index=idx)
    bench = pd.Series(100 * np.exp(np.cumsum(rng.normal(0.0004, 0.009, len(idx)))), index=idx)
    print(f"[MOCK] generated {len(prices)} ETFs + benchmark")
    return prices, bench


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


def evaluate(holdings, prices, ranks, asof):
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
        in_top6 = bool(rinfo is not None and rinfo["top6"])
        if cur <= sl_level:
            status, reason = "SELL", f"Hard SL — down {(cur/entry-1)*100:.1f}% from entry (limit -{int(SL_PCT*100)}%)"
        elif cur <= trail_level:
            status, reason = "SELL", f"Trail stop — down {(cur/peak-1)*100:.1f}% from peak (limit -{int(TRAIL_PCT*100)}%)"
        elif not in_top6:
            status, reason = "SELL", (f"Rotation — rank {rk} (out of top {N_HOLD})" if rk
                                      else f"Rotation — no longer ranked in top {N_HOLD}")
        else:
            status, reason = "HOLD", f"in top {N_HOLD} (rank {rk}), stops intact"
        out.append({
            "symbol": sym, "category": CATEGORY.get(sym, ""),
            "entry_date": (h["entry_date"].date().isoformat() if not pd.isna(h["entry_date"]) else "-"),
            "entry_price": entry, "qty": int(h["qty"]), "current": cur, "peak": peak,
            "pnl": (cur / entry - 1) * 100, "value": cur * int(h["qty"]),
            "rank": rk, "in_top6": in_top6,
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
        summ = pd.DataFrame({
            "Metric": ["As of", "Holdings", "Still on HOLD", "Flagged to SELL",
                       "Portfolio value (₹)", "Total P&L %", "Buy candidates (top-6 not held, above DMA)",
                       "Generated"],
            "Value": [meta["as_of"], meta["n_hold_total"], meta["n_hold"], meta["n_sell"],
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
        buy_block = ("<div class='empty'>No buy candidates — the top-6 not held are all below their "
                     "200-DMA, so those slots stay in cash.</div>")

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
    """

    header = (f"<h1>ETF Momentum Rotation — Live Holdings Tracker</h1>"
              f"<div class='sub'>As of {meta['as_of']} · {meta['n_hold_total']} holdings · "
              f"{meta['n_hold']} still on hold · {meta['n_sell']} flagged to sell · benchmark {BENCH_LABEL}</div>")

    rules = (f"<div class='rules'>A holding is <b>kept</b> until one of three exits fires (priority order): "
             f"<b>Hard SL</b> — price ≤ entry × (1−{int(SL_PCT*100)}%); <b>Trail stop</b> — price ≤ "
             f"peak-since-entry × (1−{int(TRAIL_PCT*100)}%); <b>Rotation</b> — it drops out of the top "
             f"{N_HOLD} by momentum score. Momentum score = 0.15·1M + 0.40·3M + 0.30·6M + 0.15·12M returns. "
             f"New buys come from the top {N_HOLD} not currently held that are above their {DMA_PERIOD}-day "
             f"average. Peak-since-entry basis: <b>{PEAK_BASIS}</b>.</div>")

    note = ("<div class='note'><b>Live monitor.</b> Status uses today's prices; the strategy rebalances at "
            "each month-start, so a flagged SELL is what you'd act on at the next rebalance (the hard/trail "
            "stops let you see a breach the moment it happens). Prices are Yahoo Finance auto-adjusted closes. "
            "Peak-since-entry is computed from price history; if you use a tighter live trailing stop, switch "
            "PEAK_BASIS. Not investment advice.</div>")

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
            f"{header}<div class='kpis'>{kpi_html}</div>"
            f"<h2>Still on hold</h2>{hold_block}"
            f"<h2>Flagged to sell</h2>{sell_block}"
            f"<h2>Buy candidates (top-6 not held, above 200-DMA)</h2>{buy_block}"
            f"<h2>Charts</h2><div class='chart'>{charts}</div>"
            f"<h2>Full ranking ({len(ranks)} ETFs)</h2>{tbl}"
            f"<h2>Rules</h2>{rules}{note}</div>{sort_js}</body></html>")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)


# =============================================================================
# MAIN
# =============================================================================
def main():
    print("ETF Momentum Rotation — Live Holdings Tracker")
    print("-" * 44)
    asof = pd.Timestamp(AS_OF) if AS_OF else None
    print("[1] Loading prices    → ", end="")
    prices, bench = load_prices()
    print(f"[2] Ranking universe  → scoring {len(C54)} ETFs, 200-DMA gate")
    ranks = compute_ranks(prices, asof)
    print("[3] Reading holdings  → ", end="")
    holdings = read_holdings()
    print(f"{len(holdings)} holding(s) in {HOLDINGS_FILE}")
    print("[4] Checking exits    → SL / trail / rotation")
    hold_df = evaluate(holdings, prices, ranks, asof) if len(holdings) else \
        pd.DataFrame(columns=["symbol", "status", "current", "pnl"])
    held_syms = set(hold_df["symbol"]) if len(hold_df) else set()
    buys = ranks[(ranks["top6"]) & (ranks["above_dma"]) & (~ranks["symbol"].isin(held_syms))]

    used = ranks["price"].notna()  # as-of label
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
    }
    print(f"    still on HOLD={n_hold} | flagged SELL={n_sell} | buy candidates={len(buys)}")
    print("[5] Excel saved       → ", end="")
    build_excel(hold_df, ranks, buys, meta, EXCEL_PATH); print(EXCEL_PATH.name, "✓")
    print("[6] HTML saved        → ", end="")
    build_html(hold_df, ranks, buys, meta, HTML_PATH); print(HTML_PATH.name, "✓")
    print("[7] Opening browser   → ", end="")
    try:
        webbrowser.open(HTML_PATH.as_uri()); print("report launched ✓")
    except Exception:
        print("no browser (headless) — open the HTML from the output folder")
    print(f"\nSaved: {EXCEL_PATH}\nSaved: {HTML_PATH}")


if __name__ == "__main__":
    main()
