#!/usr/bin/env python3
# =============================================================================
# MANJU RANA DUAL ENGINE — daily check, Telegram alerts and buttons
# Works with the Tradetron strategy "Manju Rana Dual Engine" (Shoonya, e-DIS before every sale).
#
#   python manju_engine.py daily          (scheduled ~9:35 am; "Check now" button)
#       Works out today's sales (stop-losses, monthly rotation / regime, Nifty:Gold reshuffle),
#       sends the sale flags + month-end rankings to Tradetron, and a Telegram message listing
#       what to authorise with e-DIS before 3:00 pm.
#   python manju_engine.py confirm        ("e-DIS done" button)  -> sales go through at 3:05 / 3:10 pm today
#   python manju_engine.py oldgold        ("Old gold sold" button) -> her 3,418 old GOLDBEES sold by hand
#   python manju_engine.py topup 200000   ("Top-up" button)      -> new money: 49% Part A, 49% Part B, 2% cash
#
# Secrets (GitHub > Settings > Secrets and variables > Actions):
#   TRADETRON_TOKEN_MANJU, TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
# GitHub keeps a small record of Part B's ETFs in manju/state.json (the workflow saves it back).
# Never type a token in this file.
# =============================================================================
import os, sys, json, time, urllib.request, urllib.error
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

import ETF_Momentum_Live_Tracker as T   # same rankings, regime and webhook code as the 1L strategy
import pandas as pd

STATE_FILE = Path(os.getenv("MANJU_STATE", Path(__file__).parent / "manju" / "state.json"))
DRY = os.getenv("MANJU_DRY") == "1"
FORCE = os.getenv("MANJU_FORCE") == "1"
IST = ZoneInfo("Asia/Kolkata")
SL_PCT, TRAIL_PCT, HOLD_RANK, N_SLOTS, TOPN = 0.15, 0.20, 10, 6, 6
NG_STEP = 1.0
CONFIRM_LAST = 1504          # e-DIS done must be clicked by 3:04 pm (sales run 3:05 / 3:10 pm)
SESSION = (920, 1504)        # Tradetron only accepts signals while the market is open


# ----------------------------------------------------------------------------- helpers
def now_ist():
    if os.getenv("MANJU_NOW"):                       # testing only: "2026-10-13 09:35"
        return datetime.strptime(os.getenv("MANJU_NOW"), "%Y-%m-%d %H:%M").replace(tzinfo=IST)
    return datetime.now(IST)


def load_state():
    return json.loads(STATE_FILE.read_text())


def save_state(st):
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps(st, indent=2, sort_keys=True))


def telegram(text):
    tok, chat = os.getenv("TELEGRAM_BOT_TOKEN", "").strip(), os.getenv("TELEGRAM_CHAT_ID", "").strip()
    print("\n----- Telegram -----\n" + text + "\n--------------------")
    if DRY:
        return True
    if not tok or not chat:
        print("Telegram secrets missing — message not sent")
        return False
    data = json.dumps({"chat_id": chat, "text": text}).encode()
    for attempt in range(3):
        try:
            req = urllib.request.Request(f"https://api.telegram.org/bot{tok}/sendMessage", data=data,
                                         headers={"Content-Type": "application/json"}, method="POST")
            with urllib.request.urlopen(req, timeout=20) as r:
                if r.status == 200:
                    return True
        except urllib.error.HTTPError as e:
            try:
                why = json.loads(e.read().decode()).get("description", "")
            except Exception:
                why = ""
            print(f"Telegram error: HTTP {e.code} {why}")
            if e.code in (400, 401, 403, 404):
                return False
        except Exception as e:
            print(f"Telegram error: {type(e).__name__}")
        time.sleep(5)
    return False


def tt_send(pairs):
    """Send [(variable, value), ...] to her Tradetron strategy, 3 per call."""
    if DRY:
        for k, v in pairs:
            print(f"   [dry] {k} = {v}")
        return True
    token = os.getenv("TRADETRON_TOKEN_MANJU", "").strip()
    if not token:
        print("TRADETRON_TOKEN_MANJU secret is missing")
        return False
    for i in range(0, len(pairs), 3):
        if not T._tt_post(token, pairs[i:i + 3]):
            return False
        time.sleep(2.5)
    return True


def in_session(now):
    hhmm = now.hour * 100 + now.minute
    return now.weekday() <= 4 and SESSION[0] <= hhmm <= SESSION[1]


def latest(series):
    return float(series.iloc[-1])


def fmt_units(sym):
    return f"ALL your {sym} units"


# ----------------------------------------------------------------------------- the daily decision
def decide(st, prices, regime, now):
    """Returns (plan, signal_pairs, problem). plan = what to sell today and why."""
    sig, ranks, regime_on, problem = T.month_end_signal(prices, regime, now)
    if problem:
        return None, None, problem
    if len(ranks) < T.TT_MIN_SCORED:
        return None, None, f"only {len(ranks)} ETFs could be ranked"
    stamp = now.year * 100 + now.month
    rank_of = {r["symbol"]: int(r["rank"]) for _, r in ranks.iterrows()}
    month_due = st.get("month_applied") != stamp
    today = now.date()

    # fill entry prices of ETFs bought on an earlier day (the close of the buy day)
    for sym, h in st["holdings"].items():
        if h.get("entry") is None and sym in prices:
            s = prices[sym]
            bd = pd.Timestamp(h["buy_date"])
            on = s.loc[s.index.normalize() == bd]
            if len(on) and bd.date() < today:
                h["entry"] = round(float(on.iloc[-1]), 4)

    sells = []          # (symbol, reason)
    for sym, h in sorted(st["holdings"].items()):
        s = prices.get(sym)
        if s is None or h.get("entry") is None:
            continue
        cur = latest(s)
        peak = float(s.loc[s.index >= pd.Timestamp(h["buy_date"])].max())
        rk = rank_of.get(sym)
        if month_due and not regime_on:
            sells.append((sym, "regime OFF (Nifty 50 below its 200-DMA) — all ETFs move to GOLDBEES"))
        elif cur <= h["entry"] * (1 - SL_PCT):
            sells.append((sym, f"stop-loss: {(cur / h['entry'] - 1) * 100:.1f}% from buy price"))
        elif cur <= peak * (1 - TRAIL_PCT):
            sells.append((sym, f"trailing stop: {(cur / peak - 1) * 100:.1f}% from its peak"))
        elif month_due and (rk is None or rk > HOLD_RANK):
            sells.append((sym, f"monthly rotation: rank {rk if rk else 'n/a'} (beyond {HOLD_RANK})"))

    gold_sale = bool(month_due and regime_on and st["park"] and st["park_tt"])
    old_gold_by_hand = bool(month_due and regime_on and st["park"] and st["adopt"])

    ng = None
    pn, pg = latest(prices["NIFTYBEES"]), latest(prices["GOLDBEES"])
    ratio = pn / pg
    if st.get("ng_ref") is None:
        st["ng_ref"] = round(ratio, 4)
    elif abs(ratio - st["ng_ref"]) >= NG_STEP:
        ng = "NIFTYBEES" if ratio > st["ng_ref"] else "GOLDBEES"

    # buys Tradetron will make at the monthly rebalance (mirror only)
    held_after = [s for s in st["holdings"] if s not in {x for x, _ in sells}]
    buys = []
    if month_due and regime_on:
        elig = ranks[(ranks["rank"] <= TOPN) & ranks["above_dma"] & (ranks["ret_3m"] > 0)]
        for _, r in elig.sort_values("rank").iterrows():
            if len(held_after) + len(buys) >= N_SLOTS:
                break
            if r["symbol"] not in held_after:
                buys.append(r["symbol"])

    plan = {"date": today.isoformat(), "stamp": stamp, "month_due": month_due, "regime_on": regime_on,
            "sells": sells, "gold_sale": gold_sale, "old_gold_by_hand": old_gold_by_hand,
            "ng": ng, "ratio": round(ratio, 3), "buys": buys, "signal_date": str(sig.date())}
    pairs = T.build_tt_pairs(ranks, regime_on)
    return plan, pairs, None


def edis_list(plan):
    items = [f"• {fmt_units(s)} — {why}" for s, why in plan["sells"]]
    if plan["gold_sale"]:
        items.append("• Part B's GOLDBEES (the regime-off gold, incl. her old 3,418 now in Tradetron) — regime is ON again. Part A's GOLDBEES are NOT sold: e-DIS only the quantity in Tradetron's Part B GOLDBEES sell order")
    if plan["ng"]:
        items.append(f"• ALL your {plan['ng']} units — Nifty:Gold ratio now {plan['ratio']} "
                     f"(moved 1.0 from the last reshuffle): Part A reshuffles")
    return items


def needs_confirm(plan):
    return bool(plan["sells"] or plan["gold_sale"] or plan["ng"])


def apply_plan(st, plan, now, confirmed):
    """Update GitHub's record the way Tradetron will act this afternoon."""
    sold = {s for s, _ in plan["sells"]}
    if confirmed:
        for s in sold:
            st["holdings"].pop(s, None)
        if plan["ng"]:
            st["ng_ref"] = plan["ratio"]
        if plan["gold_sale"]:
            st["park_tt"] = False
    month_done = False
    if plan["month_due"] and (confirmed or not needs_confirm(plan)):
        if not plan["regime_on"]:
            st["park"], st["park_tt"] = True, True
            month_done = True
        else:
            st["park"] = bool(st["park"] and (st["park_tt"] or st["adopt"]))
            if not st["park"]:
                for b in plan["buys"]:
                    st["holdings"][b] = {"buy_date": now.date().isoformat(), "entry": None}
                month_done = True
    if month_done:
        st["month_applied"] = plan["stamp"]
    return month_done


# ----------------------------------------------------------------------------- modes
def mode_daily(st, now, from_button=False):
    if st.get("last_daily") == now.date().isoformat() and not FORCE and not from_button:
        print("Already checked today — nothing to do.")
        if os.getenv("MANJU_MANUAL") == "1":
            lp = st.get("last_plan") or {}
            pend = st.get("pending")
            if pend:
                telegram("Today's check is done. Still waiting for e-DIS on:\n" + "\n".join(edis_list(pend)) +
                         "\nThen press 'Manju: e-DIS done' by 3:04 pm.")
            else:
                telegram("Today's check is done: nothing waiting for e-DIS." +
                         (" (Monthly rebalance at 3:05 pm.)" if lp.get("month_due") else ""))
        return True, st.get("last_plan")
    if not in_session(now):
        print(f"Skipped: market closed or too late (now {now:%a %H:%M} IST; checks run 9:20 am – 3:04 pm, Mon–Fri)")
        return None, None
    prices, _bench, regime = T.load_prices()
    last_date = max(p.index[-1] for p in prices.values() if len(p)).date()
    if last_date != now.date() and os.getenv("SCREENER_MOCK") != "1":
        print(f"Skipped: latest price is {last_date}, not today — market holiday or prices not in yet")
        return None, None
    plan, pairs, problem = decide(st, prices, regime, now)
    if problem:
        telegram(f"Manju Rana Dual Engine — daily check FAILED: {problem}. Nothing was sent to Tradetron.")
        return False, None

    # 1) sale flags (sd=1 for today's sales, 0 for every other Part B ETF GitHub knows of), Part A flag
    flags = [(f"{s}_sd", 1) for s, _ in plan["sells"]]
    flags += [(f"{s}_sd", 0) for s in st["holdings"] if s not in {x for x, _ in plan["sells"]}]
    flags.append(("ng_go", 1 if plan["ng"] else 0))
    # 2) month-end rankings + regime, then the month stamp LAST
    ok = tt_send(flags) and tt_send(pairs) and tt_send([("reb", plan["stamp"])])
    if not ok:
        telegram("Manju Rana Dual Engine — could NOT reach Tradetron this morning. "
                 "Please press 'Manju: check now' in GitHub, or check the strategy.")
        return False, None

    month_done = apply_plan(st, plan, now, confirmed=False)
    st["last_daily"] = now.date().isoformat()
    st["last_plan"] = plan
    st["pending"] = plan if needs_confirm(plan) else None

    lines = [f"Manju Rana Dual Engine — {now:%a %d %b %Y}, {now:%H:%M}"]
    if needs_confirm(plan):
        lines += ["", "SALE TODAY. Please do e-DIS in Shoonya before 3:00 pm for:"] + edis_list(plan)
        if plan["old_gold_by_hand"]:
            lines += ["", "ALSO sell by hand her OLD 3,418 GOLDBEES today (e-DIS needed), "
                      "then press 'Manju: old gold sold' by 3:04 pm."]
        lines += ["", "Then press 'Manju: e-DIS done' in GitHub (by 3:04 pm).",
                  "Sales go through at 3:05 pm (monthly) / 3:10 pm (stop-loss, reshuffle). "
                  "No click = nothing is sold today; I will ask again tomorrow."]
        if plan["buys"]:
            lines += ["", "After the sales Tradetron buys: " + ", ".join(plan["buys"])]
    elif plan["old_gold_by_hand"]:
        lines += ["", "Regime is ON again. Please sell her OLD 3,418 GOLDBEES by hand in Shoonya today "
                  "(e-DIS needed), then press 'Manju: old gold sold' by 3:04 pm. Tradetron then buys: "
                  + ", ".join(plan["buys"])]
    elif plan["month_due"] or month_done:
        lines += ["", f"Monthly rebalance today at 3:05 pm (no sale, no e-DIS needed). Regime "
                  f"{'ON' if plan['regime_on'] else 'OFF'}."]
        if plan["buys"]:
            lines.append("Tradetron buys: " + ", ".join(plan["buys"]))
    else:
        print("No sale today — no Telegram message.")
        return True, plan
    telegram("\n".join(lines))
    return True, plan


def mode_confirm(st, now):
    hhmm = now.hour * 100 + now.minute
    if now.weekday() > 4 or hhmm < SESSION[0] or hhmm > CONFIRM_LAST:
        telegram(f"'e-DIS done' was pressed at {now:%H:%M} — too late or market closed. "
                 "It must be pressed between 9:20 am and 3:04 pm. Nothing was sold; I will ask again next trading day.")
        return False
    if st.get("last_daily") != now.date().isoformat():
        print("Today's check had not run yet — running it first.")
        ok, _ = mode_daily(st, now, from_button=True)
        if not ok:
            return False
    plan = st.get("pending")
    if not plan or plan.get("date") != now.date().isoformat():
        telegram("'e-DIS done' received, but there is no sale due today. Nothing changed.")
        return True
    if not tt_send([("ok_day", now.strftime("%Y%m%d"))]):
        telegram("'e-DIS done' could NOT reach Tradetron. Please press it again.")
        return False
    apply_plan(st, plan, now, confirmed=True)
    st["pending"] = None
    when = "3:05 pm" if plan["month_due"] else "3:10 pm"
    telegram(f"Confirmed. Tradetron will sell at {when} today:\n" + "\n".join(edis_list(plan)) +
             (("\nThen it buys: " + ", ".join(plan["buys"])) if plan["buys"] and not st["park"] else ""))
    return True


def mode_oldgold(st, now):
    if not in_session(now):
        telegram("'Old gold sold' must be pressed while the market is open (9:20 am – 3:04 pm). Please press it again then.")
        return False
    if not st["adopt"]:
        telegram("Her old GOLDBEES are already marked as sold. Nothing changed.")
        return True
    if not tt_send([("adopt_g", 0)]):
        telegram("'Old gold sold' could NOT reach Tradetron. Please press it again.")
        return False
    st["adopt"] = False
    msg = "Noted: her old 3,418 GOLDBEES are sold."
    plan = st.get("last_plan")
    if plan and plan.get("month_due") and plan.get("regime_on") and not st["park_tt"]:
        st["park"] = False
        if st.get("month_applied") != plan["stamp"] and (now.hour * 100 + now.minute) <= CONFIRM_LAST:
            for b in plan["buys"]:
                st["holdings"][b] = {"buy_date": now.date().isoformat(), "entry": None}
            st["month_applied"] = plan["stamp"]
            msg += " Tradetron buys this month's ETFs at 3:05 pm today: " + ", ".join(plan["buys"])
    telegram(msg)
    return True


def mode_topup(st, now, amount):
    if not in_session(now):
        telegram("'Top-up' must be pressed while the market is open (9:20 am – 3:04 pm). Nothing changed.")
        return False
    if amount < 1000:
        telegram(f"Top-up amount {amount} looks wrong. Nothing changed.")
        return False
    tu_id = int(now.strftime("%Y%m%d%H%M"))
    if not tt_send([("tu_amt", int(amount)), ("tu_id", tu_id)]):
        telegram("'Top-up' could NOT reach Tradetron. Please press it again.")
        return False
    st.setdefault("topups", []).append({"date": now.date().isoformat(), "amount": int(amount)})
    telegram(f"Top-up of ₹{amount:,.0f} noted. At 3:05 pm today Tradetron puts 49% into Part A "
             f"(NIFTYBEES/GOLDBEES as per the ratio) and 49% into Part B "
             f"({'GOLDBEES, regime is OFF' if st['park'] else 'kept as cash until the next monthly rebalance'}); "
             f"2% stays as cash. Make sure the money is in her Shoonya account before 3:05 pm.")
    return True


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "daily"
    st = load_state()
    now = now_ist()
    print(f"Manju Rana Dual Engine — {mode} — {now:%Y-%m-%d %H:%M} IST")
    if mode == "daily":
        ok, _ = mode_daily(st, now)
    elif mode == "confirm":
        ok = mode_confirm(st, now)
    elif mode == "oldgold":
        ok = mode_oldgold(st, now)
    elif mode == "test":
        sample = {"sells": [("ITBEES", "stop-loss: -15.4% from buy price")], "gold_sale": False, "ng": None, "ratio": 2.18}
        ok = telegram("TEST MESSAGE — nothing to do, this is only a sample.\n\n"
                      f"Manju Rana Dual Engine — {now:%a %d %b %Y}, {now:%H:%M}\n\n"
                      "SALE TODAY. Please do e-DIS in Shoonya before 3:00 pm for:\n" + "\n".join(edis_list(sample)) +
                      "\n\nThen press 'Manju: e-DIS done' in GitHub (by 3:04 pm).\n"
                      "Sales go through at 3:05 pm (monthly) / 3:10 pm (stop-loss, reshuffle). "
                      "No click = nothing is sold today; I will ask again tomorrow.")
        print("sample sent" if ok else "sample NOT sent")
        if not ok:
            # Show which chats have written to the bot, so the right TELEGRAM_CHAT_ID can be set.
            try:
                tok = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
                with urllib.request.urlopen(f"https://api.telegram.org/bot{tok}/getUpdates", timeout=20) as r:
                    ups = json.loads(r.read().decode()).get("result", [])
                seen = {}
                for u in ups:
                    c = (u.get("message") or u.get("my_chat_member") or {}).get("chat") or {}
                    if c.get("id"):
                        seen[c["id"]] = c.get("first_name") or c.get("title") or ""
                print("Chats that messaged the bot:", seen if seen else "none yet - open the bot in Telegram and press Start")
                print("TELEGRAM_CHAT_ID secret has", len(os.getenv("TELEGRAM_CHAT_ID", "").strip()), "characters")
            except Exception as e:
                print(f"getUpdates failed: {type(e).__name__}")
            sys.exit(1)
        return
    elif mode == "topup":
        ok = mode_topup(st, now, float(str(sys.argv[2]).replace(",", "")) if len(sys.argv) > 2 else 0)
    else:
        print("unknown mode"); sys.exit(2)
    save_state(st)
    if ok is False:
        sys.exit(1)


if __name__ == "__main__":
    main()
