#!/usr/bin/env python3
# =============================================================================
# US GOLD RATIO (SPY + GLD) — daily check, Telegram alerts and buttons
# Works with the Tradetron strategy "US Gold Ratio SPY GLD" (template 999398326, Live Offline;
# the owner copies the orders by hand at her US broker).
#
#   python us_gold_engine.py daily            (outside timer at 9:45 am New York; "Check now" button)
#       Reads SPY and GLD prices. If the SPY/GLD ratio has moved 0.5 from the ratio at the last
#       purchase / reshuffle, it tells Tradetron to reshuffle (sends ng_ref = 0) and sends a Telegram alert.
#   python us_gold_engine.py topup 2000       ("Top-up" button)    -> new dollars go in by the current ratio
#   python us_gold_engine.py withdraw 2000    ("Withdraw" button)  -> SPY and GLD sold in proportion
#   python us_gold_engine.py status           ("Status" button)    -> where the ratio stands, sent to Telegram
#   python us_gold_engine.py test             ("Status" button, choice "test") -> one sample Telegram message
#
# Secrets (GitHub > Settings > Secrets and variables > Actions):
#   TRADETRON_TOKEN_USGOLD, TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
# GitHub keeps a small record in usgold/state.json; the workflows save it back.
# This file is separate from the Manju / Abhishek engine and changes none of their files.
# Never type a token in this file.
# Team Infinity is an algorithm / technology provider only and is not SEBI-registered. Not investment advice.
# =============================================================================
import os, sys, json, time, urllib.request, urllib.error
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

NAME = "US Gold Ratio (SPY + GLD)"
STATE_FILE = Path(os.getenv("USGOLD_STATE", Path(__file__).parent / "usgold" / "state.json"))
TOKEN_ENV = "TRADETRON_TOKEN_USGOLD"
TT_WEBHOOK_URL = "https://api.tradetron.tech/api?"
DRY = os.getenv("USGOLD_DRY") == "1"
FORCE = os.getenv("USGOLD_FORCE") == "1"
NY, IST = ZoneInfo("America/New_York"), ZoneInfo("Asia/Kolkata")
NG_MIN, NG_MAX, NG_STEP = 0.5, 3.0, 0.5
SESSION = (940, 1540)        # New York time. Tradetron only accepts signals while the US market is open.
EQUITY, GOLD = "SPY", "GLD"


# ----------------------------------------------------------------------------- helpers
def now_ny():
    if os.getenv("USGOLD_NOW"):                      # testing only: "2026-11-10 09:45" (New York time)
        return datetime.strptime(os.getenv("USGOLD_NOW"), "%Y-%m-%d %H:%M").replace(tzinfo=NY)
    return datetime.now(NY)


def both(now):
    return f"{now:%a %d %b %Y}, {now:%H:%M} New York ({now.astimezone(IST):%H:%M} IST)"


def load_state():
    return json.loads(STATE_FILE.read_text())


def save_state(st):
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps(st, indent=2, sort_keys=True))


def telegram(text):
    """Send to every chat in TELEGRAM_CHAT_ID (comma-separated). True if all got it."""
    tok, chats = os.getenv("TELEGRAM_BOT_TOKEN", "").strip(), os.getenv("TELEGRAM_CHAT_ID", "").strip()
    print("\n----- Telegram -----\n" + text + "\n--------------------")
    if DRY:
        return True
    ids = [c.strip() for c in chats.split(",") if c.strip()]
    if not tok or not ids:
        print("Telegram secrets missing — message not sent")
        return False
    ok = True
    for chat in ids:
        sent = _telegram_one(tok, chat, text)
        print(f"chat ending ...{chat[-4:]}: {'sent' if sent else 'NOT sent'}")
        ok = ok and sent
    return ok


def _telegram_one(tok, chat, text):
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
    """Send [(variable, value), ...] to the Tradetron strategy in one call. True on success."""
    if DRY:
        for k, v in pairs:
            print(f"   [dry] {k} = {v}")
        return True
    token = os.getenv(TOKEN_ENV, "").strip()
    if not token:
        print(f"{TOKEN_ENV} secret is missing")
        return False
    body = {"auth-token": token}
    for i, (k, v) in enumerate(pairs):
        sfx = "" if i == 0 else str(i)
        body["key" + sfx] = str(k)
        body["value" + sfx] = str(v)
    data = json.dumps(body).encode()
    for attempt in range(1, 4):
        try:
            req = urllib.request.Request(TT_WEBHOOK_URL, data=data, method="POST",
                                         headers={"Content-Type": "application/json", "User-Agent": "tt-webhook/1.0"})
            with urllib.request.urlopen(req, timeout=20) as resp:
                text = resp.read().decode(errors="replace")
                if resp.status == 200 and '"success":false' not in text.replace(" ", "").lower():
                    return True
                print(f"   Tradetron replied {resp.status}: {text[:200]}")
        except urllib.error.HTTPError as e:
            print(f"   Tradetron HTTP error {e.code} (attempt {attempt}/3)")
        except Exception as e:
            print(f"   Could not reach Tradetron: {type(e).__name__} (attempt {attempt}/3)")
        time.sleep(6 * attempt)
    return False


def in_session(now):
    hhmm = now.hour * 100 + now.minute
    return now.weekday() <= 4 and SESSION[0] <= hhmm <= SESSION[1]


def get_prices(now):
    """Returns (spy, gld, date of the latest price) from Yahoo's 1-minute data (actual traded prices)."""
    if os.getenv("USGOLD_PRICES"):                   # testing only: "752.69,379.53"
        a, b = os.getenv("USGOLD_PRICES").split(",")
        return float(a), float(b), now.date()
    import yfinance as yf
    out, last = {}, None
    for sym in (EQUITY, GOLD):
        df = None
        for attempt in range(3):
            try:
                df = yf.Ticker(sym).history(period="5d", interval="1m", auto_adjust=False)
                if df is not None and len(df):
                    break
            except Exception as e:
                print(f"   price error for {sym}: {type(e).__name__}")
            time.sleep(4)
        if df is None or not len(df):
            return None, None, None
        c = df["Close"].dropna()
        out[sym] = float(c.iloc[-1])
        d = c.index[-1].tz_convert(NY).date()
        last = d if last is None else min(last, d)
    return out[EQUITY], out[GOLD], last


def gold_weight(ratio):
    return min(1.0, max(0.0, (ratio - NG_MIN) / (NG_MAX - NG_MIN)))


def split_line(amount, ps, pg, w):
    """How an amount divides at the current weight: dollars and approximate fractional shares."""
    a_g, a_s = amount * w, amount * (1 - w)
    return (f"• {EQUITY}: about ${a_s:,.0f} = {a_s / ps:.4f} shares at ${ps:,.2f}\n"
            f"• {GOLD}: about ${a_g:,.0f} = {a_g / pg:.4f} shares at ${pg:,.2f}")


def levels(st):
    ref = st["ng_ref"]
    return f"next reshuffle if the ratio reaches {ref - NG_STEP:.3f} or {ref + NG_STEP:.3f}"


# ----------------------------------------------------------------------------- modes
def mode_daily(st, now):
    today = now.date().isoformat()
    if st.get("last_daily") == today and not FORCE:
        print("Already checked today — nothing to do.")
        if os.getenv("USGOLD_MANUAL") == "1":
            mode_status(st, now, title="today's check is already done")
        return True
    if not in_session(now):
        print(f"Skipped: US market closed or outside {SESSION[0]}-{SESSION[1]} New York time (now {now:%a %H:%M})")
        return None
    ps, pg, last = get_prices(now)
    if ps is None:
        telegram(f"{NAME} — daily check FAILED: could not read SPY / GLD prices. Nothing was sent to Tradetron. "
                 "Please press 'US Gold: daily check' in GitHub.")
        return False
    if last != now.date():
        print(f"Skipped: latest price is {last}, not today — US market holiday or prices not in yet")
        return None
    ratio = ps / pg
    w = gold_weight(ratio)
    st["last_ratio"], st["last_prices"] = round(ratio, 4), {EQUITY: round(ps, 2), GOLD: round(pg, 2)}
    print(f"{EQUITY} {ps:.2f}  {GOLD} {pg:.2f}  ratio {ratio:.4f}  gold weight {w * 100:.1f}%  ref {st.get('ng_ref')}")

    if st.get("ng_ref") is None:
        st["ng_ref"] = round(ratio, 4)
        st["last_daily"] = today
        cap = float(st.get("capital") or 0)
        telegram(f"{NAME} — {both(now)}\n\n"
                 f"Daily check is now running. SPY/GLD ratio {ratio:.3f}, so the target is "
                 f"{w * 100:.0f}% GLD and {(1 - w) * 100:.0f}% SPY.\n\n"
                 f"FIRST PURCHASE: Tradetron buys on its first day after 9:40 am New York. On ${cap:,.0f} that is:\n"
                 + split_line(cap, ps, pg, w) +
                 "\n\nCopy the exact fractional quantities from Tradetron's two buy orders.\n"
                 f"After this I only write when something has to be done: {levels(st)}.")
        return True

    ref = float(st["ng_ref"])
    if abs(ratio - ref) < NG_STEP:
        print(f"No reshuffle: ratio moved {ratio - ref:+.4f} from {ref} (needs {NG_STEP}). No Telegram message.")
        st["last_daily"] = today
        return True

    if not tt_send([("ng_ref", 0)]):
        telegram(f"{NAME} — a reshuffle is due (ratio {ratio:.3f}, last reshuffle at {ref:.3f}) but I could NOT "
                 "reach Tradetron. Please press 'US Gold: daily check' in GitHub now.")
        return False
    w_old = gold_weight(ref)
    more_gold = ratio > ref
    sell, buy = (EQUITY, GOLD) if more_gold else (GOLD, EQUITY)
    st.setdefault("events", []).append({"date": today, "event": "reshuffle", "from_ratio": ref,
                                        "to_ratio": round(ratio, 4), "gold_pct": round(w * 100, 1)})
    st["ng_ref"] = round(ratio, 4)
    st["last_daily"] = today
    telegram(f"{NAME} — {both(now)}\n\n"
             f"RESHUFFLE TODAY. The SPY/GLD ratio is {ratio:.3f}; it has moved {ratio - ref:+.3f} from {ref:.3f}, "
             f"the ratio at the last purchase / reshuffle.\n\n"
             f"Target changes from {w_old * 100:.0f}% GLD to {w * 100:.0f}% GLD "
             f"({(1 - w) * 100:.0f}% SPY). About {abs(w - w_old) * 100:.0f}% of the account moves from {sell} to {buy}.\n\n"
             f"Tradetron now SELLS part of {sell} first, then BUYS {buy} with the money. "
             "Place the same two orders at your broker: the exact fractional quantities are in Tradetron's orders.\n\n"
             f"New reference ratio {ratio:.3f}: {levels(st)}.")
    return True


def mode_money(st, now, amount, add):
    word, var = ("Top-up", "add_amt") if add else ("Withdrawal", "sub_amt")
    if not in_session(now):
        telegram(f"'{word}' must be pressed while the US market is open "
                 f"(9:40 am – 3:40 pm New York, Mon–Fri; now {now:%a %H:%M} New York). Nothing changed.")
        return False
    if amount < 50:
        telegram(f"{word} amount {amount} looks wrong. Nothing changed.")
        return False
    cap = float(st.get("capital") or 0)
    if not add and amount > cap * 0.95:
        telegram(f"Withdrawal of ${amount:,.0f} is more than 95% of the ${cap:,.0f} put in. Nothing changed. "
                 "To close the account, set the variable kill to 1 in Tradetron.")
        return False
    ps, pg, _ = get_prices(now)
    if not tt_send([(var, round(amount, 2))]):
        telegram(f"'{word}' could NOT reach Tradetron. Please press it again.")
        return False
    st["capital"] = round(cap + amount if add else cap - amount, 2)
    st.setdefault("events", []).append({"date": now.date().isoformat(), "event": "topup" if add else "withdraw",
                                        "amount": round(amount, 2)})
    lines = [f"{NAME} — {both(now)}", ""]
    if ps:
        w = gold_weight(ps / pg)
        if add:
            lines += [f"TOP-UP of ${amount:,.0f} noted. Tradetron buys now, by the current ratio ({w * 100:.0f}% GLD):",
                      split_line(amount, ps, pg, w)]
        else:
            lines += [f"WITHDRAWAL of ${amount:,.0f} noted. Tradetron sells SPY and GLD in the proportion held, "
                      "after using any spare cash. Roughly:", split_line(amount, ps, pg, w)]
    else:
        lines += [f"{word.upper()} of ${amount:,.0f} noted and sent to Tradetron."]
    lines += ["", "Place the same orders at your broker: the exact fractional quantities are in Tradetron's orders.",
              ("Make sure the money is in the broker account first." if add else
               "Take the money out of the broker account after the sales settle."),
              f"Total put in so far: ${st['capital']:,.0f}. The reshuffle levels do not change."]
    telegram("\n".join(lines))
    return True


def mode_status(st, now, title="status"):
    ps, pg, last = get_prices(now)
    lines = [f"{NAME} — {title} — {both(now)}", ""]
    if ps:
        ratio = ps / pg
        lines.append(f"SPY ${ps:,.2f}, GLD ${pg:,.2f} (prices of {last}). Ratio {ratio:.3f} → target "
                     f"{gold_weight(ratio) * 100:.0f}% GLD, {(1 - gold_weight(ratio)) * 100:.0f}% SPY if reshuffled now.")
    if st.get("ng_ref") is None:
        lines.append("The daily check has not run yet.")
    else:
        lines.append(f"Ratio at the last purchase / reshuffle: {float(st['ng_ref']):.3f}; {levels(st)}.")
        lines.append(f"Holdings stay as they are until then (currently set for "
                     f"{gold_weight(float(st['ng_ref'])) * 100:.0f}% GLD at that ratio).")
    lines.append(f"Last daily check: {st.get('last_daily') or 'never'}. Total put in: ${float(st.get('capital') or 0):,.0f}.")
    return telegram("\n".join(lines))


def mode_test(now):
    ok = telegram("TEST MESSAGE — nothing to do, this is only a sample.\n\n"
                  f"{NAME} — {both(now)}\n\n"
                  "RESHUFFLE TODAY. The SPY/GLD ratio is 2.492; it has moved +0.509 from 1.983, "
                  "the ratio at the last purchase / reshuffle.\n\n"
                  "Target changes from 59% GLD to 80% GLD (20% SPY). About 20% of the account moves from SPY to GLD.\n\n"
                  "Tradetron now SELLS part of SPY first, then BUYS GLD with the money. "
                  "Place the same two orders at your broker: the exact fractional quantities are in Tradetron's orders.")
    print("sample sent" if ok else "sample NOT sent")
    return ok


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "daily"
    st = load_state()
    now = now_ny()
    if st.get("active", True) is False and mode != "test":
        print(f"{NAME} is switched off in its record (active = false) — nothing done.")
        return
    print(f"{NAME} — {mode} — {now:%Y-%m-%d %H:%M} New York")
    if mode == "daily":
        ok = mode_daily(st, now)
    elif mode in ("topup", "withdraw"):
        ok = mode_money(st, now, float(str(sys.argv[2]).replace(",", "").replace("$", "")) if len(sys.argv) > 2 else 0,
                        add=(mode == "topup"))
    elif mode == "status":
        ok = mode_status(st, now)
    elif mode == "test":
        if not mode_test(now):
            sys.exit(1)
        return
    else:
        print("unknown mode"); sys.exit(2)
    save_state(st)
    if ok is False:
        sys.exit(1)


if __name__ == "__main__":
    main()
