#!/usr/bin/env python3
"""
NiftyBees / GoldBees / SilverBees Rotation Strategy Bot
Runs daily at 4 PM IST via GitHub Actions
Stores state in Cloudflare R2, sends alert to Telegram
"""

import os
import json
import boto3
import requests
import yfinance as yf
import pandas as pd
from datetime import datetime, date
import pytz

# ── CONFIG (all from GitHub Secrets) ──────────────────────────────────────────────
TELEGRAM_TOKEN  = os.environ["TELEGRAM_TOKEN"]
TELEGRAM_CHAT   = os.environ["TELEGRAM_CHAT_ID"]

R2_ACCESS_KEY   = os.environ["R2_ACCESS_KEY"]
R2_SECRET_KEY   = os.environ["R2_SECRET_KEY"]
R2_ENDPOINT     = os.environ["R2_ENDPOINT"]        # https://<account>.r2.cloudflarestorage.com
R2_BUCKET       = os.environ.get("R2_BUCKET", "rotation-signals")

IST = pytz.timezone("Asia/Kolkata")

# ── TICKERS ─────────────────────────────────────────────────────────────────────────────
TICKERS = {
    "NIFTYBEES": "NIFTYBEES.NS",
    "GOLDBEES":  "GOLDBEES.NS",
    "SILVERBEES": "SILVERBEES.NS",
}

# ── R2 CLIENT ───────────────────────────────────────────────────────────────────────────
def get_r2():
    return boto3.client(
        "s3",
        endpoint_url=R2_ENDPOINT,
        aws_access_key_id=R2_ACCESS_KEY,
        aws_secret_access_key=R2_SECRET_KEY,
        region_name="auto",
    )

def r2_read_json(key: str, default: dict) -> dict:
    try:
        obj = get_r2().get_object(Bucket=R2_BUCKET, Key=key)
        return json.loads(obj["Body"].read())
    except Exception:
        return default

def r2_write_json(key: str, data: dict):
    get_r2().put_object(
        Bucket=R2_BUCKET, Key=key,
        Body=json.dumps(data, indent=2, default=str),
        ContentType="application/json",
    )

def r2_append_csv(key: str, row: dict):
    try:
        obj = get_r2().get_object(Bucket=R2_BUCKET, Key=key)
        existing = obj["Body"].read().decode()
    except Exception:
        existing = ""
    df_new = pd.DataFrame([row])
    if existing.strip():
        csv_out = existing + df_new.to_csv(index=False, header=False)
    else:
        csv_out = df_new.to_csv(index=False, header=True)
    get_r2().put_object(
        Bucket=R2_BUCKET, Key=key,
        Body=csv_out.encode(),
        ContentType="text/csv",
    )

# ── FETCH PRICES ─────────────────────────────────────────────────────────────────────────
def fetch_prices() -> dict:
    prices = {}
    for name, ticker in TICKERS.items():
        try:
            df = yf.download(ticker, period="5d", interval="1d", progress=False)
            prices[name] = round(float(df["Close"].dropna().iloc[-1]), 2)
        except Exception as e:
            print(f"WARNING: Could not fetch {ticker}: {e}")
            prices[name] = None
    return prices

# ── SIGNAL LOGIC ─────────────────────────────────────────────────────────────────────────
def compute_signals(prices: dict, prev_state: dict) -> dict:
    nb = prices.get("NIFTYBEES")
    gb = prices.get("GOLDBEES")
    sb = prices.get("SILVERBEES")

    signals = {}

    # S1: NiftyBees vs GoldBees
    if nb and gb:
        ratio_ng = round(nb / gb, 4)
        prev_pos = prev_state.get("s1_position", "NIFTYBEES")
        if ratio_ng > 2.5 and prev_pos != "NIFTYBEES":
            signals["S1"] = {"action": "BUY NIFTYBEES", "ratio": ratio_ng, "new_pos": "NIFTYBEES"}
        elif ratio_ng < 2.0 and prev_pos != "GOLDBEES":
            signals["S1"] = {"action": "BUY GOLDBEES", "ratio": ratio_ng, "new_pos": "GOLDBEES"}
        else:
            signals["S1"] = {"action": "HOLD", "ratio": ratio_ng, "new_pos": prev_pos}
    else:
        signals["S1"] = {"action": "NO DATA", "ratio": None, "new_pos": prev_state.get("s1_position", "?")}

    # S2: NiftyBees vs SilverBees
    if nb and sb:
        ratio_ns = round(nb / sb, 4)
        prev_pos2 = prev_state.get("s2_position", "NIFTYBEES")
        if ratio_ns > 4.0 and prev_pos2 != "NIFTYBEES":
            signals["S2"] = {"action": "BUY NIFTYBEES", "ratio": ratio_ns, "new_pos": "NIFTYBEES"}
        elif ratio_ns < 3.0 and prev_pos2 != "SILVERBEES":
            signals["S2"] = {"action": "BUY SILVERBEES", "ratio": ratio_ns, "new_pos": "SILVERBEES"}
        else:
            signals["S2"] = {"action": "HOLD", "ratio": ratio_ns, "new_pos": prev_pos2}
    else:
        signals["S2"] = {"action": "NO DATA", "ratio": None, "new_pos": prev_state.get("s2_position", "?")}

    # S3: GoldBees vs SilverBees
    if gb and sb:
        ratio_gs = round(gb / sb, 4)
        prev_pos3 = prev_state.get("s3_position", "GOLDBEES")
        if ratio_gs > 1.8 and prev_pos3 != "GOLDBEES":
            signals["S3"] = {"action": "BUY GOLDBEES", "ratio": ratio_gs, "new_pos": "GOLDBEES"}
        elif ratio_gs < 1.3 and prev_pos3 != "SILVERBEES":
            signals["S3"] = {"action": "BUY SILVERBEES", "ratio": ratio_gs, "new_pos": "SILVERBEES"}
        else:
            signals["S3"] = {"action": "HOLD", "ratio": ratio_gs, "new_pos": prev_pos3}
    else:
        signals["S3"] = {"action": "NO DATA", "ratio": None, "new_pos": prev_state.get("s3_position", "?")}

    return signals

# ── TELEGRAM ─────────────────────────────────────────────────────────────────────────────
def send_telegram(msg: str):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    resp = requests.post(url, json={
        "chat_id": TELEGRAM_CHAT,
        "text": msg,
        "parse_mode": "HTML",
    }, timeout=15)
    resp.raise_for_status()

def build_message(prices, signals, today_str) -> str:
    nb = prices.get("NIFTYBEES", "N/A")
    gb = prices.get("GOLDBEES", "N/A")
    sb = prices.get("SILVERBEES", "N/A")

    new_signals = [k for k, v in signals.items() if v["action"] not in ("HOLD", "NO DATA")]
    header = "🚨 <b>NEW SIGNAL FIRED!</b>\n\n" if new_signals else "📊 <b>Daily Rotation Update</b>\n\n"

    price_block = (
        f"<b>Prices ({today_str})</b>\n"
        f"  NiftyBees : ₹{nb}\n"
        f"  GoldBees  : ₹{gb}\n"
        f"  SilverBees: ₹{sb}\n\n"
    )

    signal_block = "<b>Signal Status</b>\n"
    icons = {"BUY NIFTYBEES": "🟢", "BUY GOLDBEES": "🟡", "BUY SILVERBEES": "⚪", "HOLD": "⏸", "NO DATA": "❌"}
    for k, v in signals.items():
        icon = icons.get(v["action"], "•")
        ratio_str = f" | Ratio: {v['ratio']}" if v["ratio"] else ""
        signal_block += f"  {icon} <b>{k}</b>: {v['action']}{ratio_str}\n"
        signal_block += f"     Position → <b>{v['new_pos']}</b>\n"

    footer = ""
    if new_signals:
        footer = "\n⚡ <i>Execute switch at market open tomorrow (NSE 9:15 AM)</i>"

    return header + price_block + signal_block + footer

# ── MAIN ──────────────────────────────────────────────────────────────────────────────
def main():
    today_str = datetime.now(IST).strftime("%d-%b-%Y %I:%M %p IST")
    print(f"Running rotation bot at {today_str}")

    prev_state = r2_read_json("signal_state.json", {})
    print(f"Previous state: {prev_state}")

    prices = fetch_prices()
    print(f"Prices: {prices}")

    signals = compute_signals(prices, prev_state)
    print(f"Signals: {signals}")

    new_state = {
        "last_updated": today_str,
        "s1_position": signals["S1"]["new_pos"],
        "s2_position": signals["S2"]["new_pos"],
        "s3_position": signals["S3"]["new_pos"],
        "ratio_ng": signals["S1"].get("ratio"),
        "ratio_ns": signals["S2"].get("ratio"),
        "ratio_gs": signals["S3"].get("ratio"),
        "prices": prices,
    }

    r2_write_json("signal_state.json", new_state)

    log_row = {
        "date": today_str,
        "NIFTYBEES": prices.get("NIFTYBEES"),
        "GOLDBEES":  prices.get("GOLDBEES"),
        "SILVERBEES": prices.get("SILVERBEES"),
        "ratio_ng": signals["S1"].get("ratio"),
        "ratio_ns": signals["S2"].get("ratio"),
        "ratio_gs": signals["S3"].get("ratio"),
        "s1_action": signals["S1"]["action"],
        "s1_position": signals["S1"]["new_pos"],
        "s2_action": signals["S2"]["action"],
        "s2_position": signals["S2"]["new_pos"],
        "s3_action": signals["S3"]["action"],
        "s3_position": signals["S3"]["new_pos"],
    }
    r2_append_csv("signals_log.csv", log_row)

    msg = build_message(prices, signals, today_str)
    send_telegram(msg)
    print("Telegram message sent successfully.")

if __name__ == "__main__":
    main()
