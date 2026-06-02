import os
import time
import json
import pandas as pd
from tvdatafeed import TvDatafeed, Interval
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

# -------------------------------
# CONFIG
# -------------------------------
FILE_PATH = os.getenv("FILE_PATH", "output/rotation_invest.xlsx")
HTML_PATH = os.getenv("HTML_PATH", "output/index.html")

RAW_SHEET       = "rawdata"
GB_SIGNAL_SHEET = "GB_signal"
SB_SIGNAL_SHEET = "SB_signal"

LOOKBACK    = 20
MAX_RETRIES = 3
RETRY_SLEEP = 5

TV_SYMBOLS = {
    "NiftyBees":   "NIFTYBEES",
    "GoldBees":    "GOLDBEES",
    "SilverBees":  "SILVERBEES",
}

# -------------------------------
# INIT — anonymous mode, no login required
# -------------------------------
tv = TvDatafeed()

# Make sure output folder exists
os.makedirs(os.path.dirname(FILE_PATH), exist_ok=True)

# -------------------------------
# HELPERS
# -------------------------------
def latest_raw_date(raw_df):
    temp = raw_df.copy()
    temp["Date"] = pd.to_datetime(temp["Date"], dayfirst=True, errors="coerce")
    temp = temp.dropna(subset=["Date"])
    return None if temp.empty else temp["Date"].max().normalize()

def fetch_tv_data(symbol):
    last_err = None
    for _ in range(MAX_RETRIES):
        try:
            df = tv.get_hist(
                symbol=symbol,
                exchange="NSE",
                interval=Interval.in_daily,
                n_bars=5000
            )
            if df is None or df.empty:
                raise ValueError(f"No TradingView data for {symbol}")
            df = df.reset_index()
            if "datetime" in df.columns:
                df.rename(columns={"datetime": "Date"}, inplace=True)
            elif "Date" not in df.columns:
                raise ValueError(f"Date column not found for {symbol}")
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.normalize()
            df = df[["Date", "open", "high", "low", "close", "volume"]].copy()
            df.columns = ["Date", "Open", "High", "Low", "Close", "Volume"]
            for c in ["Open", "High", "Low", "Close"]:
                df[c] = pd.to_numeric(df[c], errors="coerce").round(2)
            return df
        except Exception as e:
            last_err = e
            time.sleep(RETRY_SLEEP)
    raise ValueError(f"Failed to fetch {symbol} after {MAX_RETRIES} retries: {last_err}")

def fmt2(x):
    return round(float(x), 2) if pd.notna(x) else None

def build_signals(df, ratio_col, long_a, long_b, a_open, b_open, a_close, b_close, out_b_close):
    d = df[["Date", a_open, b_open, a_close, b_close, ratio_col]].copy()
    d = d.dropna(subset=[ratio_col]).reset_index(drop=True)
    d["20D_H_Ratio"] = d[ratio_col].rolling(LOOKBACK).max().shift(1)
    d["20D_L_Ratio"] = d[ratio_col].rolling(LOOKBACK).min().shift(1)
    signals = []
    holding = None
    for i, row in d.iterrows():
        if pd.isna(row["20D_H_Ratio"]) or pd.isna(row["20D_L_Ratio"]):
            continue
        if row[ratio_col] > row["20D_H_Ratio"] and holding != long_a:
            holding = long_a
            signals.append((i, row["Date"], f"BUY {long_a}"))
        elif row[ratio_col] < row["20D_L_Ratio"] and holding != long_b:
            holding = long_b
            signals.append((i, row["Date"], f"BUY {long_b}"))
    rows = []
    for k, (sig_idx, sig_date, buy_signal) in enumerate(signals):
        nxt = d[d.index > sig_idx]
        if not nxt.empty:
            entry_row   = nxt.iloc[0]
            entry_date  = entry_row["Date"]
            entry_price = entry_row[a_open] if buy_signal.endswith(long_a) else entry_row[b_open]
        else:
            entry_date  = pd.NaT
            entry_price = None
        if k + 1 < len(signals):
            ni  = signals[k + 1][0]
            nx2 = d[d.index > ni]
            if not nx2.empty:
                exit_row   = nx2.iloc[0]
                exit_date  = exit_row["Date"]
                exit_price = exit_row[a_open] if buy_signal.endswith(long_a) else exit_row[b_open]
            else:
                exit_date  = pd.NaT
                exit_price = None
        else:
            exit_date  = pd.NaT
            exit_price = None
        returns = None
        if entry_price is not None and exit_price is not None and entry_price != 0:
            returns = round(((exit_price - entry_price) / entry_price) * 100, 2)
        rows.append({
            "Date":         sig_date,
            "Buy Signal":   buy_signal,
            "Entry_Date":   entry_date  if pd.notna(entry_date)  else None,
            "Entry_Price":  fmt2(entry_price),
            "Exit_Date":    exit_date   if pd.notna(exit_date)   else None,
            "Exit_Price":   fmt2(exit_price),
            "Returns":      f"{returns}%" if returns is not None else None,
            "20D_H_Ratio":  round(float(d.loc[sig_idx, "20D_H_Ratio"]), 4),
            "20D_L_Ratio":  round(float(d.loc[sig_idx, "20D_L_Ratio"]), 4),
            "NB_Close":     fmt2(d.loc[sig_idx, a_close]),
            out_b_close:    fmt2(d.loc[sig_idx, b_close]),
        })
    return pd.DataFrame(rows)

def apply_sheet_formatting(ws, date_cols):
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.column in date_cols and cell.value is not None:
                cell.number_format = "dd-mm-yy"
            elif isinstance(cell.value, (int, float)):
                cell.number_format = "0.00"
    for col in range(1, ws.max_column + 1):
        letter   = get_column_letter(col)
        max_len  = 0
        for r in range(1, min(ws.max_row, 500) + 1):
            v = ws.cell(r, col).value
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[letter].width = max(12, min(max_len + 2, 20))

def safe(v):
    return round(float(v), 4) if pd.notna(v) else None

def generate_html(raw_df, gb_sig, sb_sig):
    chart = raw_df[['Date','NB/GB','20D_H_NB_GB','20D_L_NB_GB','NB/SB','20D_H_NB_SB','20D_L_NB_SB']].dropna(subset=['NB/GB']).tail(70).copy()
    chart['Date'] = chart['Date'].dt.strftime('%d-%m-%y')
    lv = raw_df.dropna(subset=['20D_H_NB_GB','20D_L_NB_GB']).iloc[-1]
    nb_gb   = round(float(lv['NB/GB']), 4)
    nb_sb   = round(float(lv['NB/SB']), 4)   if pd.notna(lv['NB/SB'])   else 'N/A'
    h_nbgb  = round(float(lv['20D_H_NB_GB']), 4)
    l_nbgb  = round(float(lv['20D_L_NB_GB']), 4)
    h_nbsb  = round(float(lv['20D_H_NB_SB']), 4) if pd.notna(lv['20D_H_NB_SB']) else 'N/A'
    l_nbsb  = round(float(lv['20D_L_NB_SB']), 4) if pd.notna(lv['20D_L_NB_SB']) else 'N/A'
    latest_date = lv['Date'].strftime('%d-%m-%y')
    holding_gb  = gb_sig.iloc[-1]['Buy Signal'].replace('BUY ', '') if len(gb_sig) else 'N/A'
    holding_sb  = sb_sig.iloc[-1]['Buy Signal'].replace('BUY ', '') if len(sb_sig) else 'N/A'
    def sig_color(s):
        return '#1DA1F2' if 'NiftyBees' in s else '#0A2885' if 'GoldBees' in s else '#2F80ED'
    def last5_html(df, close_col):
        rows = ''
        for _, r in df.tail(5).iloc[::-1].iterrows():
            color = sig_color(r['Buy Signal'])
            ret_color = '#16a34a' if r['Returns'] and '-' not in str(r['Returns']) else '#ef4444'
            rows += f"<tr><td>{r['Date'].strftime('%d-%m-%y') if hasattr(r['Date'],'strftime') else r['Date']}</td><td><span class='badge' style='background:{color}'>{r['Buy Signal']}</span></td><td>{r['Entry_Price'] if pd.notna(r['Entry_Price']) else '&#8212;'}</td><td>{r['Exit_Price'] if pd.notna(r['Exit_Price']) else '&#8212;'}</td><td style='color:{ret_color};font-weight:700'>{r['Returns'] if pd.notna(r['Returns']) else '&#8212;'}</td><td>{r['20D_H_Ratio']}</td><td>{r['20D_L_Ratio']}</td></tr>"
        return rows
    dates  = chart['Date'].tolist()
    nbgb_v = [safe(x) for x in chart['NB/GB']]
    h1_v   = [safe(x) for x in chart['20D_H_NB_GB']]
    l1_v   = [safe(x) for x in chart['20D_L_NB_GB']]
    nbsb_v = [safe(x) for x in chart['NB/SB']]
    h2_v   = [safe(x) for x in chart['20D_H_NB_SB']]
    l2_v   = [safe(x) for x in chart['20D_L_NB_SB']]
    gb_rows = last5_html(gb_sig, 'GB_Close')
    sb_rows = last5_html(sb_sig, 'SB_Close')
    html = f"""<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<meta name='viewport' content='width=device-width, initial-scale=1.0'>
<title>Rotation Strategy Dashboard</title>
<script src='https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js'></script>
<style>
:root{{
 --bg:#edf5ff;--panel:rgba(255,255,255,.86);--text:#0f1f3a;--muted:#5f708a;
 --tw:#1DA1F2;--tech:#0A2885;--tech2:#01147C;--good:#16a34a;--bad:#ef4444;
 --shadow:0 14px 40px rgba(8,20,60,.12);
}}
*{{box-sizing:border-box;}}
body{{margin:0;font-family:Segoe UI,Inter,Arial,sans-serif;
 background:radial-gradient(circle at 16% 18%,rgba(29,161,242,.18),transparent 26%),
 radial-gradient(circle at 82% 20%,rgba(10,40,133,.18),transparent 22%),
 linear-gradient(135deg,#f7fbff 0%,#e8f1fb 100%);color:var(--text);}}
.container{{max-width:1500px;margin:0 auto;padding:22px;}}
.hero{{position:relative;overflow:hidden;border-radius:28px;padding:28px;color:#fff;box-shadow:var(--shadow);
 background:linear-gradient(135deg,var(--tech2) 0%,var(--tech) 42%,var(--tw) 100%);}}
.hero:before,.hero:after{{content:'';position:absolute;border-radius:50%;background:rgba(255,255,255,.12);}}
.hero:before{{width:240px;height:240px;right:-70px;bottom:-70px;}}
.hero:after{{width:140px;height:140px;right:140px;top:34px;background:rgba(255,255,255,.08);}}
.hero h1{{font-size:2.1rem;margin:0 0 8px;}}
.hero p{{margin:0;color:rgba(255,255,255,.93);}}
.pills{{display:flex;flex-wrap:wrap;gap:12px;margin-top:18px;}}
.pill{{padding:10px 16px;border-radius:999px;font-weight:800;background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.18);backdrop-filter:blur(8px);}}
.grid{{display:grid;gap:16px;}}
.kpis{{grid-template-columns:repeat(6,1fr);margin-top:16px;}}
@media(max-width:1100px){{.kpis{{grid-template-columns:repeat(2,1fr);}}.charts{{grid-template-columns:1fr;}}}}
.card,.chart-card,.table-wrap{{background:var(--panel);border:1px solid rgba(255,255,255,.6);border-radius:20px;box-shadow:var(--shadow);backdrop-filter:blur(10px);}}
.card{{padding:18px;}}
.label{{color:var(--muted);font-size:.78rem;text-transform:uppercase;letter-spacing:.08em;font-weight:800;margin-bottom:10px;}}
.value{{font-size:1.7rem;font-weight:900;line-height:1;}}
.value.tw{{color:var(--tw);}}.value.tech{{color:var(--tech);}}.value.bad{{color:var(--bad);}}
.small{{color:var(--muted);font-size:.85rem;margin-top:8px;}}
.section-title{{margin:24px 0 12px;font-size:1.06rem;font-weight:900;}}
.charts{{grid-template-columns:1fr 1fr;}}
.chart-card{{padding:16px;}}
.chart-head{{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;}}
.badge-soft{{display:inline-block;padding:6px 10px;border-radius:999px;background:linear-gradient(90deg,rgba(29,161,242,.12),rgba(10,40,133,.12));color:var(--tech2);font-size:.78rem;font-weight:900;border:1px solid #d8e7f7;}}
canvas{{max-height:240px;}}
.table-wrap{{overflow:hidden;margin-top:16px;}}
.table-title{{padding:16px 18px;font-weight:900;background:linear-gradient(90deg,rgba(29,161,242,.11),rgba(10,40,133,.11));border-bottom:1px solid #e6eef8;}}
table{{width:100%;border-collapse:collapse;}}
th,td{{padding:12px 10px;text-align:center;border-bottom:1px solid #eef3f8;font-size:.86rem;}}
th{{position:sticky;top:0;background:#f8fbff;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;font-size:.72rem;}}
tr:hover td{{background:#fafcff;}}
.badge{{display:inline-block;padding:6px 10px;border-radius:999px;color:#fff;font-size:.78rem;font-weight:800;}}
.holding-bar{{display:flex;gap:16px;justify-content:center;margin:16px 0;flex-wrap:wrap;}}
.holding-tag{{background:var(--panel);border:1px solid rgba(255,255,255,.6);border-radius:12px;padding:12px 24px;font-size:.9rem;box-shadow:var(--shadow);}}
.holding-tag span{{font-weight:800;font-size:1.05rem;}}
.footer-note{{margin-top:14px;color:var(--muted);font-size:.84rem;}}
</style>
</head>
<body>
<div class='container'>
  <div class='hero'>
    <h1>&#x1F504; Rotation Strategy Dashboard</h1>
    <p>NiftyBees &#x2F; GoldBees &#x2F; SilverBees &nbsp;&#x7C;&nbsp; As of {latest_date}</p>
    <div class='pills'>
      <div class='pill'>Twitter Blue: #1DA1F2</div>
      <div class='pill'>Tech Blue: #0A2885 &#x2F; #01147C</div>
      <div class='pill'>Updated Daily at 4 PM IST</div>
    </div>
  </div>
  <div class='holding-bar'>
    <div class='holding-tag'>Currently Holding (NB&#x2F;GB): <span style='color:{sig_color(holding_gb)}'>{holding_gb}</span></div>
    <div class='holding-tag'>Currently Holding (NB&#x2F;SB): <span style='color:{sig_color(holding_sb)}'>{holding_sb}</span></div>
  </div>
  <div class='grid kpis'>
    <div class='card'><div class='label'>NB / GB</div><div class='value tw'>{nb_gb}</div><div class='small'>Current ratio</div></div>
    <div class='card'><div class='label'>20D High NB/GB</div><div class='value tech'>{h_nbgb}</div><div class='small'>Breakout level</div></div>
    <div class='card'><div class='label'>20D Low NB/GB</div><div class='value bad'>{l_nbgb}</div><div class='small'>Breakdown level</div></div>
    <div class='card'><div class='label'>NB / SB</div><div class='value tw'>{nb_sb}</div><div class='small'>Current ratio</div></div>
    <div class='card'><div class='label'>20D High NB/SB</div><div class='value tech'>{h_nbsb}</div><div class='small'>Breakout level</div></div>
    <div class='card'><div class='label'>20D Low NB/SB</div><div class='value bad'>{l_nbsb}</div><div class='small'>Breakdown level</div></div>
  </div>
  <div class='section-title'>Visual Summary</div>
  <div class='grid charts'>
    <div class='chart-card'><div class='chart-head'><strong>NB / GB Ratio</strong><span class='badge-soft'>Last 70 days</span></div><canvas id='c1'></canvas></div>
    <div class='chart-card'><div class='chart-head'><strong>NB / SB Ratio</strong><span class='badge-soft'>Last 70 days</span></div><canvas id='c2'></canvas></div>
  </div>
  <div class='section-title'>Last 5 Events</div>
  <div class='table-wrap'>
    <div class='table-title'>NB / GB Events</div>
    <table><thead><tr><th>Date</th><th>Signal</th><th>Entry</th><th>Exit</th><th>Returns</th><th>20D High</th><th>20D Low</th></tr></thead>
    <tbody>{gb_rows}</tbody></table>
  </div>
  <div class='table-wrap'>
    <div class='table-title'>NB / SB Events</div>
    <table><thead><tr><th>Date</th><th>Signal</th><th>Entry</th><th>Exit</th><th>Returns</th><th>20D High</th><th>20D Low</th></tr></thead>
    <tbody>{sb_rows}</tbody></table>
  </div>
  <div class='footer-note'>Auto-refreshed daily at 4 PM IST via GitHub Actions &#x2192; Cloudflare Pages</div>
</div>
<script>
const labels={json.dumps(dates)};
const nbgb={json.dumps(nbgb_v)};const h1={json.dumps(h1_v)};const l1={json.dumps(l1_v)};
const nbsb={json.dumps(nbsb_v)};const h2={json.dumps(h2_v)};const l2={json.dumps(l2_v)};
function makeChart(id,main,hi,lo,color){{
  new Chart(document.getElementById(id),{{type:'line',
    data:{{labels,datasets:[
      {{label:'Current Ratio',data:main,borderColor:color,borderWidth:3,pointRadius:0,tension:.28,fill:false}},
      {{label:'20D High',data:hi,borderColor:'#01147C',borderWidth:2,pointRadius:0,borderDash:[6,4]}},
      {{label:'20D Low',data:lo,borderColor:'#ef4444',borderWidth:2,pointRadius:0,borderDash:[6,4]}}
    ]}},
    options:{{responsive:true,plugins:{{legend:{{labels:{{color:'#5f708a'}}}}}},
      scales:{{x:{{ticks:{{color:'#6b7b92'}},grid:{{color:'#e6eef8'}}}},y:{{ticks:{{color:'#6b7b92'}},grid:{{color:'#e6eef8'}}}}}}}}
  );
}}
makeChart('c1',nbgb,h1,l1,'#1DA1F2');
makeChart('c2',nbsb,h2,l2,'#0A2885');
</script>
</body></html>"""
    return html

# -------------------------------
# LOAD or CREATE WORKBOOK
# -------------------------------
if os.path.exists(FILE_PATH):
    wb     = load_workbook(FILE_PATH)
    raw_df = pd.read_excel(FILE_PATH, sheet_name=RAW_SHEET)
else:
    wb     = Workbook()
    wb.create_sheet(RAW_SHEET, 0)
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    raw_df = pd.DataFrame()

today     = pd.Timestamp.today().normalize()
last_date = latest_raw_date(raw_df) if not raw_df.empty else None
needs_update = (last_date is None) or (last_date < today)

print(f"Today: {today.strftime('%d-%m-%y')}")
print(f"Latest raw data: {last_date.strftime('%d-%m-%y') if last_date is not None else 'None'}")
print(f"Needs update: {needs_update}")

# -------------------------------
# UPDATE RAWDATA IF NEEDED
# -------------------------------
if needs_update:
    try:
        nb = fetch_tv_data(TV_SYMBOLS["NiftyBees"]).rename(columns={
            "Open":"NB_open","High":"NB_high","Low":"NB_low","Close":"NB_close","Volume":"NB_volume"
        })[["Date","NB_open","NB_high","NB_low","NB_close","NB_volume"]]
        gb = fetch_tv_data(TV_SYMBOLS["GoldBees"]).rename(columns={
            "Open":"GB_open","High":"GB_high","Low":"GB_low","Close":"GB_close","Volume":"GB_volume"
        })[["Date","GB_open","GB_high","GB_low","GB_close","GB_volume"]]
        sb = fetch_tv_data(TV_SYMBOLS["SilverBees"]).rename(columns={
            "Open":"SB_open","High":"SB_high","Low":"SB_low","Close":"SB_close","Volume":"SB_volume"
        })[["Date","SB_open","SB_high","SB_low","SB_close","SB_volume"]]
        raw_updated = nb.merge(gb, on="Date", how="outer").merge(sb, on="Date", how="outer")
        raw_updated = raw_updated.sort_values("Date").reset_index(drop=True)
        raw_updated["NB/GB"] = (raw_updated["NB_close"] / raw_updated["GB_close"]).round(4)
        raw_updated["NB/SB"] = (raw_updated["NB_close"] / raw_updated["SB_close"]).round(4)
        if RAW_SHEET in wb.sheetnames:
            del wb[RAW_SHEET]
        ws_raw = wb.create_sheet(RAW_SHEET, 0)
        for c_idx, h in enumerate(raw_updated.columns, 1):
            ws_raw.cell(1, c_idx, h)
        for r_idx, row in enumerate(raw_updated.itertuples(index=False), 2):
            for c_idx, val in enumerate(row, 1):
                cell = ws_raw.cell(r_idx, c_idx)
                if c_idx == 1 and pd.notna(val):
                    cell.value = pd.to_datetime(val).to_pydatetime()
                    cell.number_format = "dd-mm-yy"
                elif isinstance(val, (int, float)) and pd.notna(val):
                    if c_idx in [2,3,4,5,7,8,9,10,12,13,14,15]:
                        cell.value = round(float(val), 2)
                        cell.number_format = "0.00"
                    elif c_idx in [17,18]:
                        cell.value = round(float(val), 4)
                        cell.number_format = "0.0000"
                    else:
                        cell.value = val
                else:
                    cell.value = val
        apply_sheet_formatting(ws_raw, date_cols={1})
        print("Raw data updated successfully.")
    except Exception as e:
        print(f"Update failed, using existing raw data: {e}")
        raw_updated = raw_df.copy()
else:
    raw_updated = raw_df.copy()

raw_updated["Date"] = pd.to_datetime(raw_updated["Date"], dayfirst=True, errors="coerce")
raw_updated = raw_updated.sort_values("Date").reset_index(drop=True)

raw_updated["NB/GB"] = raw_updated["NB_close"] / raw_updated["GB_close"]
raw_updated["NB/SB"] = raw_updated["NB_close"] / raw_updated["SB_close"]
raw_updated["20D_H_NB_GB"] = raw_updated["NB/GB"].rolling(20).max().shift(1)
raw_updated["20D_L_NB_GB"] = raw_updated["NB/GB"].rolling(20).min().shift(1)
raw_updated["20D_H_NB_SB"] = raw_updated["NB/SB"].rolling(20).max().shift(1)
raw_updated["20D_L_NB_SB"] = raw_updated["NB/SB"].rolling(20).min().shift(1)

gb_signals = build_signals(raw_updated,"NB/GB","NiftyBees","GoldBees","NB_open","GB_open","NB_close","GB_close","GB_Close")
sb_signals = build_signals(raw_updated,"NB/SB","NiftyBees","SilverBees","NB_open","SB_open","NB_close","SB_close","SB_Close")

for sheet_name, sig_df in [(GB_SIGNAL_SHEET, gb_signals),(SB_SIGNAL_SHEET, sb_signals)]:
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    for c_idx, h in enumerate(sig_df.columns, 1):
        ws.cell(1, c_idx, h)
    for r_idx, row in enumerate(sig_df.itertuples(index=False), 2):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(r_idx, c_idx)
            if c_idx in [1,3,5] and val is not None:
                try:
                    cell.value = pd.to_datetime(val).to_pydatetime()
                    cell.number_format = "dd-mm-yy"
                except:
                    cell.value = val
            elif isinstance(val, str) and val.endswith("%"):
                cell.value = val
            elif isinstance(val, (int, float)) and pd.notna(val):
                if c_idx in [4,6,10,11]:
                    cell.value = round(float(val), 2)
                    cell.number_format = "0.00"
                elif c_idx in [8,9]:
                    cell.value = round(float(val), 4)
                    cell.number_format = "0.0000"
                else:
                    cell.value = round(float(val), 2)
                    cell.number_format = "0.00"
            else:
                cell.value = val
    apply_sheet_formatting(ws, date_cols={1,3,5})

if RAW_SHEET in wb.sheetnames:
    wb._sheets = [wb[RAW_SHEET]] + [wb[s] for s in wb.sheetnames if s != RAW_SHEET]

wb.save(FILE_PATH)
print(f"Workbook saved: {FILE_PATH}")
print(f"GB signals: {len(gb_signals)}")
print(f"SB signals: {len(sb_signals)}")

html_content = generate_html(raw_updated, gb_signals, sb_signals)
with open(HTML_PATH, 'w', encoding='utf-8') as f:
    f.write(html_content)
print(f"Dashboard saved: {HTML_PATH}")
