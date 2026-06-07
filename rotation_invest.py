import os
import time
import json
import pandas as pd
import yfinance as yf
from openpyxl import load_workbook, Workbook
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

YF_SYMBOLS = {
    "NiftyBees":  "NIFTYBEES.NS",
    "GoldBees":   "GOLDBEES.NS",
    "SilverBees": "SILVERBEES.NS",
}

os.makedirs(os.path.dirname(FILE_PATH), exist_ok=True)

# -------------------------------
# HELPERS
# -------------------------------
def latest_raw_date(raw_df):
    temp = raw_df.copy()
    temp["Date"] = pd.to_datetime(temp["Date"], dayfirst=True, errors="coerce")
    temp = temp.dropna(subset=["Date"])
    return None if temp.empty else temp["Date"].max().normalize()

def fetch_yf_data(ticker):
    last_err = None
    for _ in range(MAX_RETRIES):
        try:
            df = yf.download(ticker, period="20y", interval="1d", auto_adjust=True, progress=False)
            if df is None or df.empty:
                raise ValueError(f"No data for {ticker}")
            df = df.reset_index()
            # yfinance may return MultiIndex columns — flatten
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = [c[0] for c in df.columns]
            df.rename(columns={"Date": "Date", "Open": "Open", "High": "High",
                                "Low": "Low", "Close": "Close", "Volume": "Volume"}, inplace=True)
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.normalize()
            df = df[["Date", "Open", "High", "Low", "Close", "Volume"]].copy()
            for c in ["Open", "High", "Low", "Close"]:
                df[c] = pd.to_numeric(df[c], errors="coerce").round(2)
            df = df.dropna(subset=["Date", "Close"])
            return df
        except Exception as e:
            last_err = e
            time.sleep(RETRY_SLEEP)
    raise ValueError(f"Failed to fetch {ticker} after {MAX_RETRIES} retries: {last_err}")

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
        letter = get_column_letter(col)
        max_len = 0
        for r in range(1, min(ws.max_row, 500) + 1):
            v = ws.cell(r, col).value
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[letter].width = max(12, min(max_len + 2, 20))

def safe(v):
    return round(float(v), 4) if pd.notna(v) else None

def generate_html(raw_df, gb_sig, sb_sig):

    chart = raw_df[['Date','NB/GB','20D_H_NB_GB','20D_L_NB_GB',
                    'NB/SB','20D_H_NB_SB','20D_L_NB_SB',
                    'NB_close','GB_close','SB_close']].dropna(subset=['NB/GB']).tail(70).copy()

    chart['Date'] = chart['Date'].dt.strftime('%d-%m-%y')

    lv = raw_df.dropna(subset=['20D_H_NB_GB','20D_L_NB_GB']).iloc[-1]

    nb_gb  = round(float(lv['NB/GB']), 4)
    nb_sb  = round(float(lv['NB/SB']), 4)

    h_nbgb = round(float(lv['20D_H_NB_GB']), 4)
    l_nbgb = round(float(lv['20D_L_NB_GB']), 4)

    h_nbsb = round(float(lv['20D_H_NB_SB']), 4)
    l_nbsb = round(float(lv['20D_L_NB_SB']), 4)

    latest_date = lv['Date'].strftime('%d-%m-%y')

    nb_price = round(float(lv['NB_close']),2)
    gb_price = round(float(lv['GB_close']),2)
    sb_price = round(float(lv['SB_close']),2)

    holding_gb = gb_sig.iloc[-1]['Buy Signal'].replace('BUY ', '') if len(gb_sig) else 'N/A'
    holding_sb = sb_sig.iloc[-1]['Buy Signal'].replace('BUY ', '') if len(sb_sig) else 'N/A'

    def sig_color(s):
        return '#2563eb' if 'NiftyBees' in s else '#ca8a04' if 'GoldBees' in s else '#64748b'

    def last5_html(df):

        rows = ''

        for _, r in df.tail(5).iloc[::-1].iterrows():

            color = sig_color(r['Buy Signal'])

            ret_color = '#16a34a' if r['Returns'] and '-' not in str(r['Returns']) else '#ef4444'

            rows += (
                f"<tr>"
                f"<td>{r['Date'].strftime('%d-%m-%y') if hasattr(r['Date'],'strftime') else r['Date']}</td>"
                f"<td><span class='badge' style='background:{color}'>{r['Buy Signal']}</span></td>"
                f"<td>{r['Entry_Price'] if pd.notna(r['Entry_Price']) else '&#8212;'}</td>"
                f"<td>{r['Exit_Price'] if pd.notna(r['Exit_Price']) else '&#8212;'}</td>"
                f"<td style='color:{ret_color};font-weight:700'>{r['Returns'] if pd.notna(r['Returns']) else '&#8212;'}</td>"
                f"<td>{r['20D_H_Ratio']}</td>"
                f"<td>{r['20D_L_Ratio']}</td>"
                f"</tr>"
            )

        return rows

    dates  = chart['Date'].tolist()

    nbgb_v = [round(float(x),4) for x in chart['NB/GB']]
    h1_v   = [round(float(x),4) for x in chart['20D_H_NB_GB']]
    l1_v   = [round(float(x),4) for x in chart['20D_L_NB_GB']]

    nbsb_v = [round(float(x),4) for x in chart['NB/SB']]
    h2_v   = [round(float(x),4) for x in chart['20D_H_NB_SB']]
    l2_v   = [round(float(x),4) for x in chart['20D_L_NB_SB']]

    gb_rows = last5_html(gb_sig)
    sb_rows = last5_html(sb_sig)

    html = f"""
<!DOCTYPE html>
<html lang="en">

<head>

<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title>Rotation Dashboard</title>

<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

<style>

body{{
    margin:0;
    font-family:Segoe UI,Arial,sans-serif;
    background:#f1f5f9;
    color:#0f172a;
}}

.container{{
    max-width:1500px;
    margin:auto;
    padding:20px;
}}

.hero{{
    background:#ffffff;
    border-radius:24px;
    padding:24px;
    margin-bottom:18px;
    box-shadow:0 4px 18px rgba(0,0,0,.08);
}}

.hero h1{{
    margin:0;
    font-size:2rem;
}}

.hero p{{
    margin-top:8px;
    color:#64748b;
}}

.workflow-box{{
    background:#ffffff;
    border-radius:18px;
    padding:18px;
    margin-bottom:18px;
    box-shadow:0 4px 18px rgba(0,0,0,.08);
}}

.progress-wrap{{
    width:100%;
    height:14px;
    background:#e2e8f0;
    border-radius:999px;
    overflow:hidden;
    margin-top:12px;
}}

#workflowProgress{{
    width:0%;
    height:100%;
    background:#2563eb;
    transition:.4s;
}}

#workflowText{{
    margin-top:10px;
    color:#475569;
    font-size:.92rem;
}}

.grid{{
    display:grid;
    gap:16px;
}}

.kpis{{
    grid-template-columns:repeat(9,1fr);
}}

.card{{
    background:#ffffff;
    border-radius:20px;
    padding:18px;
    box-shadow:0 4px 18px rgba(0,0,0,.08);
}}

.label{{
    color:#64748b;
    font-size:.78rem;
    text-transform:uppercase;
    font-weight:700;
}}

.value{{
    font-size:2rem;
    font-weight:800;
    margin-top:10px;
}}

.small{{
    margin-top:8px;
    color:#64748b;
}}

.section-title{{
    margin:28px 0 14px;
    font-size:1.2rem;
    font-weight:800;
}}

.charts{{
    grid-template-columns:1fr 1fr;
}}

.chart-card{{
    background:#ffffff;
    border-radius:22px;
    padding:20px;
    box-shadow:0 4px 18px rgba(0,0,0,.08);
}}

.chart-container{{
    position:relative;
    height:340px;
    width:100%;
}}

canvas{{
    width:100% !important;
    height:100% !important;
}}

.table-wrap{{
    background:#ffffff;
    border-radius:22px;
    overflow:hidden;
    margin-top:18px;
    box-shadow:0 4px 18px rgba(0,0,0,.08);
}}

.table-title{{
    padding:16px 18px;
    font-weight:800;
    background:#eff6ff;
}}

table{{
    width:100%;
    border-collapse:collapse;
}}

th,td{{
    padding:12px;
    border-bottom:1px solid #e2e8f0;
    text-align:center;
}}

th{{
    background:#f8fafc;
    font-size:.75rem;
    color:#64748b;
}}

.badge{{
    color:#fff;
    padding:6px 10px;
    border-radius:999px;
    font-size:.78rem;
    font-weight:700;
}}

@media(max-width:1200px){{
    .kpis{{grid-template-columns:repeat(2,1fr);}}
    .charts{{grid-template-columns:1fr;}}
}}

</style>
</head>

<body>

<div class="container">

<div class="hero">
    <h1>Rotation Strategy Dashboard</h1>
    <p>Updated till {latest_date}</p>
</div>

<div class="workflow-box">

    <div style="font-weight:700;">
        Workflow Status
    </div>

    <div class="progress-wrap">
        <div id="workflowProgress"></div>
    </div>

    <div id="workflowText">
        Idle
    </div>

</div>

<div class="grid kpis">

<div class="card">
    <div class="label">NiftyBees</div>
    <div class="value" style="color:#2563eb;">₹{nb_price}</div>
    <div class="small">Current Price</div>
</div>

<div class="card">
    <div class="label">GoldBees</div>
    <div class="value" style="color:#ca8a04;">₹{gb_price}</div>
    <div class="small">Current Price</div>
</div>

<div class="card">
    <div class="label">SilverBees</div>
    <div class="value" style="color:#64748b;">₹{sb_price}</div>
    <div class="small">Current Price</div>
</div>

<div class="card">
    <div class="label">NB / GB</div>
    <div class="value" style="color:#2563eb;">{nb_gb}</div>
    <div class="small">Current Ratio</div>
</div>

<div class="card">
    <div class="label">20D High NB/GB</div>
    <div class="value" style="color:#16a34a;">{h_nbgb}</div>
    <div class="small">Breakout Level</div>
</div>

<div class="card">
    <div class="label">20D Low NB/GB</div>
    <div class="value" style="color:#ef4444;">{l_nbgb}</div>
    <div class="small">Breakdown Level</div>
</div>

<div class="card">
    <div class="label">NB / SB</div>
    <div class="value" style="color:#2563eb;">{nb_sb}</div>
    <div class="small">Current Ratio</div>
</div>

<div class="card">
    <div class="label">20D High NB/SB</div>
    <div class="value" style="color:#16a34a;">{h_nbsb}</div>
    <div class="small">Breakout Level</div>
</div>

<div class="card">
    <div class="label">20D Low NB/SB</div>
    <div class="value" style="color:#ef4444;">{l_nbsb}</div>
    <div class="small">Breakdown Level</div>
</div>

</div>

<div class="section-title">
Visual Summary
</div>

<div class="grid charts">

<div class="chart-card">

<h3>NB / GB Ratio</h3>

<div class="chart-container">
<canvas id="c1"></canvas>
</div>

</div>

<div class="chart-card">

<h3>NB / SB Ratio</h3>

<div class="chart-container">
<canvas id="c2"></canvas>
</div>

</div>

</div>

<div class="table-wrap">

<div class="table-title">
NB / GB Events
</div>

<table>

<thead>
<tr>
<th>Date</th>
<th>Signal</th>
<th>Entry</th>
<th>Exit</th>
<th>Returns</th>
<th>20D High</th>
<th>20D Low</th>
</tr>
</thead>

<tbody>
{gb_rows}
</tbody>

</table>

</div>

<div class="table-wrap">

<div class="table-title">
NB / SB Events
</div>

<table>

<thead>
<tr>
<th>Date</th>
<th>Signal</th>
<th>Entry</th>
<th>Exit</th>
<th>Returns</th>
<th>20D High</th>
<th>20D Low</th>
</tr>
</thead>

<tbody>
{sb_rows}
</tbody>

</table>

</div>

</div>

<script>

const labels = {json.dumps(dates)}

const nbgb = {json.dumps(nbgb_v)}
const h1 = {json.dumps(h1_v)}
const l1 = {json.dumps(l1_v)}

const nbsb = {json.dumps(nbsb_v)}
const h2 = {json.dumps(h2_v)}
const l2 = {json.dumps(l2_v)}

function makeChart(id,main,hi,lo,color){{

    new Chart(document.getElementById(id),{{

        type:'line',

        data:{{
            labels:labels,
            datasets:[
                {{
                    label:'Ratio',
                    data:main,
                    borderColor:color,
                    borderWidth:3,
                    pointRadius:0,
                    tension:.3
                }},
                {{
                    label:'20D High',
                    data:hi,
                    borderColor:'#16a34a',
                    borderDash:[6,4],
                    borderWidth:2,
                    pointRadius:0
                }},
                {{
                    label:'20D Low',
                    data:lo,
                    borderColor:'#ef4444',
                    borderDash:[6,4],
                    borderWidth:2,
                    pointRadius:0
                }}
            ]
        }},

        options:{{
            responsive:true,
            maintainAspectRatio:false
        }}

    }});

}}

makeChart('c1',nbgb,h1,l1,'#2563eb');
makeChart('c2',nbsb,h2,l2,'#0f172a');

async function fetchWorkflowStatus(){{

    try{{

        const r = await fetch("status.json?t=" + new Date().getTime());

        const s = await r.json();

        document.getElementById("workflowText").innerText = s.message;

        document.getElementById("workflowProgress").style.width = s.progress + "%";

    }}catch(e){{
        console.log(e);
    }}

}}

fetchWorkflowStatus();

setInterval(fetchWorkflowStatus,5000);

</script>

</body>
</html>
"""

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
# FETCH + UPDATE RAWDATA
# -------------------------------
try:
    nb = fetch_yf_data(YF_SYMBOLS["NiftyBees"]).rename(columns={
        "Open":"NB_open","High":"NB_high","Low":"NB_low","Close":"NB_close","Volume":"NB_volume"
    })[["Date","NB_open","NB_high","NB_low","NB_close","NB_volume"]]
    gb = fetch_yf_data(YF_SYMBOLS["GoldBees"]).rename(columns={
        "Open":"GB_open","High":"GB_high","Low":"GB_low","Close":"GB_close","Volume":"GB_volume"
    })[["Date","GB_open","GB_high","GB_low","GB_close","GB_volume"]]
    sb = fetch_yf_data(YF_SYMBOLS["SilverBees"]).rename(columns={
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
    print("Raw data fetched and updated successfully.")
except Exception as e:
    print(f"Fetch failed: {e}")
    if raw_df.empty:
        raise RuntimeError("No existing data and fresh fetch failed — cannot proceed.") from e
    raw_updated = raw_df.copy()
    print("Using existing raw data from workbook.")

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
