import os
import pandas as pd
from datetime import datetime

# 1. Define File Paths (Managed by your GitHub Actions workflow)
excel_path = os.getenv("FILE_PATH", "output/rotation_invest.xlsx")
html_path = os.getenv("HTML_PATH", "output/index.html")

# Create output directory if it doesn't exist locally
os.makedirs(os.path.dirname(html_path), exist_ok=True)

# 2. Read the Excel File (Sourced from Cloudflare R2)
if not os.path.exists(excel_path):
    raise FileNotFoundError(f"Critical Error: Excel ledger not found at {excel_path}")

df = pd.read_excel(excel_path)
df = df.dropna(subset=['NiftyBees_Close', 'GoldBees_Close', 'SilverBees_Close'])

# Grab the absolute last row for current metrics
last_row = df.iloc[-1]

# Extract Current Close Prices and Date from Excel
current_nb = float(last_row['NiftyBees_Close'])
current_gb = float(last_row['GoldBees_Close'])
current_sb = float(last_row['SilverBees_Close'])

# Format the date nicely for the dashboard headers
raw_date = last_row['Date']
if isinstance(raw_date, (datetime, pd.Timestamp)):
    formatted_date = raw_date.strftime('%d-%m-%y')
else:
    formatted_date = str(raw_date)

# 3. Generate Rolling 70-Day Data Arrays for Chart.js
df_70 = df.tail(70).copy()

# Calculate Rolling Ratios and 20-Day High/Low Boundaries
df_70['NB_GB'] = (df_70['NiftyBees_Close'] / df_70['GoldBees_Close']).round(4)
df_70['NB_SB'] = (df_70['NiftyBees_Close'] / df_70['SilverBees_Close']).round(4)

df_70['H1'] = df_70['NB_GB'].rolling(window=20, min_periods=1).max().round(4)
df_70['L1'] = df_70['NB_GB'].rolling(window=20, min_periods=1).min().round(4)
df_70['H2'] = df_70['NB_SB'].rolling(window=20, min_periods=1).max().round(4)
df_70['L2'] = df_70['NB_SB'].rolling(window=20, min_periods=1).min().round(4)

# Format dates for the chart labels
if pd.api.types.is_datetime64_any_dtype(df_70['Date']):
    chart_labels = df_70['Date'].dt.strftime('%d-%m-%y').tolist()
else:
    chart_labels = df_70['Date'].astype(str).tolist()

# Extract arrays as clean lists for Javascript injection
list_nbgb = df_70['NB_GB'].tolist()
list_h1 = df_70['H1'].tolist()
list_l1 = df_70['L1'].tolist()

list_nbsb = df_70['NB_SB'].tolist()
list_h2 = df_70['H2'].tolist()
list_l2 = df_70['L2'].tolist()

# Get the absolute newest values for the top KPI block display
latest_nbgb_ratio = list_nbgb[-1]
latest_h1 = list_h1[-1]
latest_l1 = list_l1[-1]

latest_nbsb_ratio = list_nbsb[-1]
latest_h2 = list_h2[-1]
latest_l2 = list_l2[-1]

# 4. Generate the pure HTML template
html_template = """<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<meta name='viewport' content='width=device-width, initial-scale=1.0'>
<title>Rotation Strategy Dashboard</title>
<script src='https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js'></script>
<style>
:root{
 --bg:#edf5ff;--panel:rgba(255,255,255,.86);--text:#0f1f3a;--muted:#5f708a;
 --tw:#1DA1F2;--tech:#0A2885;--tech2:#01147C;--good:#16a34a;--bad:#ef4444;
 --shadow:0 14px 40px rgba(8,20,60,.12);
}
*{box-sizing:border-box;}
body{margin:0;font-family:Segoe UI,Inter,Arial,sans-serif;
 background:radial-gradient(circle at 16% 18%,rgba(29,161,242,.18),transparent 26%),
 radial-gradient(circle at 82% 20%,rgba(10,40,133,.18),transparent 22%),
 linear-gradient(135deg,#f7fbff 0%,#e8f1fb 100%);color:var(--text);}
.container{max-width:1500px;margin:0 auto;padding:22px;}
.hero{position:relative;overflow:hidden;border-radius:28px;padding:28px;color:#fff;box-shadow:var(--shadow);
 background:linear-gradient(135deg,var(--tech2) 0%,var(--tech) 42%,var(--tw) 100%);
 display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:20px;}
.hero-content{flex:1;min-width:300px;}
.hero:before,.hero:after{content:'';position:absolute;border-radius:50%;background:rgba(255,255,255,.12);z-index:1;}
.hero:before{width:240px;height:240px;right:-70px;bottom:-70px;}
.hero:after{width:140px;height:140px;right:140px;top:34px;background:rgba(255,255,255,.08);}
.hero h1{font-size:2.1rem;margin:0 0 8px;position:relative;z-index:2;}
.hero p{margin:0;color:rgba(255,255,255,.93);position:relative;z-index:2;font-weight:500;}

.btn-runner{position:relative;z-index:2;background:#fff;color:var(--tech2);border:none;padding:12px 24px;border-radius:14px;font-weight:800;font-size:0.95rem;cursor:pointer;box-shadow:0 6px 20px rgba(0,0,0,0.15);transition:all 0.2s ease;display:flex;align-items:center;gap:8px;}
.btn-runner:hover{background:#edf5ff;transform:translateY(-2px);box-shadow:0 8px 25px rgba(0,0,0,0.2);}
.btn-runner:active{transform:translateY(0);}

.grid{display:grid;gap:16px;}
.kpis{grid-template-columns:repeat(6,1fr);margin-top:16px;}

.prices-bar{display:flex;gap:16px;justify-content:center;margin:16px 0;flex-wrap:wrap;}
.price-tag{background:var(--panel);border:1px solid rgba(255,255,255,.6);border-radius:12px;padding:12px 20px;font-size:.88rem;box-shadow:var(--shadow);font-weight:600;color:var(--muted);}
.price-tag strong{color:var(--text);font-size:1.05rem;margin-left:4px;}
.price-tag span{font-size:0.75rem;background:rgba(15,31,58,0.06);padding:3px 6px;border-radius:6px;margin-left:6px;font-weight:700;color:var(--tech2);}

@media(max-width:1100px){.kpis{grid-template-columns:repeat(2,1fr);}.charts{grid-template-columns:1fr;}.live-grid{grid-template-columns:1fr !important;}}
.card,.chart-card,.table-wrap{background:var(--panel);border:1px solid rgba(255,255,255,.6);border-radius:20px;box-shadow:var(--shadow);backdrop-filter:blur(10px);}
.card{padding:18px;}
.label{color:var(--muted);font-size:.78rem;text-transform:uppercase;letter-spacing:.08em;font-weight:800;margin-bottom:10px;}
.value{font-size:1.7rem;font-weight:900;line-height:1;}
.value.tw{color:var(--tw);}.value.tech{color:var(--tech);}.value.bad{color:var(--bad);}
.small{color:var(--muted);font-size:.85rem;margin-top:8px;}
.section-title{margin:24px 0 12px;font-size:1.06rem;font-weight:900;}
.charts{grid-template-columns:1fr 1fr;}
.chart-card{padding:16px;}
.chart-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;}
.badge-soft{display:inline-block;padding:6px 10px;border-radius:999px;background:linear-gradient(90deg,rgba(29,161,242,.12),rgba(10,40,133,.12));color:var(--tech2);font-size:.78rem;font-weight:900;border:1px solid #d8e7f7;}
canvas{max-height:240px;}
.table-wrap{overflow:hidden;margin-top:16px;}
.table-title{padding:16px 18px;font-weight:900;background:linear-gradient(90deg,rgba(29,161,242,.11),rgba(10,40,133,.11));border-bottom:1px solid #e6eef8;}
table{width:100%;border-collapse:collapse;}
th,td{padding:12px 10px;text-align:center;border-bottom:1px solid #eef3f8;font-size:.86rem;}
th{position:sticky;top:0;background:#f8fbff;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;font-size:.72rem;}
tr:hover td{background:#fafcff;}
.badge{display:inline-block;padding:6px 10px;border-radius:999px;color:#fff;font-size:.78rem;font-weight:800;}
.holding-bar{display:flex;gap:16px;justify-content:center;margin:16px 0;flex-wrap:wrap;}
.holding-tag{background:var(--panel);border:1px solid rgba(255,255,255,.6);border-radius:12px;padding:12px 24px;font-size:.9rem;box-shadow:var(--shadow);}
.holding-tag span{font-weight:800;font-size:1.05rem;}

.live-calc-card{background:var(--panel);border:1px solid rgba(255,255,255,.6);border-radius:24px;padding:24px;margin-top:16px;box-shadow:var(--shadow);}
.live-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;margin-bottom:20px;}
.input-box{display:flex;flex-direction:column;gap:8px;}
.input-box label{font-size:0.8rem;font-weight:800;color:var(--tech);text-transform:uppercase;letter-spacing:0.05em;}
.input-box input{padding:14px 18px;border-radius:14px;border:1px solid #d8e7f7;font-size:1.15rem;font-weight:700;color:var(--text);background:#fff;outline:none;transition:all 0.2s;}
.input-box input:focus{border-color:var(--tw);box-shadow:0 0 0 4px rgba(29,161,242,0.15);}
.live-results{display:flex;gap:40px;padding-top:18px;border-top:1px dashed #d8e7f7;flex-wrap:wrap;}
.result-tag{font-size:1.05rem;font-weight:600;color:var(--muted);}
.result-tag span{font-weight:900;color:var(--tech2);font-size:1.2rem;margin-left:6px;}

.footer-note{margin-top:20px;color:var(--muted);font-size:.84rem;text-align:center;}
</style>
</head>
<body>
<div class='container'>
  <div class='hero'>
    <div class='hero-content'>
      <h1>&#x1F504; Rotation Strategy Dashboard</h1>
      <p>NiftyBees / GoldBees / SilverBees &nbsp;&#x7C;&nbsp; As of <span class='latest-date-header'>{{LATEST_DATE}}</span></p>
    </div>
    <button class='btn-runner' onclick='runGitHubWorkflow()'>🔄 Trigger Screening Workflow</button>
  </div>

  <div class='holding-bar'>
    <div class='holding-tag'>Currently Holding (NB&#x2F;GB): <span style='color:#0A2885'>GoldBees</span></div>
    <div class='holding-tag'>Currently Holding (NB&#x2F;SB): <span style='color:#2F80ED'>SilverBees</span></div>
  </div>

  <div class='prices-bar'>
    <div class='price-tag'>NiftyBees Close: <strong id='close-nb'>{{CLOSE_NB}}</strong><span class='latest-date'>{{LATEST_DATE}}</span></div>
    <div class='price-tag'>GoldBees Close: <strong id='close-gb'>{{CLOSE_GB}}</strong><span class='latest-date'>{{LATEST_DATE}}</span></div>
    <div class='price-tag'>SilverBees Close: <strong id='close-sb'>{{CLOSE_SB}}</strong><span class='latest-date'>{{LATEST_DATE}}</span></div>
  </div>

  <div class='grid kpis'>
    <div class='card'><div class='label'>NB / GB</div><div class='value tw' id='kpi-nbgb'>{{LATEST_NBGB}}</div><div class='small'>Current ratio</div></div>
    <div class='card'><div class='label'>20D High NB/GB</div><div class='value tech'>{{LATEST_H1}}</div><div class='small'>Breakout level</div></div>
    <div class='card'><div class='label'>20D Low NB/GB</div><div class='value bad'>{{LATEST_L1}}</div><div class='small'>Breakdown level</div></div>
    <div class='card'><div class='label'>NB / SB</div><div class='value tw' id='kpi-nbsb'>{{LATEST_NBSB}}</div><div class='small'>Current ratio</div></div>
    <div class='card'><div class='label'>20D High NB/SB</div><div class='value tech'>{{LATEST_H2}}</div><div class='small'>Breakout level</div></div>
    <div class='card'><div class='label'>20D Low NB/SB</div><div class='value bad'>{{LATEST_L2}}</div><div class='small'>Breakdown level</div></div>
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
    <tbody><tr><td>12-05-26</td><td><span class='badge' style='background:#0A2885'>BUY GoldBees</span></td><td>131.1</td><td>&#8212;</td><td style='color:#16a34a;font-weight:700'>&#8212;</td><td>2.2465</td><td>2.1778</td></tr><tr><td>23-03-26</td><td><span class='badge' style='background:#1DA1F2'>BUY NiftyBees</span></td><td>259.0</td><td>265.04</td><td style='color:#16a34a;font-weight:700'>2.33%</td><td>2.2678</td><td>2.0157</td></tr><tr><td>01-12-25</td><td><span class='badge' style='background:#0A2885'>BUY GoldBees</span></td><td>106.65</td><td>111.62</td><td style='color:#16a34a;font-weight:700'>4.66%</td><td>2.912</td><td>2.7749</td></tr><tr><td>28-10-25</td><td><span class='badge' style='background:#1DA1F2'>BUY NiftyBees</span></td><td>291.01</td><td>296.18</td><td style='color:#16a34a;font-weight:700'>1.78%</td><td>2.9783</td><td>2.684</td></tr><tr><td>29-08-25</td><td><span class='badge' style='background:#0A2885'>BUY GoldBees</span></td><td>86.76</td><td>98.0</td><td style='color:#16a34a;font-weight:700'>12.96%</td><td>3.442</td><td>3.2615</td></tr></tbody></table>
  </div>

  <div class='table-wrap'>
    <div class='table-title'>NB / SB Events</div>
    <table><thead><tr><th>Date</th><th>Signal</th><th>Entry</th><th>Exit</th><th>Returns</th><th>20D High</th><th>20D Low</th></tr></thead>
    <tbody><tr><td>07-05-26</td><td><span class='badge' style='background:#2F80ED'>BUY SilverBees</span></td><td>240.03</td><td>&#8212;</td><td style='color:#16a34a;font-weight:700'>&#8212;</td><td>1.2182</td><td>1.1512</td></tr><tr><td>23-03-26</td><td><span class='badge' style='background:#1DA1F2'>BUY NiftyBees</span></td><td>259.0</td><td>273.62</td><td style='color:#16a34a;font-weight:700'>5.64%</td><td>1.2223</td><td>1.0265</td></tr><tr><td>02-03-26</td><td><span class='badge' style='background:#2F80ED'>BUY SilverBees</span></td><td>251.05</td><td>199.0</td><td style='color:#ef4444;font-weight:700'>-20.73%</td><td>1.3249</td><td>1.0982</td></tr><tr><td>17-02-26</td><td><span class='badge' style='background:#1DA1F2'>BUY NiftyBees</span></td><td>291.0</td><td>280.99</td><td style='color:#ef4444;font-weight:700'>-3.44%</td><td>1.2949</td><td>0.8166</td></tr><tr><td>01-12-25</td><td><span class='badge' style='background:#2F80ED'>BUY SilverBees</span></td><td>167.01</td><td>220.89</td><td style='color:#16a34a;font-weight:700'>32.26%</td><td>2.0692</td><td>1.8613</td></tr></tbody></table>
  </div>

  <div class='section-title'>📅 Daily Tracking Data (Last 5 Columns)</div>
  <div class='table-wrap'>
    <div class='table-title'>Recent Historical Matrix Profile</div>
    <table>
      <thead>
        <tr>
          <th>Date</th>
          <th>NB / GB Ratio</th>
          <th>20D High (NB/GB)</th>
          <th>20D Low (NB/GB)</th>
          <th>NB / SB Ratio</th>
          <th>20D High (NB/SB)</th>
          <th>20D Low (NB/SB)</th>
        </tr>
      </thead>
      <tbody id="daily-data-rows">
      </tbody>
    </table>
  </div>

  <div class='section-title'>📊 Live Valuation Matrix</div>
  <div class='live-calc-card'>
    <div class='live-grid'>
      <div class='input-box'>
        <label>NiftyBees Live Price</label>
        <input type='number' id='nb-input' value='{{VAL_NB}}' step='0.05' oninput='runLiveCalculation()'>
      </div>
      <div class='input-box'>
        <label>GoldBees Live Price</label>
        <input type='number' id='gb-input' value='{{VAL_GB}}' step='0.05' oninput='runLiveCalculation()'>
      </div>
      <div class='input-box'>
        <label>SilverBees Live Price</label>
        <input type='number' id='sb-input' value='{{VAL_SB}}' step='0.05' oninput='runLiveCalculation()'>
      </div>
    </div>
    <div class='live-results'>
      <div class='result-tag'>Live NB / GB Ratio: <span id='live-nbgb-label'>{{LATEST_NBGB}}</span></div>
      <div class='result-tag'>Live NB / SB Ratio: <span id='live-nbsb-label'>{{LATEST_NBSB}}</span></div>
    </div>
  </div>

  <div class='footer-note'>Auto-refreshed daily at 4 PM IST via GitHub Actions &#x2192; Cloudflare R2 Bucket Storage</div>
</div>

<script>
const labels = {{CHART_LABELS}};
const nbgb = {{LIST_NBGB}};
const h1 = {{LIST_H1}};
const l1 = {{LIST_L1}};
const nbsb = {{LIST_NBSB}};
const h2 = {{LIST_H2}};
const l2 = {{LIST_L2}};

const tbody = document.getElementById('daily-data-rows');
for (let i = labels.length - 5; i < labels.length; i++) {
  if(labels[i]) {
    const row = `<tr>
      <td style="font-weight:700;">${labels[i]}</td>
      <td style="color:var(--tw); font-weight:700;">${nbgb[i]}</td>
      <td>${h1[i] || '—'}</td>
      <td>${l1[i] || '—'}</td>
      <td style="color:var(--tech); font-weight:700;">${nbsb[i]}</td>
      <td>${h2[i] || '—'}</td>
      <td>${l2[i] || '—'}</td>
    </tr>`;
    tbody.insertAdjacentHTML('beforeend', row);
  }
}

function runLiveCalculation() {
  const nb = parseFloat(document.getElementById('nb-input').value) || 0;
  const gb = parseFloat(document.getElementById('gb-input').value) || 0;
  const sb = parseFloat(document.getElementById('sb-input').value) || 0;

  if (gb > 0) {
    const calculatedNBGB = (nb / gb).toFixed(4);
    document.getElementById('live-nbgb-label').innerText = calculatedNBGB;
    document.getElementById('kpi-nbgb').innerText = calculatedNBGB;
  }
  if (sb > 0) {
    const calculatedNBSB = (nb / sb).toFixed(4);
    document.getElementById('live-nbsb-label').innerText = calculatedNBSB;
    document.getElementById('kpi-nbsb').innerText = calculatedNBSB;
  }
}

function makeChart(id,main,hi,lo,color){
  new Chart(document.getElementById(id),{type:'line',
    data:{labels,datasets:[
      {label:'Current Ratio',data:main,borderColor:color,borderWidth:3,pointRadius:0,tension:.28,fill:false},
      {label:'20D High',data:hi,borderColor:'#01147C',borderWidth:2,pointRadius:0,borderDash:[6,4]},
      {label:'20D Low',data:lo,borderColor:'#ef4444',borderWidth:2,pointRadius:0,borderDash:[6,4]}
    ]},
    options:{responsive:true,plugins:{legend:{labels:{color:'#5f708a'}}},
      scales:{x:{ticks:{color:'#6b7b92'},grid:{color:'#e6eef8'}},y:{ticks:{color:'#6b7b92'},grid:{color:'#e6eef8'}}}}
  );
}
makeChart('c1',nbgb,h1,l1,'#1DA1F2');
makeChart('c2',nbsb,h2,l2,'#0A2885');

function runGitHubWorkflow() {
  const USER = "YOUR_GITHUB_USERNAME";         
  const REPOSITORY = "YOUR_REPOSITORY_NAME";   
  const WORKFLOW_FILE = "YOUR_WORKFLOW.yml";   
  const AUTH_TOKEN = "YOUR_PERSONAL_ACCESS_TOKEN"; 

  if (USER === "YOUR_GITHUB_USERNAME") {
    alert("Configuration Needed:\\nOpen this HTML file in a text editor and update the 'runGitHubWorkflow' variables inside the script tag with your actual GitHub repository credentials.");
    return;
  }

  if (!confirm("Confirm execution: Trigger screening engine execution pipeline on GitHub?")) return;

  fetch(`https://api.github.com/repos/${USER}/${REPOSITORY}/actions/workflows/${WORKFLOW_FILE}/dispatches`, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${AUTH_TOKEN}`,
      "Accept": "application/vnd.github+v3+json",
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ ref: "main" })
  })
  .then(res => {
    if (res.status === 204) {
      alert("✅ Screening automation pipeline triggered successfully on GitHub Actions.");
    } else {
      res.json().then(err => alert(`❌ API Rejection: ${err.message || res.statusText}`));
    }
  })
  .catch(err => {
    console.error("Workflow activation fault:", err);
    alert("❌ Connection Refused: Check your internet connection or verify token CORS settings.");
  });
}
</script>
</body>
</html>"""

# 5. Inject calculated python variables safely into the placeholders
replacements = {
    "{{LATEST_DATE}}": str(formatted_date),
    "{{CLOSE_NB}}": f"{current_nb:.2f}",
    "{{CLOSE_GB}}": f"{current_gb:.2f}",
    "{{CLOSE_SB}}": f"{current_sb:.2f}",
    "{{LATEST_NBGB}}": f"{latest_nbgb_ratio:.4f}",
    "{{LATEST_H1}}": f"{latest_h1:.4f}",
    "{{LATEST_L1}}": f"{latest_l1:.4f}",
    "{{LATEST_NBSB}}": f"{latest_nbsb_ratio:.4f}",
    "{{LATEST_H2}}": f"{latest_h2:.4f}",
    "{{LATEST_L2}}": f"{latest_l2:.4f}",
    "{{VAL_NB}}": str(current_nb),
    "{{VAL_GB}}": str(current_gb),
    "{{VAL_SB}}": str(current_sb),
    "{{CHART_LABELS}}": str(chart_labels),
    "{{LIST_NBGB}}": str(list_nbgb),
    "{{LIST_H1}}": str(list_h1),
    "{{LIST_L1}}": str(list_l1),
    "{{LIST_NBSB}}": str(list_nbsb),
    "{{LIST_H2}}": str(list_h2),
    "{{LIST_L2}}": str(list_l2)
}

final_html = html_template
for placeholder, value in replacements.items():
    final_html = final_html.replace(placeholder, value)

# 6. Save the actual text out as structural HTML
with open(html_path, "w", encoding="utf-8") as f:
    f.write(final_html)

print(f"Success: Up-to-date Dashboard HTML generated completely inside {html_path}")
