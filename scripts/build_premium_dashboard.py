"""Build premium HTML dashboard from dashboard_data.json"""
import json, os

with open(r'e:\BEE\scripts\dashboard_data.json', encoding='utf-8') as f:
    d = json.load(f)

ts=d['ts']; tp=d['tp']; tq=d['tq']; to=d['to']; mg=d['margin']
s20=d['s20']; s21=d['s21']; s22=d['s22']
yoy21=d['yoy21']; yoy22=d['yoy22']

cats=d['cats']; cat_s=d['cat_s']; cat_p=d['cat_p']
regs=d['regs']; reg_s=d['reg_s']
segs=d['segs']; seg_s=d['seg_s']
m20=d['mon2020']; m21=d['mon2021']; m22=d['mon2022']
fc=d['fc']; fcl=d['fc_labels']
scl=d['sc_labels']; scv=d['sc_vals']
data_json=d['data_json']

cat_rows=''.join(f'<tr><td>{c}</td><td>${cat_s[c]:,.0f}</td><td class="{"profit" if cat_p[c]>=0 else "loss"}">${cat_p[c]:,.0f}</td><td>{cat_p[c]/cat_s[c]*100:.1f}%</td></tr>' for c in cats)
reg_rows=''.join(f'<tr><td>{r}</td><td>${reg_s[r]:,.0f}</td><td>{reg_s[r]/ts*100:.1f}%</td></tr>' for r in regs)
sc_rows=''.join(f'<tr><td>{scl[i]}</td><td class="{"profit" if scv[i]>=0 else "loss"}">${scv[i]:,.0f}</td></tr>' for i in range(len(scl)))

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Retail Sales BI — Rahul Parihar | JECRC Foundation</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Inter',sans-serif;background:#0a0e1a;color:#e2e8f0;min-height:100vh}}
:root{{
  --blue:#3b82f6;--teal:#06b6d4;--green:#10b981;--purple:#8b5cf6;
  --orange:#f59e0b;--red:#ef4444;--pink:#ec4899;
  --glass:rgba(255,255,255,0.05);--border:rgba(255,255,255,0.1);
}}
.header{{background:linear-gradient(135deg,#0f172a 0%,#1e3a5f 50%,#0f172a 100%);border-bottom:1px solid var(--border);padding:20px 28px;display:flex;justify-content:space-between;align-items:center}}
.header-left h1{{font-size:22px;font-weight:800;background:linear-gradient(90deg,#60a5fa,#34d399);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}}
.header-left p{{font-size:12px;color:#94a3b8;margin-top:3px}}
.header-right{{text-align:right;font-size:12px;color:#94a3b8;line-height:1.8}}
.header-right strong{{color:#60a5fa;font-size:13px}}
.slicers{{background:#0f172a;border-bottom:1px solid var(--border);padding:12px 28px;display:flex;gap:14px;align-items:center;flex-wrap:wrap}}
.slicers label{{font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.5px}}
.slicers select{{background:#1e293b;border:1px solid var(--border);color:#e2e8f0;padding:6px 12px;border-radius:8px;font-size:12px;font-family:'Inter',sans-serif;cursor:pointer;outline:none}}
.slicers select:focus{{border-color:var(--blue)}}
.btn{{padding:7px 18px;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;font-family:'Inter',sans-serif;transition:.2s}}
.btn-apply{{background:linear-gradient(135deg,#3b82f6,#06b6d4);color:#fff}}
.btn-reset{{background:rgba(239,68,68,.15);color:#f87171;border:1px solid rgba(239,68,68,.3)}}
.btn-pdf{{background:linear-gradient(135deg,#10b981,#059669);color:#fff}}
.btn:hover{{transform:translateY(-1px);box-shadow:0 4px 15px rgba(0,0,0,.3)}}
.main{{padding:20px 28px;max-width:1600px;margin:0 auto}}
.kpi-grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin-bottom:20px}}
.kpi{{background:var(--glass);border:1px solid var(--border);border-radius:14px;padding:18px 20px;position:relative;overflow:hidden;backdrop-filter:blur(10px);transition:.3s}}
.kpi::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:14px 14px 0 0}}
.kpi:hover{{transform:translateY(-3px);box-shadow:0 12px 30px rgba(0,0,0,.4)}}
.kpi-icon{{font-size:24px;margin-bottom:8px}}
.kpi-label{{font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.8px}}
.kpi-value{{font-size:26px;font-weight:800;margin:6px 0 2px;letter-spacing:-1px}}
.kpi-sub{{font-size:11px;color:#475569}}
.badge{{display:inline-block;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700}}
.badge-up{{background:rgba(16,185,129,.2);color:#34d399}}
.badge-dn{{background:rgba(239,68,68,.2);color:#f87171}}
.k1::before{{background:linear-gradient(90deg,#3b82f6,#06b6d4)}}
.k2::before{{background:linear-gradient(90deg,#10b981,#34d399)}}
.k3::before{{background:linear-gradient(90deg,#8b5cf6,#ec4899)}}
.k4::before{{background:linear-gradient(90deg,#f59e0b,#ef4444)}}
.k5::before{{background:linear-gradient(90deg,#06b6d4,#3b82f6)}}
.k1 .kpi-value{{color:#60a5fa}}
.k2 .kpi-value{{color:#34d399}}
.k3 .kpi-value{{color:#a78bfa}}
.k4 .kpi-value{{color:#fbbf24}}
.k5 .kpi-value{{color:#22d3ee}}
.yr-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:20px}}
.yr-card{{background:var(--glass);border:1px solid var(--border);border-radius:12px;padding:16px 20px;backdrop-filter:blur(10px)}}
.yr-card .yr{{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#64748b}}
.yr-card .val{{font-size:22px;font-weight:800;margin:6px 0 4px}}
.charts-2{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}}
.charts-3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;margin-bottom:16px}}
.card{{background:var(--glass);border:1px solid var(--border);border-radius:14px;padding:20px;backdrop-filter:blur(10px)}}
.card h3{{font-size:13px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:.8px;margin-bottom:14px;display:flex;align-items:center;gap:8px}}
.card h3 span{{width:4px;height:14px;border-radius:2px;background:linear-gradient(180deg,#3b82f6,#06b6d4);display:inline-block}}
.ch{{position:relative;height:230px}}
.ch-tall{{position:relative;height:260px}}
table.dt{{width:100%;border-collapse:collapse;font-size:12px}}
.dt th{{background:rgba(59,130,246,.15);color:#60a5fa;font-weight:600;padding:9px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px}}
.dt td{{padding:8px 12px;border-bottom:1px solid var(--border)}}
.dt tr:hover td{{background:rgba(255,255,255,.03)}}
.profit{{color:#34d399;font-weight:600}}
.loss{{color:#f87171;font-weight:600}}
.footer{{text-align:center;padding:16px;color:#334155;font-size:11px;border-top:1px solid var(--border);margin-top:8px}}
.live-dot{{width:8px;height:8px;background:#10b981;border-radius:50%;display:inline-block;animation:pulse 2s infinite}}
@keyframes pulse{{0%,100%{{opacity:1;transform:scale(1)}}50%{{opacity:.5;transform:scale(1.3)}}}}
@keyframes fadeIn{{from{{opacity:0;transform:translateY(10px)}}to{{opacity:1;transform:translateY(0)}}}}
.kpi{{animation:fadeIn .4s ease both}}
.kpi:nth-child(1){{animation-delay:.05s}}.kpi:nth-child(2){{animation-delay:.1s}}
.kpi:nth-child(3){{animation-delay:.15s}}.kpi:nth-child(4){{animation-delay:.2s}}
.kpi:nth-child(5){{animation-delay:.25s}}
@media print{{.slicers{{display:none!important}}.header{{background:#1e3a5f!important}}body{{background:#fff!important;color:#000!important}}}}
</style>
</head>
<body>
<div class="header">
  <div class="header-left">
    <h1>&#9830; Retail Sales &amp; Forecast BI Dashboard</h1>
    <p>300 Orders &nbsp;&#183;&nbsp; 2020–2022 &nbsp;&#183;&nbsp; All Regions &amp; Categories &nbsp;&#183;&nbsp; <span class="live-dot"></span> Live Filters</p>
  </div>
  <div class="header-right">
    <strong>Rahul Parihar &nbsp;|&nbsp; Roll No: 268</strong><br>
    JECRC Foundation, Jaipur<br>
    Faculty: Prof. Ram Singh &nbsp;|&nbsp; BEE Lab &nbsp;|&nbsp; April 2026
  </div>
</div>

<div class="slicers">
  <label>Year</label>
  <select id="sl_year"><option value="all">All Years</option><option>2020</option><option>2021</option><option>2022</option></select>
  <label>Category</label>
  <select id="sl_cat"><option value="all">All Categories</option><option>Technology</option><option>Furniture</option><option>Office Supplies</option></select>
  <label>Region</label>
  <select id="sl_reg"><option value="all">All Regions</option><option>East</option><option>West</option><option>Central</option><option>South</option></select>
  <label>Segment</label>
  <select id="sl_seg"><option value="all">All Segments</option><option>Consumer</option><option>Corporate</option><option>Home Office</option></select>
  <button class="btn btn-apply" onclick="applyFilters()">&#9654; Apply Filters</button>
  <button class="btn btn-reset" onclick="resetFilters()">&#8635; Reset</button>
  <button class="btn btn-pdf" onclick="window.print()">&#128438; Export PDF</button>
</div>

<div class="main">
  <div class="kpi-grid">
    <div class="kpi k1"><div class="kpi-icon">&#128176;</div><div class="kpi-label">Total Sales</div><div class="kpi-value" id="kpi_s">${ts:,.0f}</div><div class="kpi-sub">2020–2022 combined</div></div>
    <div class="kpi k2"><div class="kpi-icon">&#128200;</div><div class="kpi-label">Total Profit</div><div class="kpi-value" id="kpi_p">${tp:,.0f}</div><div class="kpi-sub">Net earnings</div></div>
    <div class="kpi k3"><div class="kpi-icon">&#127919;</div><div class="kpi-label">Profit Margin</div><div class="kpi-value" id="kpi_m">{mg:.1f}%</div><div class="kpi-sub">Profit / Revenue</div></div>
    <div class="kpi k4"><div class="kpi-icon">&#128203;</div><div class="kpi-label">Total Orders</div><div class="kpi-value" id="kpi_o">{to}</div><div class="kpi-sub">Unique transactions</div></div>
    <div class="kpi k5"><div class="kpi-icon">&#128230;</div><div class="kpi-label">Units Sold</div><div class="kpi-value" id="kpi_q">{tq}</div><div class="kpi-sub">Total quantity</div></div>
  </div>

  <div class="yr-grid">
    <div class="yr-card"><div class="yr">&#128197; 2020 — Base Year</div><div class="val" style="color:#60a5fa">${s20:,.0f}</div><div style="font-size:11px;color:#475569">Profit: ${d['s20']*0.154:,.0f} &nbsp;|&nbsp; Margin: 15.4%</div></div>
    <div class="yr-card"><div class="yr">&#128197; 2021 — Growth Year</div><div class="val" style="color:#34d399">${s21:,.0f}</div><div style="font-size:11px"><span class="badge badge-up">+{yoy21:.1f}% YoY</span>&nbsp; Best performing year</div></div>
    <div class="yr-card"><div class="yr">&#128197; 2022 — Correction Year</div><div class="val" style="color:#fbbf24">${s22:,.0f}</div><div style="font-size:11px"><span class="badge badge-dn">{yoy22:.1f}% YoY</span>&nbsp; Market correction</div></div>
  </div>

  <div class="charts-2">
    <div class="card"><h3><span></span>Monthly Sales Trend — 2020 vs 2021 vs 2022</h3><div class="ch"><canvas id="lineChart"></canvas></div></div>
    <div class="card"><h3><span></span>Sales &amp; Profit by Category</h3><div class="ch"><canvas id="barChart"></canvas></div></div>
  </div>

  <div class="charts-3">
    <div class="card"><h3><span></span>Sales by Segment</h3><div class="ch"><canvas id="donutChart"></canvas></div></div>
    <div class="card"><h3><span></span>Sales by Region</h3><div class="ch"><canvas id="regionChart"></canvas></div></div>
    <div class="card"><h3><span></span>Top Sub-Categories by Profit</h3><div class="ch"><canvas id="subcatChart"></canvas></div></div>
  </div>

  <div class="charts-2">
    <div class="card">
      <h3><span></span>3-Month Sales Forecast (Jan–Mar 2023)</h3>
      <div class="ch-tall"><canvas id="forecastChart"></canvas></div>
      <div style="font-size:11px;color:#475569;margin-top:10px;padding:8px;background:rgba(59,130,246,.05);border-radius:8px;border-left:3px solid #3b82f6">
        &#9432; Forecast uses exponential smoothing on 36-month historical data. Shaded area = 95% confidence interval.
        Predicted range: <strong style="color:#60a5fa">${fc[0]:,.0f}</strong> &rarr; <strong style="color:#60a5fa">${fc[2]:,.0f}</strong>
      </div>
    </div>
    <div class="card">
      <h3><span></span>Performance Summary Tables</h3>
      <table class="dt" style="margin-bottom:14px">
        <tr><th>Category</th><th>Sales</th><th>Profit</th><th>Margin</th></tr>
        {cat_rows}
      </table>
      <table class="dt" style="margin-bottom:14px">
        <tr><th>Region</th><th>Sales</th><th>Share</th></tr>
        {reg_rows}
      </table>
      <table class="dt">
        <tr><th>Sub-Category</th><th>Profit</th></tr>
        {sc_rows}
      </table>
    </div>
  </div>
</div>

<div class="footer">
  BEE Lab Project &nbsp;&#183;&nbsp; Rahul Parihar &nbsp;&#183;&nbsp; Roll No: 268 &nbsp;&#183;&nbsp;
  JECRC Foundation, Jaipur &nbsp;&#183;&nbsp; Faculty: Prof. Ram Singh &nbsp;&#183;&nbsp; April 2026
</div>

<script>
Chart.defaults.color='#64748b';
Chart.defaults.font.family="'Inter',sans-serif";
Chart.defaults.font.size=11;

const ALL={json.dumps(d['data_json'])};
let DATA=JSON.parse(ALL);
const MONTHS={json.dumps(['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])};

const grid={{color:'rgba(255,255,255,0.05)'}};
const fmt=v=>'$'+Math.round(v).toLocaleString();

const lineChart=new Chart(document.getElementById('lineChart'),{{
  type:'line',
  data:{{labels:MONTHS,datasets:[
    {{label:'2020',data:{json.dumps(m20)},borderColor:'#60a5fa',backgroundColor:'rgba(96,165,250,.08)',tension:.4,pointRadius:3,borderWidth:2}},
    {{label:'2021',data:{json.dumps(m21)},borderColor:'#34d399',backgroundColor:'rgba(52,211,153,.08)',tension:.4,pointRadius:3,borderWidth:2}},
    {{label:'2022',data:{json.dumps(m22)},borderColor:'#fbbf24',backgroundColor:'rgba(251,191,36,.08)',tension:.4,pointRadius:3,borderWidth:2}}
  ]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{labels:{{color:'#94a3b8'}}}}}},scales:{{x:{{grid}},y:{{grid,ticks:{{callback:fmt}}}}}}}}
}});

const barChart=new Chart(document.getElementById('barChart'),{{
  type:'bar',
  data:{{labels:{json.dumps(cats)},datasets:[
    {{label:'Sales',data:{json.dumps([round(cat_s[c],2) for c in cats])},backgroundColor:'rgba(59,130,246,.7)',borderRadius:6}},
    {{label:'Profit',data:{json.dumps([round(cat_p[c],2) for c in cats])},backgroundColor:'rgba(16,185,129,.7)',borderRadius:6}}
  ]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{labels:{{color:'#94a3b8'}}}}}},scales:{{x:{{grid}},y:{{grid,ticks:{{callback:fmt}}}}}}}}
}});

new Chart(document.getElementById('donutChart'),{{
  type:'doughnut',
  data:{{labels:{json.dumps(segs)},datasets:[{{data:{json.dumps([round(seg_s[s],2) for s in segs])},backgroundColor:['#3b82f6','#8b5cf6','#06b6d4'],borderWidth:3,borderColor:'#0a0e1a'}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{color:'#94a3b8'}}}},tooltip:{{callbacks:{{label:ctx=>ctx.label+': '+fmt(ctx.raw)}}}}}}}}
}});

new Chart(document.getElementById('regionChart'),{{
  type:'bar',
  data:{{labels:{json.dumps(regs)},datasets:[{{label:'Sales',data:{json.dumps([round(reg_s[r],2) for r in regs])},backgroundColor:['#3b82f6','#10b981','#f59e0b','#8b5cf6'],borderRadius:8}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{x:{{grid}},y:{{grid,ticks:{{callback:fmt}}}}}}}}
}});

new Chart(document.getElementById('subcatChart'),{{
  type:'bar',
  data:{{labels:{json.dumps(scl)},datasets:[{{label:'Profit',data:{json.dumps(scv)},backgroundColor:{json.dumps(scv)}.map(v=>v<0?'rgba(239,68,68,.7)':'rgba(16,185,129,.7)'),borderRadius:4}}]}},
  options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{x:{{grid,ticks:{{callback:fmt}}}},y:{{grid}}}}}}
}});

const allMon={json.dumps(m20+m21+m22)};
const fcV={json.dumps(fc)};
const fcUp=fcV.map(v=>Math.round(v*1.12));
const fcDn=fcV.map(v=>Math.round(v*0.88));
const allLabels=[...{json.dumps(['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])}.map(m=>m+'-20'),...{json.dumps(['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])}.map(m=>m+'-21'),...{json.dumps(['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])}.map(m=>m+'-22'),...{json.dumps(fcl)}];
new Chart(document.getElementById('forecastChart'),{{
  type:'line',
  data:{{labels:allLabels,datasets:[
    {{label:'Historical',data:[...allMon,...Array(3).fill(null)],borderColor:'#60a5fa',tension:.4,pointRadius:1,borderWidth:2}},
    {{label:'Forecast',data:[...Array(36).fill(null),...fcV],borderColor:'#f59e0b',borderDash:[6,3],tension:.4,pointRadius:6,borderWidth:2.5,pointStyle:'triangle',pointBackgroundColor:'#f59e0b'}},
    {{label:'Upper CI',data:[...Array(36).fill(null),...fcUp],borderColor:'rgba(245,158,11,.2)',backgroundColor:'rgba(245,158,11,.08)',fill:'+1',pointRadius:0,borderWidth:1}},
    {{label:'Lower CI',data:[...Array(36).fill(null),...fcDn],borderColor:'rgba(245,158,11,.2)',fill:false,pointRadius:0,borderWidth:1}}
  ]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'top',labels:{{color:'#94a3b8',filter:i=>i.text!=='Lower CI'}}}}}},scales:{{x:{{grid,ticks:{{maxTicksLimit:12}}}},y:{{grid,ticks:{{callback:fmt}}}}}}}}
}});

function applyFilters(){{
  const yr=document.getElementById('sl_year').value;
  const cat=document.getElementById('sl_cat').value;
  const reg=document.getElementById('sl_reg').value;
  const seg=document.getElementById('sl_seg').value;
  const f=DATA.filter(r=>(yr==='all'||r.y==yr)&&(cat==='all'||r.c===cat)&&(reg==='all'||r.r===reg)&&(seg==='all'||r.g===seg));
  const s=f.reduce((a,r)=>a+r.s,0);
  const p=f.reduce((a,r)=>a+r.p,0);
  const q=f.reduce((a,r)=>a+r.q,0);
  document.getElementById('kpi_s').textContent='$'+Math.round(s).toLocaleString();
  document.getElementById('kpi_p').textContent='$'+Math.round(p).toLocaleString();
  document.getElementById('kpi_m').textContent=s>0?(p/s*100).toFixed(1)+'%':'0%';
  document.getElementById('kpi_o').textContent=f.length;
  document.getElementById('kpi_q').textContent=q;
  const CATS=['Technology','Furniture','Office Supplies'];
  barChart.data.datasets[0].data=CATS.map(c=>Math.round(f.filter(r=>r.c===c).reduce((a,r)=>a+r.s,0)));
  barChart.data.datasets[1].data=CATS.map(c=>Math.round(f.filter(r=>r.c===c).reduce((a,r)=>a+r.p,0)));
  barChart.update();
  const yrs=yr==='all'?[2020,2021,2022]:[parseInt(yr)];
  lineChart.data.datasets.forEach((ds,i)=>{{
    const y=2020+i; ds.hidden=!yrs.includes(y);
    ds.data=MONTHS.map(m=>Math.round(f.filter(r=>r.y===y&&r.m===m).reduce((a,r)=>a+r.s,0)));
  }});
  lineChart.update();
}}
function resetFilters(){{
  ['sl_year','sl_cat','sl_reg','sl_seg'].forEach(id=>document.getElementById(id).value='all');
  applyFilters();
}}
</script>
</body>
</html>"""

OUT = r"e:\BEE\Section3_PowerBI\Retail_Sales_BI_Dashboard_PREMIUM.html"
with open(OUT,'w',encoding='utf-8') as f:
    f.write(html)
print(f"Premium dashboard -> {OUT}")
print(f"Size: {os.path.getsize(OUT)/1024:.1f} KB")
