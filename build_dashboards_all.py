import pandas as pd
import numpy as np
import json, os, sys, datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(BASE_DIR, 'QA_Agent'))
from widget import qa_widget as _qa_widget
date_str = os.environ.get('HPI_OUTPUT_DATE', datetime.date.today().strftime('%Y%m%d'))

# ── Load clean data ──────────────────────────────────────────────────────────
df = pd.read_excel(os.path.join(BASE_DIR, 'HPIAdmins_v1.xlsx'))
df['Date'] = pd.to_datetime(df['Date'])
df_valid = df[df['DataGapFlag'] == False].copy()

FOLDER = os.path.join(BASE_DIR, date_str)
os.makedirs(FOLDER, exist_ok=True)

regions = sorted(df['RegionName'].unique().tolist())
COLORS = {
    'Greater Manchester': '#1B3A5C',   # Navy Base
    'Merseyside':         '#2D6A4F',   # Green Base
    'South Yorkshire':    '#C44E2C',   # Terracotta Base
    'Tyne and Wear':      '#D4A017',   # Gold Base
    'West Midlands':      '#4A6B8C',   # Navy Light
    'West Yorkshire':     '#5C9A7E',   # Green Light
}

# ── Pre-compute shared data ───────────────────────────────────────────────────
dates_all = sorted(df_valid['Date'].dt.strftime('%Y-%m').unique().tolist())

price_by_region, vol_by_region, yoy_by_region, mom_by_region = {}, {}, {}, {}
for r in regions:
    rdf = df_valid[df_valid['RegionName'] == r].sort_values('Date')
    price_by_region[r] = rdf['AveragePrice'].round(0).tolist()
    vol_by_region[r]   = rdf['SalesVolume'].tolist()
    yoy_by_region[r]   = rdf['12mPctChange'].round(2).tolist()
    mom_by_region[r]   = rdf['1mPctChange'].round(2).tolist()

# Latest full month
latest_date = df_valid['Date'].max()
latest = df_valid[df_valid['Date'] == latest_date].sort_values('AveragePrice', ascending=False)
first_date  = df_valid['Date'].min()
first       = df_valid[df_valid['Date'] == first_date]

# KPIs
latest_avg = round(latest['AveragePrice'].mean(), 0)
first_avg  = round(first['AveragePrice'].mean(), 0)
total_growth = round((latest_avg / first_avg - 1) * 100, 1)
total_vol  = int(latest['SalesVolume'].sum())

first_year   = first_date.year
latest_year  = latest_date.year
latest_label = latest_date.strftime('%B %Y')   # e.g. "December 2017"
latest_short = latest_date.strftime('%b %Y')   # e.g. "Dec 2017"

# Per-region growth (dynamic range)
region_growth = []
for r in regions:
    p0 = float(first[first['RegionName']==r]['AveragePrice'].values[0])
    p1 = float(latest[latest['RegionName']==r]['AveragePrice'].values[0])
    region_growth.append({'region': r, 'start': round(p0,0), 'end': round(p1,0),
                          'growth': round((p1/p0-1)*100,1), 'color': COLORS[r]})
region_growth_sorted = sorted(region_growth, key=lambda x: x['growth'], reverse=True)
fastest = region_growth_sorted[0]
slowest = region_growth_sorted[-1]

# Yearly average price (all regions combined)
yearly = df_valid.groupby(df_valid['Date'].dt.year)['AveragePrice'].mean().round(0)
yearly_labels = yearly.index.tolist()
yearly_values = yearly.values.tolist()

# Seasonal pattern: avg by month-of-year (all regions)
df_valid['Month'] = df_valid['Date'].dt.month
monthly_seasonal = df_valid.groupby('Month')['AveragePrice'].mean().round(0)
month_names = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

# Property types latest snapshot
prop_cols   = ['DetachedPrice','SemiDetachedPrice','TerracedPrice','FlatPrice']
prop_labels = ['Detached','Semi-Detached','Terraced','Flat']
prop_colors = ['#3B82F6','#10B981','#F59E0B','#EC4899']
prop_data_by_region = {}
for r in regions:
    row = latest[latest['RegionName']==r]
    prop_data_by_region[r] = [round(float(row[p].values[0]),0) for p in prop_cols] if len(row) > 0 else [0]*4

# Market avg price by property type over years
prop_yearly = {}
for pc, pl in zip(prop_cols, prop_labels):
    yr_avg = df_valid.groupby(df_valid['Date'].dt.year)[pc].mean().round(0)
    prop_yearly[pl] = yr_avg.values.tolist()

# Latest price & volume per region
latest_by_region = {
    r: {
        'price': round(float(latest[latest['RegionName']==r]['AveragePrice'].values[0]),0),
        'vol':   int(latest[latest['RegionName']==r]['SalesVolume'].values[0]),
        'yoy':   round(float(latest[latest['RegionName']==r]['12mPctChange'].values[0]),1)
                 if pd.notna(latest[latest['RegionName']==r]['12mPctChange'].values[0]) else 0,
        'det':   round(float(latest[latest['RegionName']==r]['DetachedPrice'].values[0]),0),
        'semi':  round(float(latest[latest['RegionName']==r]['SemiDetachedPrice'].values[0]),0),
        'terr':  round(float(latest[latest['RegionName']==r]['TerracedPrice'].values[0]),0),
        'flat':  round(float(latest[latest['RegionName']==r]['FlatPrice'].values[0]),0),
        'color': COLORS[r],
    }
    for r in regions
}

# YoY history for all regions combined (for ops table)
ops_rows = []
for _, row in df_valid.sort_values(['Date','RegionName']).iterrows():
    ops_rows.append({
        'date': row['Date'].strftime('%Y-%m'),
        'region': row['RegionName'],
        'area': row['Area'],
        'price': round(float(row['AveragePrice']),0),
        'vol': int(row['SalesVolume']),
        'yoy': round(float(row['12mPctChange']),2) if pd.notna(row['12mPctChange']) else None,
        'mom': round(float(row['1mPctChange']),2) if pd.notna(row['1mPctChange']) else None,
        'det': round(float(row['DetachedPrice']),0),
        'semi': round(float(row['SemiDetachedPrice']),0),
        'terr': round(float(row['TerracedPrice']),0),
        'flat': round(float(row['FlatPrice']),0),
    })

print("Data preparation complete.")

# ── Shared JS data block ──────────────────────────────────────────────────────
def js_block():
    return f"""
const DATES={json.dumps(dates_all)};
const REGIONS={json.dumps(regions)};
const COLORS={json.dumps(COLORS)};
const PRICE_BY_REGION={json.dumps(price_by_region)};
const VOL_BY_REGION={json.dumps(vol_by_region)};
const YOY_BY_REGION={json.dumps(yoy_by_region)};
const MOM_BY_REGION={json.dumps(mom_by_region)};
const YEARLY_LABELS={json.dumps(yearly_labels)};
const YEARLY_VALUES={json.dumps(yearly_values)};
const PROP_LABELS={json.dumps(prop_labels)};
const PROP_COLORS={json.dumps(prop_colors)};
const PROP_DATA={json.dumps(prop_data_by_region)};
const PROP_YEARLY={json.dumps(prop_yearly)};
const REGION_GROWTH={json.dumps(region_growth)};
const REGION_GROWTH_SORTED={json.dumps(region_growth_sorted)};
const LATEST_BY_REGION={json.dumps(latest_by_region)};
const MONTH_NAMES={json.dumps(month_names)};
const SEASONAL={json.dumps(monthly_seasonal.values.tolist())};
const OPS_ROWS={json.dumps(ops_rows)};
const KPI={{
  latestAvg:{latest_avg}, firstAvg:{first_avg},
  totalGrowth:{total_growth}, totalVol:{total_vol},
  fastest:{json.dumps(fastest)}, slowest:{json.dumps(slowest)},
  latestDate:"{latest_date.strftime('%B %Y')}",
  firstDate:"{first_date.strftime('%B %Y')}",
}};
const fmt = v => v==null?'—':'£'+Number(v).toLocaleString('en-GB',{{maximumFractionDigits:0}});
const fmtK = v => v==null?'—':'£'+(v/1000).toFixed(0)+'k';
const fmtPct = v => v==null?'—':(v>0?'+':'')+Number(v).toFixed(1)+'%';
const fmtN = v => v==null?'—':Number(v).toLocaleString('en-GB');
Chart.defaults.color='#64748B';
Chart.defaults.borderColor='rgba(71,85,105,0.15)';
Chart.defaults.font.family="'DM Sans',sans-serif";
"""

##############################################################################
# 1. CEO DASHBOARD
##############################################################################
ceo_html = """<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>CEO · UK HPI Market Overview</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,700;0,900;1,400&family=Raleway:wght@300;400;500;600;700&display=swap" rel="stylesheet"/>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0E2138;--surface:#132840;--card:#1B3A5C;--card2:#1f4572;
  --border:rgba(212,160,23,0.16);--border2:rgba(74,107,140,0.35);
  --gold:#D4A017;--gold-l:#EBC75A;--gold-s:#9B7510;
  --navy:#1B3A5C;--navy-l:#4A6B8C;
  --green-l:#5C9A7E;--terra:#C44E2C;
  --text:#F7F9FC;--muted:#94A3B8;
}
body{font-family:'Raleway',sans-serif;background:var(--bg);color:var(--text);min-height:100vh}
header{padding:36px 60px 28px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:flex-end;background:linear-gradient(160deg,#0E2138 0%,#132840 100%)}
.logo{font-family:'Playfair Display',serif;font-size:1.9rem;color:#fff;letter-spacing:-.5px;line-height:1}
.logo span{color:var(--gold)}
.logo-sub{font-size:.65rem;color:var(--muted);letter-spacing:2.5px;text-transform:uppercase;margin-top:5px}
.report-meta{text-align:right;font-size:.7rem;color:var(--muted);line-height:2}
.report-meta strong{color:var(--gold-l);display:block;font-size:.8rem;font-family:'Playfair Display',serif;letter-spacing:1px;text-transform:uppercase}
.subtitle{padding:11px 60px;font-size:.65rem;color:var(--muted);letter-spacing:3px;text-transform:uppercase;border-bottom:1px solid var(--border);background:var(--surface)}
main{padding:40px 60px;max-width:1500px;margin:0 auto}

.hero-kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:0;margin-bottom:40px;border:1px solid var(--border2);border-radius:12px;overflow:hidden}
.hero-kpi{background:var(--card);padding:30px 28px;position:relative;border-right:1px solid var(--border2)}
.hero-kpi:last-child{border-right:none}
.hero-kpi:first-child{background:linear-gradient(135deg,var(--card) 0%,var(--card2) 100%)}
.hero-kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:var(--border2)}
.hero-kpi:first-child::before{background:linear-gradient(90deg,var(--gold-s),var(--gold),var(--gold-l))}
.hkpi-label{font-size:.6rem;letter-spacing:2.5px;text-transform:uppercase;color:var(--muted);margin-bottom:14px;font-weight:600}
.hkpi-value{font-family:'Playfair Display',serif;font-size:2.6rem;color:#fff;line-height:1;font-weight:700}
.hkpi-value.gold{color:var(--gold)}
.hkpi-value.green{color:var(--green-l)}
.hkpi-sub{font-size:.68rem;color:var(--muted);margin-top:10px}

.section-label{font-size:.6rem;letter-spacing:3px;text-transform:uppercase;color:var(--gold);margin-bottom:20px;padding-bottom:8px;border-bottom:1px solid var(--border);font-weight:700}

.grid-2{display:grid;grid-template-columns:2fr 1fr;gap:20px;margin-bottom:24px}
.grid-3{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;margin-bottom:24px}
.card{background:var(--card);border:1px solid var(--border2);border-radius:10px;padding:24px}
.card-title{font-family:'Playfair Display',serif;font-size:.95rem;color:#fff;margin-bottom:4px;font-weight:700}
.card-sub{font-size:.67rem;color:var(--muted);margin-bottom:18px;letter-spacing:.3px}
.chart-wrap{position:relative}

footer{padding:20px 60px;text-align:center;font-size:.67rem;color:var(--muted);border-top:1px solid var(--border);letter-spacing:1px;background:var(--surface)}
</style>
</head><body>
<header>
  <div>
    <div class="logo">UK <span>House Price</span> Index</div>
    <div class="logo-sub">Market Intelligence Report &nbsp;·&nbsp; Metropolitan Areas</div>
  </div>
  <div class="report-meta">
    <strong>CEO Briefing</strong>
    Data through October 2016<br/>6 Metropolitan Regions · Jan 1995 – Oct 2016
  </div>
</header>
<div class="subtitle">Strategic Market Overview &nbsp;·&nbsp; Executive Summary</div>
<main>

  <!-- Hero KPIs -->
  <div class="hero-kpis">
    <div class="hero-kpi">
      <div class="hkpi-label">Market Average Price</div>
      <div class="hkpi-value gold" id="kpi1">—</div>
      <div class="hkpi-sub" id="kpi1s">as at latest month</div>
    </div>
    <div class="hero-kpi">
      <div class="hkpi-label">Total Market Growth</div>
      <div class="hkpi-value green" id="kpi2">—</div>
      <div class="hkpi-sub">1995 → 2016, all regions</div>
    </div>
    <div class="hero-kpi">
      <div class="hkpi-label">Leading Region (growth)</div>
      <div class="hkpi-value" id="kpi3" style="font-size:1.6rem">—</div>
      <div class="hkpi-sub" id="kpi3s">highest appreciation</div>
    </div>
    <div class="hero-kpi">
      <div class="hkpi-label">Monthly Market Volume</div>
      <div class="hkpi-value" id="kpi4">—</div>
      <div class="hkpi-sub" id="kpi4s">transactions, latest month</div>
    </div>
  </div>

  <div class="section-label">Market Trajectory</div>
  <div class="grid-2">
    <div class="card">
      <div class="card-title">Market Average Price — 21 Year Trend</div>
      <div class="card-sub">Annual average across all 6 metropolitan regions · 1995–2016</div>
      <div class="chart-wrap" style="height:280px"><canvas id="ceoTrend"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Regional Growth Ranking</div>
      <div class="card-sub">Total price appreciation 1995 → 2016</div>
      <div id="growthTable" style="margin-top:8px"></div>
    </div>
  </div>

  <div class="section-label">Regional Performance</div>
  <div class="grid-3">
    <div class="card" style="grid-column:1/3">
      <div class="card-title">Regional Price Trajectories</div>
      <div class="card-sub">Average price per region — all years</div>
      <div class="chart-wrap" style="height:240px"><canvas id="ceoRegions"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Current Price Spread</div>
      <div class="card-sub">Latest average price per region</div>
      <div class="chart-wrap" style="height:240px"><canvas id="ceoBar"></canvas></div>
    </div>
  </div>

</main>
<footer>CONFIDENTIAL · CEO LEVEL · UK House Price Index · HPI Admins Dashboard · Generated 2026-05-15</footer>
<script>
""" + js_block() + """
// KPIs
document.getElementById('kpi1').textContent = fmt(KPI.latestAvg);
document.getElementById('kpi1s').textContent = 'as at ' + KPI.latestDate;
document.getElementById('kpi2').textContent = '+' + KPI.totalGrowth + '%';
document.getElementById('kpi3').textContent = KPI.fastest.region;
document.getElementById('kpi3s').textContent = '+' + KPI.fastest.growth + '% since 1995';
document.getElementById('kpi4').textContent = fmtN(KPI.totalVol);
document.getElementById('kpi4s').textContent = 'transactions, ' + KPI.latestDate;

// Growth table
const maxG = Math.max(...REGION_GROWTH_SORTED.map(x=>x.growth));
const gt = document.getElementById('growthTable');
gt.innerHTML = REGION_GROWTH_SORTED.map(r => `
  <div style="display:flex;align-items:center;justify-content:space-between;padding:10px 0;border-bottom:1px solid rgba(201,168,76,0.1)">
    <div style="display:flex;align-items:center;gap:8px;font-size:.82rem">
      <span style="width:8px;height:8px;border-radius:50%;background:${r.color};display:inline-block;flex-shrink:0"></span>
      ${r.region}
    </div>
    <div style="display:flex;align-items:center;gap:12px">
      <div style="width:80px;height:5px;background:rgba(255,255,255,0.06);border-radius:3px;overflow:hidden">
        <div style="width:${(r.growth/maxG*100).toFixed(0)}%;height:100%;background:${r.color};border-radius:3px"></div>
      </div>
      <span style="font-size:.82rem;color:#D4A017;min-width:48px;text-align:right">+${r.growth}%</span>
    </div>
  </div>
`).join('');

// Trend chart (annual)
new Chart(document.getElementById('ceoTrend').getContext('2d'),{
  type:'line',
  data:{
    labels:YEARLY_LABELS,
    datasets:[{
      label:'Market Average',
      data:YEARLY_VALUES,
      borderColor:'#D4A017',
      backgroundColor:'rgba(212,160,23,0.1)',
      borderWidth:2.5,
      pointRadius:3,
      pointBackgroundColor:'#c9a84c',
      fill:true,
      tension:0.4,
    }]
  },
  options:{
    responsive:true,maintainAspectRatio:false,
    plugins:{legend:{display:false},tooltip:{backgroundColor:'#0E2138',borderColor:'rgba(212,160,23,0.35)',borderWidth:1,callbacks:{label:c=>' '+fmt(c.parsed.y)}}},
    scales:{
      x:{grid:{color:'rgba(74,107,140,0.12)'},ticks:{font:{size:11}}},
      y:{grid:{color:'rgba(74,107,140,0.12)'},ticks:{callback:v=>'£'+(v/1000).toFixed(0)+'k',font:{size:11}}}
    }
  }
});

// Regional lines
new Chart(document.getElementById('ceoRegions').getContext('2d'),{
  type:'line',
  data:{
    labels:DATES,
    datasets:REGIONS.map(r=>({
      label:r,data:PRICE_BY_REGION[r],
      borderColor:COLORS[r],borderWidth:2,
      pointRadius:0,tension:0.3,
    }))
  },
  options:{
    responsive:true,maintainAspectRatio:false,
    interaction:{mode:'index',intersect:false},
    plugins:{legend:{display:true,position:'bottom',labels:{boxWidth:10,font:{size:10},color:'#8892a4'}},
      tooltip:{backgroundColor:'#0E2138',borderColor:'rgba(74,107,140,0.4)',borderWidth:1,callbacks:{label:c=>' '+c.dataset.label+': '+fmt(c.parsed.y)}}},
    scales:{
      x:{grid:{color:'rgba(74,107,140,0.1)'},ticks:{maxTicksLimit:10,font:{size:10}}},
      y:{grid:{color:'rgba(74,107,140,0.1)'},ticks:{callback:v=>'£'+(v/1000).toFixed(0)+'k',font:{size:10}}}
    }
  }
});

// Bar — current prices
const latestPrices = REGIONS.map(r=>LATEST_BY_REGION[r].price);
const sortedIdx = [...Array(REGIONS.length).keys()].sort((a,b)=>latestPrices[b]-latestPrices[a]);
new Chart(document.getElementById('ceoBar').getContext('2d'),{
  type:'bar',
  data:{
    labels:sortedIdx.map(i=>REGIONS[i].split(' ')[0]),
    datasets:[{
      data:sortedIdx.map(i=>latestPrices[i]),
      backgroundColor:sortedIdx.map(i=>COLORS[REGIONS[i]]+'99'),
      borderColor:sortedIdx.map(i=>COLORS[REGIONS[i]]),
      borderWidth:1,borderRadius:6,
    }]
  },
  options:{
    responsive:true,maintainAspectRatio:false,
    plugins:{legend:{display:false},tooltip:{backgroundColor:'#0E2138',borderColor:'rgba(74,107,140,0.4)',borderWidth:1,callbacks:{label:c=>' '+fmt(c.parsed.y)}}},
    scales:{
      x:{grid:{display:false},ticks:{font:{size:10}}},
      y:{grid:{color:'rgba(74,107,140,0.1)'},ticks:{callback:v=>'£'+(v/1000).toFixed(0)+'k',font:{size:10}}}
    }
  }
});
</script></body></html>"""

##############################################################################
# 2. EXECUTIVE COMMITTEE DASHBOARD
##############################################################################
exec_html = """<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>ExCo · UK HPI Strategic Analysis</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&display=swap" rel="stylesheet"/>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#F7F9FC;--card:#FFFFFF;--card2:#F0F4F8;--border:#E2E8F0;
  --navy:#1B3A5C;--navy-s:#0E2138;--navy-l:#4A6B8C;
  --green:#2D6A4F;--green-l:#5C9A7E;
  --terra:#C44E2C;--gold:#D4A017;--gold-l:#EBC75A;
  --text:#0F172A;--muted:#64748B;
  --success:#16A34A;--danger:#DC2626;--warn:#F59E0B;
}
body{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text);min-height:100vh}
header{background:var(--navy);padding:16px 40px;display:flex;justify-content:space-between;align-items:center;border-bottom:3px solid var(--gold)}
.brand{font-weight:600;font-size:1rem;color:#fff;letter-spacing:.3px}
.brand em{color:var(--gold-l);font-style:normal}
.badge-exec{background:rgba(235,199,90,0.15);border:1px solid rgba(235,199,90,0.4);color:var(--gold-l);font-size:.65rem;font-weight:700;padding:3px 12px;border-radius:20px;letter-spacing:1.5px;text-transform:uppercase}
.tabs{background:var(--navy-s);border-bottom:1px solid rgba(255,255,255,0.08);padding:0 40px;display:flex;gap:0}
.tab{padding:13px 22px;font-size:.78rem;font-weight:500;color:rgba(255,255,255,0.45);cursor:pointer;border-bottom:3px solid transparent;transition:all .2s;letter-spacing:.3px}
.tab.active{color:#fff;border-bottom-color:var(--gold)}
.tab:hover:not(.active){color:rgba(255,255,255,0.75)}
main{padding:28px 40px;max-width:1500px;margin:0 auto}
.tab-panel{display:none}.tab-panel.active{display:block}

.kpi-strip{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:24px}
.kpi{background:var(--navy);border-radius:10px;padding:16px 18px;position:relative;overflow:hidden}
.kpi::after{content:'';position:absolute;right:-8px;top:-8px;width:56px;height:56px;border-radius:50%;background:rgba(255,255,255,0.04)}
.kpi-l{font-size:.6rem;color:rgba(255,255,255,0.45);text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;font-weight:600}
.kpi-v{font-size:1.5rem;font-weight:700;color:#fff}
.kpi-v.blue{color:var(--gold-l)}.kpi-v.green{color:#5C9A7E}.kpi-v.amber{color:var(--gold)}
.kpi-s{font-size:.66rem;color:rgba(255,255,255,0.38);margin-top:4px}

.g2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
.g3{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:16px}
.g21{display:grid;grid-template-columns:2fr 1fr;gap:16px;margin-bottom:16px}
.card{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:20px;box-shadow:0 1px 3px rgba(15,23,42,0.05)}
.ct{font-size:.8rem;font-weight:600;color:var(--text);margin-bottom:2px}
.cs{font-size:.67rem;color:var(--muted);margin-bottom:14px}
.cw{position:relative}

.region-pills{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px}
.pill{font-size:.67rem;padding:3px 10px;border-radius:20px;border:1px solid var(--border);color:var(--muted);cursor:pointer;transition:all .15s;font-weight:500}
.pill.on{color:#fff;border-color:transparent}
.pill:hover{border-color:var(--navy-l)}

footer{text-align:center;padding:20px;font-size:.65rem;color:var(--muted);border-top:1px solid var(--border);background:var(--card);letter-spacing:.5px}
</style></head><body>
<header>
  <div class="brand">UK <em>House Price Index</em> &nbsp;·&nbsp; Strategic Analysis</div>
  <span class="badge-exec">Executive Committee</span>
</header>
<div class="tabs">
  <div class="tab active" data-tab="market">Market Overview</div>
  <div class="tab" data-tab="regions">Regional Analysis</div>
  <div class="tab" data-tab="segments">Property Segments</div>
  <div class="tab" data-tab="growth">Growth Dynamics</div>
</div>
<main>

<!-- TAB: Market Overview -->
<div class="tab-panel active" id="tab-market">
  <div class="kpi-strip">
    <div class="kpi"><div class="kpi-l">Avg Market Price</div><div class="kpi-v blue" id="e1">—</div><div class="kpi-s" id="e1s">—</div></div>
    <div class="kpi"><div class="kpi-l">21-Year Growth</div><div class="kpi-v green" id="e2">—</div><div class="kpi-s">1995 baseline</div></div>
    <div class="kpi"><div class="kpi-l">Top Region (YoY)</div><div class="kpi-v amber" id="e3" style="font-size:1rem">—</div><div class="kpi-s" id="e3s">—</div></div>
    <div class="kpi"><div class="kpi-l">Total Transactions</div><div class="kpi-v" id="e4">—</div><div class="kpi-s" id="e4s">—</div></div>
    <div class="kpi"><div class="kpi-l">Price Range (regions)</div><div class="kpi-v" id="e5" style="font-size:1rem">—</div><div class="kpi-s">spread latest month</div></div>
  </div>
  <div class="g21">
    <div class="card">
      <div class="ct">Average Price Trend by Region</div>
      <div class="cs">Monthly data 1995–2016</div>
      <div class="region-pills" id="mktPills"></div>
      <div class="cw" style="height:260px"><canvas id="mktTrend"></canvas></div>
    </div>
    <div class="card">
      <div class="ct">Year-on-Year Growth</div>
      <div class="cs">12-month % change — latest readings</div>
      <div class="cw" style="height:260px"><canvas id="mktYoY"></canvas></div>
    </div>
  </div>
  <div class="g2">
    <div class="card">
      <div class="ct">Annual Market Average</div>
      <div class="cs">Combined average across all regions · area chart</div>
      <div class="cw" style="height:200px"><canvas id="mktAnnual"></canvas></div>
    </div>
    <div class="card">
      <div class="ct">Monthly Sales Volume</div>
      <div class="cs">Total transactions per month (stacked)</div>
      <div class="cw" style="height:200px"><canvas id="mktVol"></canvas></div>
    </div>
  </div>
</div>

<!-- TAB: Regional Analysis -->
<div class="tab-panel" id="tab-regions">
  <div class="g3" id="regionCards"></div>
  <div class="g2">
    <div class="card">
      <div class="ct">Regional Price Comparison — All Time</div>
      <div class="cs">Monthly average by region — hover to compare</div>
      <div class="cw" style="height:280px"><canvas id="regAllTime"></canvas></div>
    </div>
    <div class="card">
      <div class="ct">Regional Growth: Start vs End</div>
      <div class="cs">1995 opening price vs 2016 closing price</div>
      <div class="cw" style="height:280px"><canvas id="regStartEnd"></canvas></div>
    </div>
  </div>
</div>

<!-- TAB: Property Segments -->
<div class="tab-panel" id="tab-segments">
  <div class="g21">
    <div class="card">
      <div class="ct">Property Type Prices by Region — Latest</div>
      <div class="cs">Detached · Semi-Detached · Terraced · Flat</div>
      <div class="cw" style="height:300px"><canvas id="segGrouped"></canvas></div>
    </div>
    <div class="card">
      <div class="ct">Market Segment Share</div>
      <div class="cs">Price premium: Detached vs Flat (latest)</div>
      <div class="cw" style="height:300px"><canvas id="segPremium"></canvas></div>
    </div>
  </div>
  <div class="card">
    <div class="ct">Property Type Price Trends — Annual Average</div>
    <div class="cs">All 4 property segments 1995–2016 (market average)</div>
    <div class="cw" style="height:240px"><canvas id="segTrend"></canvas></div>
  </div>
</div>

<!-- TAB: Growth Dynamics -->
<div class="tab-panel" id="tab-growth">
  <div class="g21">
    <div class="card">
      <div class="ct">YoY % Change — All Regions</div>
      <div class="cs">12-month rolling change per region 1996–2016</div>
      <div class="cw" style="height:280px"><canvas id="gYoY"></canvas></div>
    </div>
    <div class="card">
      <div class="ct">Total Appreciation 1995→2016</div>
      <div class="cs">Start price · end price · % growth</div>
      <div id="gTable" style="margin-top:4px"></div>
    </div>
  </div>
</div>

</main>
<footer>EXECUTIVE COMMITTEE · CONFIDENTIAL · UK House Price Index · HPI Admins Dashboard · Generated 2026-05-15</footer>
<script>
""" + js_block() + """
// KPIs
const latestPrices = REGIONS.map(r=>LATEST_BY_REGION[r].price);
const priceSpread = Math.max(...latestPrices) - Math.min(...latestPrices);
document.getElementById('e1').textContent = fmt(KPI.latestAvg);
document.getElementById('e1s').textContent = 'as at ' + KPI.latestDate;
document.getElementById('e2').textContent = '+'+KPI.totalGrowth+'%';
document.getElementById('e3').textContent = KPI.fastest.region;
document.getElementById('e3s').textContent = '+'+KPI.fastest.growth+'% total growth';
document.getElementById('e4').textContent = fmtN(KPI.totalVol);
document.getElementById('e4s').textContent = KPI.latestDate;
document.getElementById('e5').textContent = fmtK(priceSpread);

// Tabs
document.querySelectorAll('.tab').forEach(t=>{
  t.addEventListener('click',()=>{
    document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(x=>x.classList.remove('active'));
    t.classList.add('active');
    document.getElementById('tab-'+t.dataset.tab).classList.add('active');
  });
});

// Region pills (market tab)
const pillsEl = document.getElementById('mktPills');
let activeRegions = new Set(REGIONS);
REGIONS.forEach(r=>{
  const p = document.createElement('div');
  p.className='pill on'; p.textContent=r; p.style.background=COLORS[r]+'22';
  p.style.borderColor=COLORS[r]+'55'; p.style.color=COLORS[r];
  p.addEventListener('click',()=>{
    if(activeRegions.has(r)){activeRegions.delete(r);p.classList.remove('on');p.style.background='';p.style.color='var(--muted)';}
    else{activeRegions.add(r);p.classList.add('on');p.style.background=COLORS[r]+'22';p.style.color=COLORS[r];}
    updateMktTrend();
  });
  pillsEl.appendChild(p);
});

let mktTrendChart;
function buildMktTrendDatasets(){
  return REGIONS.filter(r=>activeRegions.has(r)).map(r=>({
    label:r,data:PRICE_BY_REGION[r],borderColor:COLORS[r],borderWidth:1.5,pointRadius:0,tension:0.3,
  }));
}
function updateMktTrend(){
  mktTrendChart.data.datasets=buildMktTrendDatasets();
  mktTrendChart.update();
}
const optBase={responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
  plugins:{legend:{display:false},tooltip:{backgroundColor:'#1B3A5C',borderColor:'rgba(74,107,140,0.5)',borderWidth:1,titleColor:'#F7F9FC',bodyColor:'#94A3B8'}},
  scales:{x:{grid:{color:'rgba(15,23,42,0.06)'},ticks:{maxTicksLimit:10,font:{size:10}}},
          y:{grid:{color:'rgba(15,23,42,0.06)'},ticks:{font:{size:10}}}}};

mktTrendChart = new Chart(document.getElementById('mktTrend'),{type:'line',
  data:{labels:DATES,datasets:buildMktTrendDatasets()},
  options:{...optBase,scales:{...optBase.scales,y:{...optBase.scales.y,ticks:{callback:v=>'£'+(v/1000).toFixed(0)+'k',font:{size:10}}}},
    plugins:{...optBase.plugins,tooltip:{...optBase.plugins.tooltip,callbacks:{label:c=>' '+c.dataset.label+': '+fmt(c.parsed.y)}}}}
});

// YoY latest bar
const yoyLatest = REGIONS.map(r=>LATEST_BY_REGION[r].yoy);
new Chart(document.getElementById('mktYoY'),{type:'bar',
  data:{labels:REGIONS.map(r=>r.split(' ')[0]),
    datasets:[{data:yoyLatest,backgroundColor:REGIONS.map(r=>COLORS[r]+'99'),borderColor:REGIONS.map(r=>COLORS[r]),borderWidth:1,borderRadius:6}]},
  options:{...optBase,plugins:{...optBase.plugins,tooltip:{...optBase.plugins.tooltip,callbacks:{label:c=>' '+fmtPct(c.parsed.y)}}},
    scales:{...optBase.scales,y:{...optBase.scales.y,ticks:{callback:v=>v+'%',font:{size:10}}}}}
});

// Annual area
new Chart(document.getElementById('mktAnnual'),{type:'line',
  data:{labels:YEARLY_LABELS,datasets:[{label:'Market Avg',data:YEARLY_VALUES,borderColor:'#1B3A5C',backgroundColor:'rgba(27,58,92,0.12)',borderWidth:2,fill:true,tension:0.4,pointRadius:2}]},
  options:{...optBase,plugins:{...optBase.plugins,legend:{display:false},tooltip:{...optBase.plugins.tooltip,callbacks:{label:c=>' '+fmt(c.parsed.y)}}},
    scales:{...optBase.scales,y:{...optBase.scales.y,ticks:{callback:v=>'£'+(v/1000).toFixed(0)+'k',font:{size:10}}}}}
});

// Volume stacked
const volTotal = DATES.map((_,i)=>REGIONS.reduce((s,r)=>s+(VOL_BY_REGION[r][i]||0),0));
new Chart(document.getElementById('mktVol'),{type:'line',
  data:{labels:DATES,datasets:[{label:'Total Vol',data:volTotal,borderColor:'#2D6A4F',backgroundColor:'rgba(45,106,79,0.1)',borderWidth:1.5,fill:true,tension:0.3,pointRadius:0}]},
  options:{...optBase,plugins:{...optBase.plugins,legend:{display:false},tooltip:{...optBase.plugins.tooltip,callbacks:{label:c=>' '+fmtN(c.parsed.y)+' transactions'}}},
    scales:{...optBase.scales,y:{...optBase.scales.y,ticks:{callback:v=>fmtN(v),font:{size:10}}}}}
});

// Region cards
const rc = document.getElementById('regionCards');
REGIONS.forEach(r=>{
  const d = LATEST_BY_REGION[r];
  rc.innerHTML += `<div class="card" style="border-top:3px solid ${d.color}">
    <div style="font-size:.78rem;font-weight:600;color:${d.color};margin-bottom:10px">${r}</div>
    <div style="font-size:1.6rem;font-weight:700;color:#fff;margin-bottom:2px">${fmt(d.price)}</div>
    <div style="font-size:.68rem;color:var(--muted);margin-bottom:12px">avg price · ${KPI.latestDate}</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:.68rem">
      <div><span style="color:var(--muted)">YoY </span><span style="color:${d.yoy>=0?'#5C9A7E':'#DC2626'};font-weight:600">${fmtPct(d.yoy)}</span></div>
      <div><span style="color:var(--muted)">Vol </span><span style="font-weight:600">${fmtN(d.vol)}</span></div>
      <div><span style="color:var(--muted)">Detached </span>${fmtK(d.det)}</div>
      <div><span style="color:var(--muted)">Flat </span>${fmtK(d.flat)}</div>
    </div>
  </div>`;
});

// Regional all-time
new Chart(document.getElementById('regAllTime'),{type:'line',
  data:{labels:DATES,datasets:REGIONS.map(r=>({label:r,data:PRICE_BY_REGION[r],borderColor:COLORS[r],borderWidth:1.5,pointRadius:0,tension:0.3}))},
  options:{...optBase,interaction:{mode:'index',intersect:false},
    plugins:{...optBase.plugins,legend:{display:true,position:'bottom',labels:{boxWidth:10,font:{size:10},color:'#64748B'}},
      tooltip:{...optBase.plugins.tooltip,callbacks:{label:c=>' '+c.dataset.label+': '+fmt(c.parsed.y)}}},
    scales:{...optBase.scales,y:{...optBase.scales.y,ticks:{callback:v=>'£'+(v/1000).toFixed(0)+'k',font:{size:10}}}}}
});

// Start vs end grouped bar
new Chart(document.getElementById('regStartEnd'),{type:'bar',
  data:{labels:REGION_GROWTH_SORTED.map(x=>x.region.split(' ')[0]),
    datasets:[
      {label:'1995 Price',data:REGION_GROWTH_SORTED.map(x=>x.start),backgroundColor:'rgba(27,58,92,0.25)',borderColor:'rgba(74,107,140,0.6)',borderWidth:1,borderRadius:4},
      {label:'2016 Price',data:REGION_GROWTH_SORTED.map(x=>x.end),backgroundColor:REGION_GROWTH_SORTED.map(x=>x.color+'99'),borderColor:REGION_GROWTH_SORTED.map(x=>x.color),borderWidth:1,borderRadius:4},
    ]},
  options:{...optBase,plugins:{...optBase.plugins,legend:{display:true,position:'top',labels:{boxWidth:10,font:{size:10},color:'#64748B'}},
    tooltip:{...optBase.plugins.tooltip,callbacks:{label:c=>' '+c.dataset.label+': '+fmt(c.parsed.y)}}},
    scales:{...optBase.scales,y:{...optBase.scales.y,ticks:{callback:v=>'£'+(v/1000).toFixed(0)+'k',font:{size:10}}}}}
});

// Segment grouped
new Chart(document.getElementById('segGrouped'),{type:'bar',
  data:{labels:REGIONS.map(r=>r.split(' ')[0]),
    datasets:PROP_LABELS.map((l,i)=>({label:l,data:REGIONS.map(r=>PROP_DATA[r][i]),backgroundColor:PROP_COLORS[i]+'88',borderColor:PROP_COLORS[i],borderWidth:1,borderRadius:4}))},
  options:{...optBase,plugins:{...optBase.plugins,legend:{display:true,position:'top',labels:{boxWidth:10,font:{size:10},color:'#64748B'}},
    tooltip:{...optBase.plugins.tooltip,callbacks:{label:c=>' '+c.dataset.label+': '+fmt(c.parsed.y)}}},
    scales:{...optBase.scales,y:{...optBase.scales.y,ticks:{callback:v=>'£'+(v/1000).toFixed(0)+'k',font:{size:10}}}}}
});

// Detached vs Flat premium
const detFlatData = REGIONS.map(r=>({r,premium:PROP_DATA[r][0]-PROP_DATA[r][3],color:COLORS[r]}));
detFlatData.sort((a,b)=>b.premium-a.premium);
new Chart(document.getElementById('segPremium'),{type:'bar',
  data:{labels:detFlatData.map(x=>x.r.split(' ')[0]),
    datasets:[{label:'Detached-Flat Premium',data:detFlatData.map(x=>x.premium),backgroundColor:detFlatData.map(x=>x.color+'99'),borderColor:detFlatData.map(x=>x.color),borderWidth:1,borderRadius:4}]},
  options:{...optBase,indexAxis:'y',plugins:{...optBase.plugins,legend:{display:false},
    tooltip:{...optBase.plugins.tooltip,callbacks:{label:c=>' Premium: '+fmt(c.parsed.x)}}},
    scales:{...optBase.scales,x:{...optBase.scales.x,ticks:{callback:v=>'£'+(v/1000).toFixed(0)+'k',font:{size:10}}}}}
});

// Segment trend annual
new Chart(document.getElementById('segTrend'),{type:'line',
  data:{labels:YEARLY_LABELS,
    datasets:PROP_LABELS.map((l,i)=>({label:l,data:PROP_YEARLY[l],borderColor:PROP_COLORS[i],borderWidth:2,pointRadius:0,tension:0.4}))},
  options:{...optBase,plugins:{...optBase.plugins,legend:{display:true,position:'top',labels:{boxWidth:10,font:{size:10},color:'#64748B'}},
    tooltip:{...optBase.plugins.tooltip,callbacks:{label:c=>' '+c.dataset.label+': '+fmt(c.parsed.y)}}},
    scales:{...optBase.scales,y:{...optBase.scales.y,ticks:{callback:v=>'£'+(v/1000).toFixed(0)+'k',font:{size:10}}}}}
});

// Segment trend legend color fix
// Growth dynamics YoY
new Chart(document.getElementById('gYoY'),{type:'line',
  data:{labels:DATES,datasets:REGIONS.map(r=>({label:r,data:YOY_BY_REGION[r],borderColor:COLORS[r],borderWidth:1.5,pointRadius:0,tension:0.3}))},
  options:{...optBase,interaction:{mode:'index',intersect:false},
    plugins:{...optBase.plugins,legend:{display:true,position:'bottom',labels:{boxWidth:10,font:{size:10},color:'#64748B'}},
      tooltip:{...optBase.plugins.tooltip,callbacks:{label:c=>' '+c.dataset.label+': '+fmtPct(c.parsed.y)}}},
    scales:{...optBase.scales,y:{...optBase.scales.y,ticks:{callback:v=>v+'%',font:{size:10}}}}}
});

// Growth table
const gtEl = document.getElementById('gTable');
gtEl.innerHTML = REGION_GROWTH_SORTED.map(r=>`
  <div style="display:flex;align-items:center;justify-content:space-between;padding:8px 0;border-bottom:1px solid var(--border);font-size:.78rem">
    <span style="color:${r.color};font-weight:500;min-width:140px">${r.region}</span>
    <span style="color:var(--muted)">${fmt(r.start)}</span>
    <span style="color:#fff">→ ${fmt(r.end)}</span>
    <span style="color:#5C9A7E;font-weight:600">+${r.growth}%</span>
  </div>`).join('');
</script></body></html>"""

##############################################################################
# 3. MANAGERS DASHBOARD
##############################################################################
mgr_html = """<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Managers · UK HPI Regional Analysis</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet"/>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#F7F9FC;--card:#FFFFFF;--border:#E2E8F0;
  --navy:#1B3A5C;--navy-s:#0E2138;--navy-l:#4A6B8C;
  --green:#2D6A4F;--green-l:#5C9A7E;
  --terra:#C44E2C;--gold:#D4A017;--gold-l:#EBC75A;
  --text:#0F172A;--muted:#64748B;
  --success:#16A34A;--danger:#DC2626;
}
body{font-family:'IBM Plex Sans',sans-serif;background:var(--bg);color:var(--text);min-height:100vh}
header{background:var(--navy);padding:13px 32px;display:flex;justify-content:space-between;align-items:center;border-bottom:3px solid var(--gold)}
.brand{color:#fff;font-weight:600;font-size:.95rem;letter-spacing:.3px}
.brand em{color:var(--gold-l);font-style:normal}
.badge-mgr{background:rgba(235,199,90,0.15);border:1px solid rgba(235,199,90,0.35);color:var(--gold-l);font-size:.65rem;font-weight:700;padding:3px 10px;border-radius:20px;letter-spacing:1.5px;text-transform:uppercase}

.filterbar{background:var(--navy-s);border-bottom:1px solid rgba(255,255,255,0.08);padding:10px 32px;display:flex;align-items:center;gap:16px;flex-wrap:wrap}
.fb-label{font-size:.65rem;color:rgba(255,255,255,0.45);text-transform:uppercase;letter-spacing:1px;font-weight:500}
select,input[type=range]{background:rgba(255,255,255,0.07);border:1px solid rgba(255,255,255,0.14);color:#E2E8F0;border-radius:6px;padding:5px 10px;font-size:.78rem;font-family:'IBM Plex Sans',sans-serif;cursor:pointer;outline:none}
select:focus,input:focus{border-color:var(--gold-l)}
.year-val{font-size:.78rem;color:var(--gold-l);min-width:36px;font-weight:600}

main{padding:22px 32px}
.kpi-row{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:18px}
.kpi{background:var(--card);border:1px solid var(--border);border-radius:8px;padding:14px 16px;border-left:4px solid var(--kpi-color,var(--navy))}
.kl{font-size:.6rem;color:var(--muted);text-transform:uppercase;letter-spacing:.8px;margin-bottom:4px;font-weight:500}
.kv{font-size:1.3rem;font-weight:600;color:var(--text)}
.ks{font-size:.65rem;color:var(--muted);margin-top:3px}

.g2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}
.g3{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:14px}
.g12{display:grid;grid-template-columns:1fr 2fr;gap:14px;margin-bottom:14px}
.card{background:var(--card);border:1px solid var(--border);border-radius:8px;padding:16px;box-shadow:0 1px 2px rgba(15,23,42,0.04)}
.ct{font-size:.78rem;font-weight:600;color:var(--text);margin-bottom:2px}
.cs{font-size:.65rem;color:var(--muted);margin-bottom:12px}
.cw{position:relative}

.region-status{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:14px}
.rs-card{background:var(--card);border:1px solid var(--border);border-radius:8px;padding:14px;border-top:3px solid var(--rc);box-shadow:0 1px 2px rgba(15,23,42,0.04)}
.rs-name{font-size:.72rem;font-weight:600;color:var(--text);margin-bottom:8px;letter-spacing:.3px}
.rs-price{font-size:1.2rem;font-weight:700;color:var(--text)}
.rs-meta{display:flex;gap:8px;margin-top:6px;font-size:.65rem}
.rs-pill{padding:2px 7px;border-radius:20px;font-weight:600}
.up{background:rgba(22,163,74,0.1);color:var(--success)}
.dn{background:rgba(220,38,38,0.1);color:var(--danger)}
.neutral{background:rgba(27,58,92,0.08);color:var(--navy-l)}

.rank-table{width:100%;border-collapse:collapse;font-size:.76rem}
.rank-table th{font-size:.62rem;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);padding:8px 0;border-bottom:2px solid var(--border);text-align:left;font-weight:600}
.rank-table td{padding:8px 0;border-bottom:1px solid var(--bg);vertical-align:middle}
.rank-table tr:last-child td{border-bottom:none}
.region-tag{display:inline-flex;align-items:center;gap:5px}
.rdot{width:7px;height:7px;border-radius:50%;flex-shrink:0}
.bar-cell{width:120px}
.bar-bg{height:5px;background:var(--bg);border-radius:3px;overflow:hidden}
.bar-fill{height:100%;border-radius:3px}

footer{text-align:center;padding:14px 32px;font-size:.65rem;color:var(--muted);border-top:1px solid var(--border);margin-top:6px;letter-spacing:.5px}
</style></head><body>
<header>
  <div class="brand">UK <em>House Price Index</em> &nbsp;·&nbsp; Regional Performance</div>
  <span class="badge-mgr">Managers</span>
</header>
<div class="filterbar">
  <span class="fb-label">Region</span>
  <select id="mgrRegion"><option value="ALL">All Regions</option></select>
  <span class="fb-label">From</span>
  <input type="range" id="mgrFrom" min="1995" max="2016" value="1995"/>
  <span class="year-val" id="mgrFromV">1995</span>
  <span class="fb-label">To</span>
  <input type="range" id="mgrTo" min="1995" max="2016" value="2016"/>
  <span class="year-val" id="mgrToV">2016</span>
  <span class="fb-label">Property Type</span>
  <select id="mgrProp"><option value="0">All Types</option><option value="0">Detached</option><option value="1">Semi-Detached</option><option value="2">Terraced</option><option value="3">Flat</option></select>
</div>
<main>

<!-- Region status cards -->
<div class="region-status" id="rsCards"></div>

<!-- KPIs -->
<div class="kpi-row">
  <div class="kpi" style="--kpi-color:#1B3A5C"><div class="kl">Market Avg Price</div><div class="kv" id="mk1">—</div><div class="ks" id="mk1s">—</div></div>
  <div class="kpi" style="--kpi-color:#2D6A4F"><div class="kl">Overall Growth</div><div class="kv" id="mk2">—</div><div class="ks">since 1995</div></div>
  <div class="kpi" style="--kpi-color:#D4A017"><div class="kl">Highest Price Region</div><div class="kv" id="mk3" style="font-size:.9rem">—</div><div class="ks" id="mk3s">—</div></div>
  <div class="kpi" style="--kpi-color:#C44E2C"><div class="kl">Lowest Price Region</div><div class="kv" id="mk4" style="font-size:.9rem">—</div><div class="ks" id="mk4s">—</div></div>
  <div class="kpi" style="--kpi-color:#4A6B8C"><div class="kl">Fastest Growing</div><div class="kv" id="mk5" style="font-size:.9rem">—</div><div class="ks" id="mk5s">21yr growth</div></div>
  <div class="kpi" style="--kpi-color:#5C9A7E"><div class="kl">Total Monthly Vol</div><div class="kv" id="mk6">—</div><div class="ks" id="mk6s">—</div></div>
</div>

<div class="g2">
  <div class="card">
    <div class="ct">Average Price Trend</div>
    <div class="cs">Filtered by region & year range · monthly</div>
    <div class="cw" style="height:220px"><canvas id="mgrTrend"></canvas></div>
  </div>
  <div class="card">
    <div class="ct">Monthly Volume by Region</div>
    <div class="cs">Sales transactions — filtered range</div>
    <div class="cw" style="height:220px"><canvas id="mgrVol"></canvas></div>
  </div>
</div>

<div class="g2">
  <div class="card">
    <div class="ct">Year-on-Year % Change</div>
    <div class="cs">12-month rolling price change — filtered</div>
    <div class="cw" style="height:200px"><canvas id="mgrYoY"></canvas></div>
  </div>
  <div class="card">
    <div class="ct">Month-on-Month % Change</div>
    <div class="cs">1-month price change — short-term signals</div>
    <div class="cw" style="height:200px"><canvas id="mgrMoM"></canvas></div>
  </div>
</div>

<div class="g12">
  <div class="card">
    <div class="ct">Property Type Breakdown</div>
    <div class="cs">Avg price by type — latest month</div>
    <div class="cw" style="height:240px"><canvas id="mgrPropType"></canvas></div>
  </div>
  <div class="card">
    <div class="ct">Regional Ranking Table</div>
    <div class="cs">Latest month — sortable comparison</div>
    <table class="rank-table">
      <thead><tr>
        <th>Region</th><th>Avg Price</th><th>Vol</th><th>YoY</th><th>Detached</th><th>Flat</th><th>Price bar</th>
      </tr></thead>
      <tbody id="rankBody"></tbody>
    </table>
  </div>
</div>

</main>
<footer>MANAGERS · UK House Price Index · HPI Admins Dashboard · Generated 2026-05-15</footer>
<script>
""" + js_block() + """
// Populate region selector
const mgrRegEl = document.getElementById('mgrRegion');
REGIONS.forEach(r=>{const o=document.createElement('option');o.value=r;o.textContent=r;mgrRegEl.appendChild(o);});

// KPIs (static)
const byPrice = [...REGIONS].sort((a,b)=>LATEST_BY_REGION[b].price-LATEST_BY_REGION[a].price);
document.getElementById('mk1').textContent = fmt(KPI.latestAvg);
document.getElementById('mk1s').textContent = KPI.latestDate;
document.getElementById('mk2').textContent = '+'+KPI.totalGrowth+'%';
document.getElementById('mk3').textContent = byPrice[0];
document.getElementById('mk3s').textContent = fmt(LATEST_BY_REGION[byPrice[0]].price);
document.getElementById('mk4').textContent = byPrice[byPrice.length-1];
document.getElementById('mk4s').textContent = fmt(LATEST_BY_REGION[byPrice[byPrice.length-1]].price);
document.getElementById('mk5').textContent = KPI.fastest.region.split(' ')[0];
document.getElementById('mk5s').textContent = '+'+KPI.fastest.growth+'% total';
document.getElementById('mk6').textContent = fmtN(KPI.totalVol);
document.getElementById('mk6s').textContent = KPI.latestDate;

// Region status cards
const rsEl = document.getElementById('rsCards');
REGIONS.forEach(r=>{
  const d=LATEST_BY_REGION[r];
  const g=REGION_GROWTH.find(x=>x.region===r);
  rsEl.innerHTML+=`<div class="rs-card" style="--rc:${d.color}">
    <div class="rs-name">${r}</div>
    <div class="rs-price">${fmt(d.price)}</div>
    <div class="rs-meta">
      <span class="rs-pill ${d.yoy>=0?'up':'dn'}">${fmtPct(d.yoy)} YoY</span>
      <span class="rs-pill neutral">${fmtN(d.vol)} sales</span>
    </div>
    <div style="font-size:.67rem;color:var(--muted);margin-top:6px">+${g.growth}% since 1995</div>
  </div>`;
});

// Ranking table
const maxP = Math.max(...REGIONS.map(r=>LATEST_BY_REGION[r].price));
const rankBody = document.getElementById('rankBody');
[...REGIONS].sort((a,b)=>LATEST_BY_REGION[b].price-LATEST_BY_REGION[a].price).forEach(r=>{
  const d=LATEST_BY_REGION[r];
  rankBody.innerHTML+=`<tr>
    <td><span class="region-tag"><span class="rdot" style="background:${d.color}"></span>${r}</span></td>
    <td style="font-weight:600">${fmt(d.price)}</td>
    <td>${fmtN(d.vol)}</td>
    <td style="color:${d.yoy>=0?'#2D6A4F':'#DC2626'};font-weight:600">${fmtPct(d.yoy)}</td>
    <td>${fmt(d.det)}</td>
    <td>${fmt(d.flat)}</td>
    <td class="bar-cell"><div class="bar-bg"><div class="bar-fill" style="width:${(d.price/maxP*100).toFixed(0)}%;background:${d.color}"></div></div></td>
  </tr>`;
});

// Charts with filter
const mgrFromEl=document.getElementById('mgrFrom'), mgrToEl=document.getElementById('mgrTo');
const mgrFromV=document.getElementById('mgrFromV'), mgrToV=document.getElementById('mgrToV');
mgrFromEl.addEventListener('input',()=>{mgrFromV.textContent=mgrFromEl.value;updateMgr();});
mgrToEl.addEventListener('input',()=>{mgrToV.textContent=mgrToEl.value;updateMgr();});
mgrRegEl.addEventListener('change',updateMgr);

const optM={responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
  plugins:{legend:{display:false},tooltip:{backgroundColor:'#1B3A5C',titleColor:'#F7F9FC',bodyColor:'#94A3B8',borderColor:'rgba(74,107,140,0.5)',borderWidth:1}},
  scales:{x:{grid:{color:'rgba(15,23,42,0.05)'},ticks:{maxTicksLimit:8,font:{size:10}}},
          y:{grid:{color:'rgba(15,23,42,0.05)'},ticks:{font:{size:10}}}}};

let mgrCharts={};
function initMgrChart(id,type,opts){
  if(mgrCharts[id]) mgrCharts[id].destroy();
  mgrCharts[id]=new Chart(document.getElementById(id),{type,data:{labels:[],datasets:[]},options:opts});
}
function getIdxs(){
  const y0=parseInt(mgrFromEl.value),y1=parseInt(mgrToEl.value);
  return DATES.map((d,i)=>[parseInt(d.slice(0,4)),i]).filter(([y])=>y>=y0&&y<=y1).map(([,i])=>i);
}
function getRegs(){const s=mgrRegEl.value;return s==='ALL'?REGIONS:[s];}

initMgrChart('mgrTrend','line',{...optM,scales:{...optM.scales,y:{...optM.scales.y,ticks:{callback:v=>'£'+(v/1000).toFixed(0)+'k',font:{size:10}}}}});
initMgrChart('mgrVol','line',{...optM,scales:{...optM.scales,y:{...optM.scales.y,ticks:{callback:v=>fmtN(v),font:{size:10}}}}});
initMgrChart('mgrYoY','line',{...optM,scales:{...optM.scales,y:{...optM.scales.y,ticks:{callback:v=>v+'%',font:{size:10}}}}});
initMgrChart('mgrMoM','line',{...optM,scales:{...optM.scales,y:{...optM.scales.y,ticks:{callback:v=>v+'%',font:{size:10}}}}});

new Chart(document.getElementById('mgrPropType'),{type:'bar',
  data:{labels:PROP_LABELS,
    datasets:REGIONS.map(r=>({label:r,data:PROP_DATA[r],backgroundColor:COLORS[r]+'88',borderColor:COLORS[r],borderWidth:1,borderRadius:4}))},
  options:{...optM,plugins:{...optM.plugins,legend:{display:true,position:'bottom',labels:{boxWidth:8,font:{size:9}}}},
    scales:{...optM.scales,y:{...optM.scales.y,ticks:{callback:v=>'£'+(v/1000).toFixed(0)+'k',font:{size:10}}}}}
});

function updateMgr(){
  const idxs=getIdxs(), regs=getRegs();
  const labels=idxs.map(i=>DATES[i]);
  ['mgrTrend','mgrVol','mgrYoY','mgrMoM'].forEach(id=>{
    const ch=mgrCharts[id];
    ch.data.labels=labels;
    const src={mgrTrend:PRICE_BY_REGION,mgrVol:VOL_BY_REGION,mgrYoY:YOY_BY_REGION,mgrMoM:MOM_BY_REGION}[id];
    ch.data.datasets=regs.map(r=>({label:r,data:idxs.map(i=>src[r][i]),borderColor:COLORS[r],borderWidth:1.5,pointRadius:0,tension:0.3,fill:false}));
    ch.update();
  });
}
updateMgr();
</script></body></html>"""

##############################################################################
# 4. OPERATIONS DASHBOARD
##############################################################################
ops_html = """<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Operations · UK HPI Data Console</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@300;400;500&family=Inter:wght@400;500;600&display=swap" rel="stylesheet"/>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0E2138;--s1:#132840;--s2:#1B3A5C;
  --border:rgba(74,107,140,0.28);--border2:rgba(74,107,140,0.5);
  --text:#E2E8F0;--muted:#94A3B8;
  --gold:#D4A017;--gold-l:#EBC75A;
  --green:#5C9A7E;--terra:#E58367;--navy-l:#4A6B8C;
  --mono:'JetBrains Mono',monospace;
  --success:#16A34A;--danger:#DC2626;
}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;font-size:.85rem}
header{background:var(--s2);border-bottom:3px solid var(--gold);padding:12px 28px;display:flex;justify-content:space-between;align-items:center}
.brand{font-family:var(--mono);font-size:.86rem;color:var(--gold-l);letter-spacing:.5px}
.brand .sep{color:var(--muted);margin:0 5px}
.badge-ops{background:rgba(212,160,23,0.12);border:1px solid rgba(212,160,23,0.35);color:var(--gold-l);font-size:.62rem;font-weight:600;padding:2px 8px;border-radius:4px;font-family:var(--mono);letter-spacing:1px}

.controls{background:var(--s1);border-bottom:1px solid var(--border);padding:10px 28px;display:flex;flex-wrap:wrap;gap:12px;align-items:center}
.ctrl-group{display:flex;align-items:center;gap:6px}
.ctrl-label{font-family:var(--mono);font-size:.62rem;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
select,input[type=range],input[type=text]{background:var(--s2);border:1px solid var(--border);color:var(--text);border-radius:4px;padding:4px 8px;font-size:.72rem;font-family:var(--mono);outline:none;cursor:pointer}
select:focus,input:focus{border-color:var(--gold)}
.yval{font-family:var(--mono);font-size:.75rem;color:var(--gold-l);min-width:32px;font-weight:600}
.sep-v{height:20px;width:1px;background:var(--border);margin:0 4px}

main{padding:18px 28px}

.status-bar{display:grid;grid-template-columns:repeat(7,1fr);gap:8px;margin-bottom:14px}
.stat{background:var(--s1);border:1px solid var(--border);border-radius:6px;padding:10px 12px}
.stat-l{font-family:var(--mono);font-size:.57rem;color:var(--muted);text-transform:uppercase;margin-bottom:4px;letter-spacing:.5px}
.stat-v{font-family:var(--mono);font-size:1rem;font-weight:500;color:var(--text)}
.stat-v.green{color:var(--green)}.stat-v.red{color:var(--terra)}.stat-v.cyan{color:var(--gold-l)}.stat-v.amber{color:var(--gold)}

.charts-row{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:10px}
.charts-row2{display:grid;grid-template-columns:2fr 1fr;gap:10px;margin-bottom:10px}
.card{background:var(--s1);border:1px solid var(--border);border-radius:6px;padding:14px}
.ct{font-family:var(--mono);font-size:.67rem;color:var(--gold-l);margin-bottom:2px;text-transform:uppercase;letter-spacing:.8px;font-weight:500}
.cs{font-size:.65rem;color:var(--muted);margin-bottom:10px}
.cw{position:relative}

.tbl-wrap{overflow-x:auto;overflow-y:auto;max-height:340px;border:1px solid var(--border);border-radius:6px}
.data-table{width:100%;border-collapse:collapse;font-family:var(--mono);font-size:.7rem}
.data-table th{position:sticky;top:0;background:var(--s2);color:var(--muted);font-weight:500;padding:8px 10px;text-align:left;border-bottom:1px solid var(--border2);white-space:nowrap;cursor:pointer;user-select:none;letter-spacing:.3px}
.data-table th:hover{color:var(--gold-l)}
.data-table td{padding:6px 10px;border-bottom:1px solid rgba(74,107,140,0.13);white-space:nowrap}
.data-table tr:hover td{background:rgba(27,58,92,0.35)}
.data-table tr:last-child td{border-bottom:none}
.up{color:var(--green)}.dn{color:var(--terra)}.na{color:var(--muted)}.flag{color:var(--gold)}

.paging{display:flex;align-items:center;gap:8px;padding:10px 0;font-size:.7rem;color:var(--muted);font-family:var(--mono)}
.pg-btn{background:var(--s2);border:1px solid var(--border);color:var(--text);padding:3px 10px;border-radius:4px;cursor:pointer;font-family:var(--mono);font-size:.7rem}
.pg-btn:hover{border-color:var(--gold);color:var(--gold-l)}
.pg-btn:disabled{opacity:.28;cursor:not-allowed}

footer{text-align:center;padding:12px 28px;font-size:.62rem;color:var(--muted);border-top:1px solid var(--border);font-family:var(--mono);letter-spacing:.5px}
</style></head><body>
<header>
  <div class="brand">hpi<span class="sep">//</span>admins<span class="sep">::</span>data-console</div>
  <span class="badge-ops">OPERATIONS</span>
</header>
<div class="controls">
  <div class="ctrl-group"><span class="ctrl-label">Region</span>
    <select id="opsReg"><option value="ALL">ALL</option></select></div>
  <div class="sep-v"></div>
  <div class="ctrl-group"><span class="ctrl-label">From</span>
    <input type="range" id="opsFrom" min="1995" max="2016" value="1995"/>
    <span class="yval" id="opsFromV">1995</span></div>
  <div class="ctrl-group"><span class="ctrl-label">To</span>
    <input type="range" id="opsTo" min="1995" max="2016" value="2016"/>
    <span class="yval" id="opsToV">2016</span></div>
  <div class="sep-v"></div>
  <div class="ctrl-group"><span class="ctrl-label">Search</span>
    <input type="text" id="opsSearch" placeholder="region / date..." style="width:160px"/></div>
  <div class="ctrl-group"><span class="ctrl-label">Rows/page</span>
    <select id="opsPageSize"><option>20</option><option>50</option><option>100</option></select></div>
  <div class="ctrl-group">
    <select id="opsMetric">
      <option value="price">Avg Price</option>
      <option value="vol">Volume</option>
      <option value="yoy">YoY %</option>
      <option value="mom">MoM %</option>
    </select>
  </div>
</div>
<main>

<!-- Status bar -->
<div class="status-bar">
  <div class="stat"><div class="stat-l">Total Rows</div><div class="stat-v cyan" id="os1">—</div></div>
  <div class="stat"><div class="stat-l">Filtered Rows</div><div class="stat-v" id="os2">—</div></div>
  <div class="stat"><div class="stat-l">Data Gaps</div><div class="stat-v amber" id="os3">—</div></div>
  <div class="stat"><div class="stat-l">Avg Price</div><div class="stat-v" id="os4">—</div></div>
  <div class="stat"><div class="stat-l">Avg YoY</div><div class="stat-v green" id="os5">—</div></div>
  <div class="stat"><div class="stat-l">Total Vol</div><div class="stat-v" id="os6">—</div></div>
  <div class="stat"><div class="stat-l">Date Range</div><div class="stat-v" id="os7" style="font-size:.75rem">—</div></div>
</div>

<!-- Charts -->
<div class="charts-row">
  <div class="card">
    <div class="ct">Avg Price — Selected Metric</div>
    <div class="cs">Filtered data · by region</div>
    <div class="cw" style="height:180px"><canvas id="opsTrend"></canvas></div>
  </div>
  <div class="card">
    <div class="ct">YoY % Change</div>
    <div class="cs">12-month rolling · filtered range</div>
    <div class="cw" style="height:180px"><canvas id="opsYoY"></canvas></div>
  </div>
  <div class="card">
    <div class="ct">MoM % Change</div>
    <div class="cs">1-month rolling · short-term signal</div>
    <div class="cw" style="height:180px"><canvas id="opsMoM"></canvas></div>
  </div>
</div>
<div class="charts-row2">
  <div class="card">
    <div class="ct">Sales Volume Trend</div>
    <div class="cs">Transactions per month · filtered</div>
    <div class="cw" style="height:150px"><canvas id="opsVol"></canvas></div>
  </div>
  <div class="card">
    <div class="ct">Property Type — Latest</div>
    <div class="cs">Avg price by segment (filtered region)</div>
    <div class="cw" style="height:150px"><canvas id="opsProp"></canvas></div>
  </div>
</div>

<!-- Data table -->
<div class="card" style="padding:0">
  <div style="padding:12px 14px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:center">
    <span class="ct" style="margin-bottom:0">Raw Data Table</span>
    <span style="font-family:var(--mono);font-size:.68rem;color:var(--muted)" id="tblMeta">—</span>
  </div>
  <div class="tbl-wrap">
    <table class="data-table">
      <thead><tr>
        <th onclick="sortTable('date')">Date ↕</th>
        <th onclick="sortTable('region')">Region ↕</th>
        <th onclick="sortTable('area')">Area ↕</th>
        <th onclick="sortTable('price')">Avg Price ↕</th>
        <th onclick="sortTable('vol')">Volume ↕</th>
        <th onclick="sortTable('yoy')">YoY % ↕</th>
        <th onclick="sortTable('mom')">MoM % ↕</th>
        <th onclick="sortTable('det')">Detached ↕</th>
        <th onclick="sortTable('semi')">Semi-Det ↕</th>
        <th onclick="sortTable('terr')">Terraced ↕</th>
        <th onclick="sortTable('flat')">Flat ↕</th>
      </tr></thead>
      <tbody id="tblBody"></tbody>
    </table>
  </div>
  <div class="paging">
    <button class="pg-btn" id="pgPrev" onclick="changePage(-1)">← Prev</button>
    <span id="pgInfo">Page 1</span>
    <button class="pg-btn" id="pgNext" onclick="changePage(1)">Next →</button>
  </div>
</div>
</main>
<footer>OPERATIONS · ALL DATA CONFIDENTIAL · UK House Price Index · HPIAdmins_v1 · Generated 2026-05-15 · Rows: 1,572 (excl. data gaps)</footer>
<script>
""" + js_block() + """
// Populate region
const opsRegEl=document.getElementById('opsReg');
REGIONS.forEach(r=>{const o=document.createElement('option');o.value=r;o.textContent=r;opsRegEl.appendChild(o);});

// State
let filteredRows=[...OPS_ROWS], sortKey='date', sortDir=1, page=0;
let opsCharts={};

// Controls
const opsFromEl=document.getElementById('opsFrom'),opsToEl=document.getElementById('opsTo');
const opsFromV=document.getElementById('opsFromV'),opsToV=document.getElementById('opsToV');
opsFromEl.addEventListener('input',()=>{opsFromV.textContent=opsFromEl.value;applyFilters();});
opsToEl.addEventListener('input',()=>{opsToV.textContent=opsToEl.value;applyFilters();});
opsRegEl.addEventListener('change',applyFilters);
document.getElementById('opsSearch').addEventListener('input',applyFilters);
document.getElementById('opsPageSize').addEventListener('change',()=>{page=0;renderTable();});
document.getElementById('opsMetric').addEventListener('change',updateOpsCharts);

function applyFilters(){
  const reg=opsRegEl.value, y0=parseInt(opsFromEl.value), y1=parseInt(opsToEl.value);
  const q=document.getElementById('opsSearch').value.toLowerCase();
  filteredRows=OPS_ROWS.filter(r=>{
    const yr=parseInt(r.date.slice(0,4));
    return (reg==='ALL'||r.region===reg) && yr>=y0 && yr<=y1 &&
      (q===''||(r.region.toLowerCase().includes(q)||r.date.includes(q)||r.area.toLowerCase().includes(q)));
  });
  page=0;
  updateStats();
  updateOpsCharts();
  renderTable();
}

function updateStats(){
  const fr=filteredRows;
  const gaps=fr.filter(r=>r.vol===0).length;
  const avgP=fr.length?Math.round(fr.reduce((s,r)=>s+r.price,0)/fr.length):0;
  const yoyVals=fr.filter(r=>r.yoy!=null).map(r=>r.yoy);
  const avgYoY=yoyVals.length?(yoyVals.reduce((s,v)=>s+v,0)/yoyVals.length):0;
  const totalVol=fr.reduce((s,r)=>s+r.vol,0);
  const dates=fr.map(r=>r.date).sort();
  document.getElementById('os1').textContent=fmtN(OPS_ROWS.length);
  document.getElementById('os2').textContent=fmtN(fr.length);
  document.getElementById('os3').textContent=gaps>0?gaps+' rows':'0';
  document.getElementById('os4').textContent=fmt(avgP);
  document.getElementById('os5').textContent=fmtPct(Math.round(avgYoY*10)/10);
  document.getElementById('os6').textContent=fmtN(totalVol);
  document.getElementById('os7').textContent=dates.length?(dates[0]+' → '+dates[dates.length-1]):'—';
}

function getChartData(){
  const reg=opsRegEl.value==='ALL'?REGIONS:[opsRegEl.value];
  const y0=parseInt(opsFromEl.value),y1=parseInt(opsToEl.value);
  const idxs=DATES.map((d,i)=>[parseInt(d.slice(0,4)),i]).filter(([y])=>y>=y0&&y<=y1).map(([,i])=>i);
  return {reg,idxs};
}

const optOps={responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
  plugins:{legend:{display:false},tooltip:{backgroundColor:'#0E2138',borderColor:'rgba(74,107,140,0.5)',borderWidth:1,titleColor:'#E2E8F0',bodyColor:'#94A3B8',titleFont:{family:'JetBrains Mono',size:10},bodyFont:{family:'JetBrains Mono',size:10}}},
  scales:{x:{grid:{color:'rgba(74,107,140,0.1)'},ticks:{maxTicksLimit:8,font:{family:'JetBrains Mono',size:9},color:'#94A3B8'}},
          y:{grid:{color:'rgba(74,107,140,0.1)'},ticks:{font:{family:'JetBrains Mono',size:9},color:'#94A3B8'}}}};

['opsTrend','opsYoY','opsMoM','opsVol'].forEach(id=>{
  opsCharts[id]=new Chart(document.getElementById(id),{type:'line',data:{labels:[],datasets:[]},options:{...optOps}});
});
opsCharts['opsProp']=new Chart(document.getElementById('opsProp'),{type:'bar',
  data:{labels:PROP_LABELS,datasets:[]},
  options:{...optOps,plugins:{...optOps.plugins,legend:{display:true,position:'bottom',labels:{boxWidth:8,font:{family:'JetBrains Mono',size:9},color:'#94A3B8'}}},
    scales:{...optOps.scales,y:{...optOps.scales.y,ticks:{callback:v=>'£'+(v/1000).toFixed(0)+'k',font:{family:'JetBrains Mono',size:9},color:'#94A3B8'}}}}});

function updateOpsCharts(){
  const {reg,idxs}=getChartData();
  const metric=document.getElementById('opsMetric').value;
  const srcMap={price:PRICE_BY_REGION,vol:VOL_BY_REGION,yoy:YOY_BY_REGION,mom:MOM_BY_REGION};
  const labels=idxs.map(i=>DATES[i]);

  // Trend
  opsCharts['opsTrend'].data.labels=labels;
  opsCharts['opsTrend'].data.datasets=reg.map(r=>({label:r,data:idxs.map(i=>srcMap[metric][r][i]),borderColor:COLORS[r],borderWidth:1.5,pointRadius:0,tension:0.3}));
  opsCharts['opsTrend'].update();

  // YoY
  opsCharts['opsYoY'].data.labels=labels;
  opsCharts['opsYoY'].data.datasets=reg.map(r=>({label:r,data:idxs.map(i=>YOY_BY_REGION[r][i]),borderColor:COLORS[r],borderWidth:1.5,pointRadius:0,tension:0.3}));
  opsCharts['opsYoY'].update();

  // MoM
  opsCharts['opsMoM'].data.labels=labels;
  opsCharts['opsMoM'].data.datasets=reg.map(r=>({label:r,data:idxs.map(i=>MOM_BY_REGION[r][i]),borderColor:COLORS[r],borderWidth:1.5,pointRadius:0,tension:0.3}));
  opsCharts['opsMoM'].update();

  // Vol
  opsCharts['opsVol'].data.labels=labels;
  opsCharts['opsVol'].data.datasets=reg.map(r=>({label:r,data:idxs.map(i=>VOL_BY_REGION[r][i]),borderColor:COLORS[r],backgroundColor:COLORS[r]+'18',borderWidth:1.5,pointRadius:0,tension:0.3,fill:true}));
  opsCharts['opsVol'].update();

  // Prop type
  opsCharts['opsProp'].data.datasets=reg.map(r=>({label:r,data:PROP_DATA[r],backgroundColor:COLORS[r]+'66',borderColor:COLORS[r],borderWidth:1,borderRadius:3}));
  opsCharts['opsProp'].update();
}

// Table
let sortKey2='date', sortDir2=1;
function sortTable(key){
  if(sortKey2===key) sortDir2*=-1; else{sortKey2=key;sortDir2=1;}
  filteredRows.sort((a,b)=>{
    const av=a[key]??-Infinity, bv=b[key]??-Infinity;
    return typeof av==='string'?av.localeCompare(bv)*sortDir2:(av-bv)*sortDir2;
  });
  renderTable();
}
function changePage(d){
  const ps=parseInt(document.getElementById('opsPageSize').value);
  const maxP=Math.ceil(filteredRows.length/ps)-1;
  page=Math.max(0,Math.min(maxP,page+d));
  renderTable();
}
function renderTable(){
  const ps=parseInt(document.getElementById('opsPageSize').value);
  const start=page*ps, end=start+ps;
  const slice=filteredRows.slice(start,end);
  const body=document.getElementById('tblBody');
  body.innerHTML=slice.map(r=>`<tr>
    <td>${r.date}</td>
    <td><span style="color:${COLORS[r.region]}">${r.region}</span></td>
    <td>${r.area}</td>
    <td style="font-weight:500">${fmt(r.price)}</td>
    <td>${r.vol===0?'<span class="flag">0 ⚠</span>':fmtN(r.vol)}</td>
    <td class="${r.yoy==null?'na':r.yoy>=0?'up':'dn'}">${r.yoy==null?'—':fmtPct(r.yoy)}</td>
    <td class="${r.mom==null?'na':r.mom>=0?'up':'dn'}">${r.mom==null?'—':fmtPct(r.mom)}</td>
    <td>${fmt(r.det)}</td><td>${fmt(r.semi)}</td><td>${fmt(r.terr)}</td><td>${fmt(r.flat)}</td>
  </tr>`).join('');
  const total=filteredRows.length, pages=Math.ceil(total/ps)||1;
  document.getElementById('pgInfo').textContent=`Page ${page+1} of ${pages} (${fmtN(total)} rows)`;
  document.getElementById('tblMeta').textContent=`Showing ${start+1}–${Math.min(end,total)} of ${fmtN(total)} records`;
  document.getElementById('pgPrev').disabled=page===0;
  document.getElementById('pgNext').disabled=page>=pages-1;
}

applyFilters();
</script></body></html>"""

# ── Fix hardcoded year references to be dynamic ──────────────────────────────
def _fix_years(html):
    return (html
        # Range sliders — critical: prevent user from reaching latest year
        .replace('id="mgrFrom" min="1995" max="2016" value="1995"',
                 f'id="mgrFrom" min="{first_year}" max="{latest_year}" value="{first_year}"')
        .replace('id="mgrTo" min="1995" max="2016" value="2016"',
                 f'id="mgrTo" min="{first_year}" max="{latest_year}" value="{latest_year}"')
        .replace('id="mgrToV">2016<', f'id="mgrToV">{latest_year}<')
        .replace('id="opsFrom" min="1995" max="2016" value="1995"',
                 f'id="opsFrom" min="{first_year}" max="{latest_year}" value="{first_year}"')
        .replace('id="opsTo" min="1995" max="2016" value="2016"',
                 f'id="opsTo" min="{first_year}" max="{latest_year}" value="{latest_year}"')
        .replace('id="opsToV">2016<', f'id="opsToV">{latest_year}<')
        # Text labels — cosmetic but confusing if wrong
        .replace('Data through October 2016', f'Data through {latest_label}')
        .replace('Jan 1995 – Oct 2016', f'Jan {first_year} – {latest_short}')
        .replace('1995 → 2016, all regions', f'{first_year} → {latest_year}, all regions')
        .replace('1995 opening price vs 2016 closing price',
                 f'{first_year} opening price vs {latest_year} closing price')
        .replace('Total Appreciation 1995→2016', f'Total Appreciation {first_year}→{latest_year}')
        .replace("label:'2016 Price'", f"label:'{latest_year} Price'")
        .replace('1995–2016', f'{first_year}–{latest_year}')
        .replace('1996–2016', f'{first_year + 1}–{latest_year}')
        .replace('1995 → 2016', f'{first_year} → {latest_year}')
    )

# ── Write files ─────────────────────────────────────────────────────────────
files = [
    ('CEO_HPIAdmins.html',   ceo_html,  'CEO'),
    ('ExCo_HPIAdmins.html',  exec_html, 'Executive Committee'),
    ('Mgr_HPIAdmins.html',   mgr_html,  'Managers'),
    ('Ops_HPIAdmins.html',   ops_html,  'Operations'),
]
_qa = _qa_widget()
for fname, content, label in files:
    path = os.path.join(FOLDER, fname)
    with open(path, 'w', encoding='utf-8') as f:
        html = _fix_years(content).replace('</body></html>', _qa + '\n</body></html>')
        f.write(html)
    size = os.path.getsize(path) / 1024
    print(f"[{label:22s}] {fname}  ({size:.0f} KB)")

print(f"\nAll dashboards saved to: {FOLDER}")
