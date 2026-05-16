import pandas as pd
import json
import os
import sys
import datetime

_BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_BASE, 'QA_Agent'))
from widget import qa_widget as _qa_widget

df = pd.read_excel(r'C:\MyData\ClaudeProjects\DashboardTest\HPIAdmins_v1.xlsx')
df['Date'] = pd.to_datetime(df['Date'])

regions = sorted(df['RegionName'].unique().tolist())
colors = {
    'Greater Manchester': '#3B82F6',
    'Merseyside':         '#10B981',
    'South Yorkshire':    '#F59E0B',
    'Tyne and Wear':      '#EF4444',
    'West Midlands':      '#8B5CF6',
    'West Yorkshire':     '#EC4899',
}

# ---- Data structures for JS ----

# 1. Monthly avg price per region
dates_all = sorted(df['Date'].dt.strftime('%Y-%m').unique().tolist())

price_by_region = {}
vol_by_region = {}
yoy_by_region = {}
for r in regions:
    rdf = df[df['RegionName'] == r].sort_values('Date')
    price_by_region[r] = rdf['AveragePrice'].round(0).tolist()
    vol_by_region[r]   = rdf['SalesVolume'].tolist()
    yoy_by_region[r]   = rdf['12mPctChange'].round(2).tolist()

# 2. Latest snapshot (last full data month = Oct 2016 since Nov/Dec have 0 vol)
latest_full = df[df['DataGapFlag'] == False]['Date'].max()
latest = df[df['Date'] == latest_full].sort_values('AveragePrice', ascending=False)
latest_labels = latest['RegionName'].tolist()
latest_prices = latest['AveragePrice'].round(0).tolist()
latest_vols   = latest['SalesVolume'].tolist()

# 3. Property type prices - latest snapshot
prop_types  = ['DetachedPrice','SemiDetachedPrice','TerracedPrice','FlatPrice']
prop_labels = ['Detached','Semi-Detached','Terraced','Flat']
prop_data   = {}
for r in regions:
    row = latest[latest['RegionName'] == r]
    if len(row) > 0:
        prop_data[r] = [round(float(row[p].values[0]), 0) for p in prop_types]
    else:
        prop_data[r] = [0]*4

# 4. KPIs
avg_latest_price = round(latest['AveragePrice'].mean(), 0)
max_growth_row = df[(df['Date'] == latest_full) & df['12mPctChange'].notna()].nlargest(1, '12mPctChange')
if len(max_growth_row) > 0:
    max_growth_region = max_growth_row['RegionName'].values[0]
    max_growth_val    = round(float(max_growth_row['12mPctChange'].values[0]), 1)
else:
    max_growth_region = 'N/A'
    max_growth_val    = 0

total_vol = int(latest['SalesVolume'].sum())
price_1995 = df[df['Date'] == df['Date'].min()]['AveragePrice'].mean()
overall_growth = round((avg_latest_price / price_1995 - 1) * 100, 1)
date_label   = latest_full.strftime('%B %Y')
date_label_s = latest_full.strftime('%b %Y')    # e.g. "Dec 2017"
first_year   = df['Date'].min().year
latest_year  = latest_full.year

# ---- Serialize ----
js_data = f"""
const DATES = {json.dumps(dates_all)};
const REGIONS = {json.dumps(regions)};
const COLORS = {json.dumps(colors)};
const PRICE_BY_REGION = {json.dumps(price_by_region)};
const VOL_BY_REGION = {json.dumps(vol_by_region)};
const YOY_BY_REGION = {json.dumps(yoy_by_region)};
const LATEST_LABELS = {json.dumps(latest_labels)};
const LATEST_PRICES = {json.dumps(latest_prices)};
const LATEST_VOLS = {json.dumps(latest_vols)};
const PROP_LABELS = {json.dumps(prop_labels)};
const PROP_DATA = {json.dumps(prop_data)};
const KPI = {{
  avgPrice: {avg_latest_price},
  maxGrowthRegion: "{max_growth_region}",
  maxGrowthVal: {max_growth_val},
  totalVol: {total_vol},
  overallGrowth: {overall_growth},
  dateLabel: "{date_label}"
}};
"""

html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>UK House Price Index — HPI Admins Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600&display=swap');

  :root {
    --bg: #0f1117;
    --surface: #1a1d27;
    --surface2: #222636;
    --border: #2d3148;
    --text: #e8eaf0;
    --muted: #8892a4;
    --accent: #4f8ef7;
    --accent2: #a78bfa;
    --green: #34d399;
    --red: #f87171;
    --amber: #fbbf24;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: 'DM Sans', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
  }

  /* ── Header ── */
  header {
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    padding: 20px 36px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    z-index: 100;
  }
  header h1 {
    font-family: 'DM Serif Display', serif;
    font-size: 1.5rem;
    letter-spacing: -0.3px;
    color: #fff;
  }
  header h1 span { color: var(--accent); }
  .header-meta { font-size: .8rem; color: var(--muted); }

  /* ── Controls ── */
  .controls {
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    padding: 14px 36px;
    display: flex;
    gap: 20px;
    align-items: center;
    flex-wrap: wrap;
  }
  .control-group { display: flex; align-items: center; gap: 8px; }
  .control-group label { font-size: .8rem; color: var(--muted); font-weight: 500; text-transform: uppercase; letter-spacing: .5px; }
  select, input[type=range] {
    background: var(--surface2);
    border: 1px solid var(--border);
    color: var(--text);
    border-radius: 6px;
    padding: 5px 10px;
    font-size: .85rem;
    font-family: 'DM Sans', sans-serif;
    cursor: pointer;
    outline: none;
  }
  select:focus { border-color: var(--accent); }
  #yearDisplay { font-size: .85rem; color: var(--accent); min-width: 60px; }

  /* ── Main layout ── */
  main { padding: 28px 36px; max-width: 1600px; margin: 0 auto; }

  /* ── KPI cards ── */
  .kpi-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 16px;
    margin-bottom: 28px;
  }
  .kpi-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 20px 22px;
    position: relative;
    overflow: hidden;
  }
  .kpi-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    background: var(--accent-line, var(--accent));
    border-radius: 12px 12px 0 0;
  }
  .kpi-label { font-size: .72rem; color: var(--muted); text-transform: uppercase; letter-spacing: .8px; font-weight: 600; margin-bottom: 8px; }
  .kpi-value { font-size: 1.75rem; font-weight: 600; color: #fff; line-height: 1; }
  .kpi-value.green { color: var(--green); }
  .kpi-value.amber { color: var(--amber); }
  .kpi-sub { font-size: .78rem; color: var(--muted); margin-top: 6px; }

  /* ── Chart grid ── */
  .charts-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 20px;
    margin-bottom: 20px;
  }
  .charts-grid.wide { grid-template-columns: 2fr 1fr; }

  .chart-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 22px;
  }
  .chart-card.full { grid-column: 1 / -1; }
  .chart-title {
    font-size: .85rem;
    font-weight: 600;
    color: var(--text);
    margin-bottom: 4px;
    text-transform: uppercase;
    letter-spacing: .5px;
  }
  .chart-subtitle { font-size: .75rem; color: var(--muted); margin-bottom: 16px; }
  .chart-wrap { position: relative; }

  /* ── Region badge legend ── */
  .region-legend {
    display: flex; flex-wrap: wrap; gap: 10px;
    margin-top: 12px;
  }
  .region-badge {
    display: flex; align-items: center; gap: 5px;
    font-size: .73rem; color: var(--muted);
    cursor: pointer;
    padding: 3px 8px;
    border-radius: 20px;
    border: 1px solid var(--border);
    transition: all .2s;
    user-select: none;
  }
  .region-badge:hover { border-color: var(--accent); color: var(--text); }
  .region-badge.active { color: var(--text); border-color: var(--badge-color); }
  .badge-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }

  /* ── Footer ── */
  footer {
    text-align: center;
    color: var(--muted);
    font-size: .72rem;
    padding: 28px 36px;
    border-top: 1px solid var(--border);
    margin-top: 8px;
  }
</style>
</head>
<body>

<header>
  <h1>UK <span>HPI</span> Admins Dashboard</h1>
  <div class="header-meta">House Price Index · Metropolitan Areas · Jan 1995 – Dec 2016</div>
</header>

<div class="controls">
  <div class="control-group">
    <label>Region</label>
    <select id="regionFilter">
      <option value="ALL">All Regions</option>
    </select>
  </div>
  <div class="control-group">
    <label>From year</label>
    <input type="range" id="yearFrom" min="1995" max="2016" value="1995" step="1"/>
    <span id="yearFromDisplay">1995</span>
  </div>
  <div class="control-group">
    <label>To year</label>
    <input type="range" id="yearTo" min="1995" max="2016" value="2016" step="1"/>
    <span id="yearToDisplay">2016</span>
  </div>
  <div class="control-group">
    <label>Property type snapshot</label>
    <select id="propFilter">
      <option value="ALL">All Types</option>
      <option value="DetachedPrice">Detached</option>
      <option value="SemiDetachedPrice">Semi-Detached</option>
      <option value="TerracedPrice">Terraced</option>
      <option value="FlatPrice">Flat</option>
    </select>
  </div>
</div>

<main>

  <!-- KPI Row -->
  <div class="kpi-grid">
    <div class="kpi-card" style="--accent-line: var(--accent)">
      <div class="kpi-label">Avg Price (latest)</div>
      <div class="kpi-value" id="kpiAvgPrice">—</div>
      <div class="kpi-sub" id="kpiPriceDate">—</div>
    </div>
    <div class="kpi-card" style="--accent-line: var(--green)">
      <div class="kpi-label">Price Growth since 1995</div>
      <div class="kpi-value green" id="kpiGrowth">—</div>
      <div class="kpi-sub">Overall across all regions</div>
    </div>
    <div class="kpi-card" style="--accent-line: var(--amber)">
      <div class="kpi-label">Fastest Growing Region</div>
      <div class="kpi-value amber" id="kpiFastRegion" style="font-size:1.2rem">—</div>
      <div class="kpi-sub" id="kpiFastPct">—</div>
    </div>
    <div class="kpi-card" style="--accent-line: var(--accent2)">
      <div class="kpi-label">Total Monthly Sales Volume</div>
      <div class="kpi-value" id="kpiVol">—</div>
      <div class="kpi-sub" id="kpiVolDate">Combined across all regions</div>
    </div>
  </div>

  <!-- Row 1: Price trend (full width) -->
  <div class="charts-grid" style="grid-template-columns: 1fr; margin-bottom:20px">
    <div class="chart-card">
      <div class="chart-title">Average House Price Trend</div>
      <div class="chart-subtitle">Monthly average price per region · use filter to isolate</div>
      <div class="chart-wrap" style="height:280px">
        <canvas id="priceTrendChart"></canvas>
      </div>
      <div class="region-legend" id="priceLegend"></div>
    </div>
  </div>

  <!-- Row 2: YoY + Sales volume -->
  <div class="charts-grid" style="margin-bottom:20px">
    <div class="chart-card">
      <div class="chart-title">Year-on-Year Price Change (%)</div>
      <div class="chart-subtitle">12-month rolling percentage change</div>
      <div class="chart-wrap" style="height:240px">
        <canvas id="yoyChart"></canvas>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title">Monthly Sales Volume</div>
      <div class="chart-subtitle">Transactions per month by region</div>
      <div class="chart-wrap" style="height:240px">
        <canvas id="volumeChart"></canvas>
      </div>
    </div>
  </div>

  <!-- Row 3: Regional snapshot + Property type breakdown -->
  <div class="charts-grid wide" style="margin-bottom:20px">
    <div class="chart-card">
      <div class="chart-title">Property Type Prices — Latest Snapshot</div>
      <div class="chart-subtitle">Detached · Semi-Detached · Terraced · Flat by region</div>
      <div class="chart-wrap" style="height:260px">
        <canvas id="propTypeChart"></canvas>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title">Regional Price Ranking</div>
      <div class="chart-subtitle">Average price — latest month</div>
      <div class="chart-wrap" style="height:260px">
        <canvas id="rankingChart"></canvas>
      </div>
    </div>
  </div>

</main>

<footer>
  UK House Price Index (HPI) · Data: ONS/Land Registry · Metropolitan regions 1995–2016 · HPIAdmins_v1 · Dashboard generated 2026-05-15
</footer>

<script>
""" + js_data + """

// ---- Helpers ----
const fmt = v => v == null ? '' : '£' + Number(v).toLocaleString('en-GB', {maximumFractionDigits:0});
const fmtPct = v => v == null ? '' : (v > 0 ? '+' : '') + v.toFixed(1) + '%';

// ---- Init KPIs ----
document.getElementById('kpiAvgPrice').textContent = fmt(KPI.avgPrice);
document.getElementById('kpiPriceDate').textContent = 'as at ' + KPI.dateLabel;
document.getElementById('kpiGrowth').textContent = '+' + KPI.overallGrowth + '%';
document.getElementById('kpiFastRegion').textContent = KPI.maxGrowthRegion;
document.getElementById('kpiFastPct').textContent = fmtPct(KPI.maxGrowthVal) + ' YoY growth';
document.getElementById('kpiVol').textContent = Number(KPI.totalVol).toLocaleString('en-GB');
document.getElementById('kpiVolDate').textContent = 'as at ' + KPI.dateLabel;

// ---- Populate region filter ----
const regionSel = document.getElementById('regionFilter');
REGIONS.forEach(r => {
  const o = document.createElement('option'); o.value = r; o.textContent = r;
  regionSel.appendChild(o);
});

// ---- Year sliders ----
const yearFromEl = document.getElementById('yearFrom');
const yearToEl   = document.getElementById('yearTo');
const yrFromDisp = document.getElementById('yearFromDisplay');
const yrToDisp   = document.getElementById('yearToDisplay');
yearFromEl.addEventListener('input', () => { yrFromDisp.textContent = yearFromEl.value; updateCharts(); });
yearToEl.addEventListener('input',   () => { yrToDisp.textContent   = yearToEl.value;   updateCharts(); });

// ---- Chart defaults ----
Chart.defaults.color = '#8892a4';
Chart.defaults.borderColor = '#2d3148';
Chart.defaults.font.family = "'DM Sans', sans-serif";

// ---- Chart instances ----
let priceTrendChart, yoyChart, volumeChart, propTypeChart, rankingChart;

function filteredDateIndices() {
  const yr0 = parseInt(yearFromEl.value);
  const yr1 = parseInt(yearToEl.value);
  const idxs = [];
  DATES.forEach((d, i) => {
    const yr = parseInt(d.slice(0,4));
    if (yr >= yr0 && yr <= yr1) idxs.push(i);
  });
  return idxs;
}

function activeRegions() {
  const sel = regionSel.value;
  return sel === 'ALL' ? REGIONS : [sel];
}

function getDataset(key, idxs) {
  return {
    PRICE: PRICE_BY_REGION,
    VOL:   VOL_BY_REGION,
    YOY:   YOY_BY_REGION,
  }[key];
}

function buildLineDatasets(sourceMap, idxs, fill=false) {
  return activeRegions().map(r => ({
    label: r,
    data: idxs.map(i => sourceMap[r][i]),
    borderColor: COLORS[r],
    backgroundColor: COLORS[r] + '18',
    borderWidth: 2,
    pointRadius: 0,
    pointHoverRadius: 4,
    fill: fill,
    tension: 0.3,
  }));
}

// ---- Build region legend ----
function buildLegend() {
  const el = document.getElementById('priceLegend');
  el.innerHTML = '';
  REGIONS.forEach(r => {
    const badge = document.createElement('div');
    badge.className = 'region-badge active';
    badge.style.setProperty('--badge-color', COLORS[r]);
    badge.innerHTML = `<span class="badge-dot" style="background:${COLORS[r]}"></span>${r}`;
    badge.addEventListener('click', () => {
      regionSel.value = (regionSel.value === r) ? 'ALL' : r;
      updateCharts();
    });
    el.appendChild(badge);
  });
}
buildLegend();

const commonOptions = {
  responsive: true, maintainAspectRatio: false,
  interaction: { mode: 'index', intersect: false },
  plugins: { legend: { display: false }, tooltip: { backgroundColor: '#1a1d27', borderColor: '#2d3148', borderWidth: 1, titleFont: { weight: '600' } } },
  scales: {
    x: { grid: { color: '#2d3148' }, ticks: { maxTicksLimit: 12, font: { size: 11 } } },
    y: { grid: { color: '#2d3148' }, ticks: { font: { size: 11 } } }
  }
};

// ---- Price trend chart ----
function initPriceTrend() {
  const idxs = filteredDateIndices();
  const ctx = document.getElementById('priceTrendChart').getContext('2d');
  priceTrendChart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: idxs.map(i => DATES[i]),
      datasets: buildLineDatasets(PRICE_BY_REGION, idxs),
    },
    options: {
      ...commonOptions,
      scales: {
        ...commonOptions.scales,
        y: { ...commonOptions.scales.y, ticks: { ...commonOptions.scales.y.ticks, callback: v => '£' + (v/1000).toFixed(0)+'k' } }
      },
      plugins: { ...commonOptions.plugins, tooltip: { ...commonOptions.plugins.tooltip, callbacks: { label: ctx => ' ' + ctx.dataset.label + ': ' + fmt(ctx.parsed.y) } } }
    }
  });
}

// ---- YoY chart ----
function initYoY() {
  const idxs = filteredDateIndices();
  const ctx = document.getElementById('yoyChart').getContext('2d');
  yoyChart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: idxs.map(i => DATES[i]),
      datasets: buildLineDatasets(YOY_BY_REGION, idxs),
    },
    options: {
      ...commonOptions,
      scales: {
        ...commonOptions.scales,
        y: { ...commonOptions.scales.y, ticks: { callback: v => v + '%' } }
      },
      plugins: { ...commonOptions.plugins, tooltip: { ...commonOptions.plugins.tooltip, callbacks: { label: ctx => ' ' + ctx.dataset.label + ': ' + fmtPct(ctx.parsed.y) } } }
    }
  });
}

// ---- Volume chart ----
function initVolume() {
  const idxs = filteredDateIndices();
  const ctx = document.getElementById('volumeChart').getContext('2d');
  volumeChart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: idxs.map(i => DATES[i]),
      datasets: buildLineDatasets(VOL_BY_REGION, idxs),
    },
    options: {
      ...commonOptions,
      scales: {
        ...commonOptions.scales,
        y: { ...commonOptions.scales.y, ticks: { callback: v => v.toLocaleString() } }
      }
    }
  });
}

// ---- Property type grouped bar ----
function initPropType() {
  const regs = activeRegions();
  const ctx = document.getElementById('propTypeChart').getContext('2d');
  propTypeChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: regs,
      datasets: PROP_LABELS.map((label, pi) => ({
        label,
        data: regs.map(r => PROP_DATA[r][pi]),
        backgroundColor: ['#3B82F6','#10B981','#F59E0B','#EC4899'][pi] + 'cc',
        borderColor:     ['#3B82F6','#10B981','#F59E0B','#EC4899'][pi],
        borderWidth: 1,
        borderRadius: 4,
      })),
    },
    options: {
      ...commonOptions,
      plugins: {
        ...commonOptions.plugins,
        legend: { display: true, position: 'top', labels: { font: { size: 11 }, color: '#8892a4', boxWidth: 12 } },
        tooltip: { ...commonOptions.plugins.tooltip, callbacks: { label: ctx => ' ' + ctx.dataset.label + ': ' + fmt(ctx.parsed.y) } }
      },
      scales: {
        ...commonOptions.scales,
        x: { ...commonOptions.scales.x, ticks: { font: { size: 10 } } },
        y: { ...commonOptions.scales.y, ticks: { callback: v => '£' + (v/1000).toFixed(0)+'k' } }
      }
    }
  });
}

// ---- Ranking horizontal bar ----
function initRanking() {
  const regs = activeRegions();
  const prices = regs.map(r => {
    const idx = LATEST_LABELS.indexOf(r);
    return idx >= 0 ? LATEST_PRICES[idx] : 0;
  });
  const sorted = regs.map((r,i) => ({r, p: prices[i]})).sort((a,b) => a.p - b.p);
  const ctx = document.getElementById('rankingChart').getContext('2d');
  rankingChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: sorted.map(x => x.r),
      datasets: [{
        label: 'Avg Price',
        data: sorted.map(x => x.p),
        backgroundColor: sorted.map(x => COLORS[x.r] + 'bb'),
        borderColor: sorted.map(x => COLORS[x.r]),
        borderWidth: 1,
        borderRadius: 4,
      }]
    },
    options: {
      ...commonOptions,
      indexAxis: 'y',
      plugins: { ...commonOptions.plugins, tooltip: { ...commonOptions.plugins.tooltip, callbacks: { label: ctx => ' ' + fmt(ctx.parsed.x) } } },
      scales: {
        x: { ...commonOptions.scales.x, ticks: { callback: v => '£' + (v/1000).toFixed(0)+'k' } },
        y: { ...commonOptions.scales.y, ticks: { font: { size: 11 } } }
      }
    }
  });
}

// ---- Init all ----
function initCharts() {
  initPriceTrend();
  initYoY();
  initVolume();
  initPropType();
  initRanking();
}

// ---- Update all ----
function updateCharts() {
  const idxs = filteredDateIndices();
  const regs = activeRegions();
  const labels = idxs.map(i => DATES[i]);

  // Price trend
  priceTrendChart.data.labels = labels;
  priceTrendChart.data.datasets = buildLineDatasets(PRICE_BY_REGION, idxs);
  priceTrendChart.update();

  // YoY
  yoyChart.data.labels = labels;
  yoyChart.data.datasets = buildLineDatasets(YOY_BY_REGION, idxs);
  yoyChart.update();

  // Volume
  volumeChart.data.labels = labels;
  volumeChart.data.datasets = buildLineDatasets(VOL_BY_REGION, idxs);
  volumeChart.update();

  // Prop type
  propTypeChart.data.labels = regs;
  propTypeChart.data.datasets = PROP_LABELS.map((label, pi) => ({
    label,
    data: regs.map(r => PROP_DATA[r][pi]),
    backgroundColor: ['#3B82F6','#10B981','#F59E0B','#EC4899'][pi] + 'cc',
    borderColor:     ['#3B82F6','#10B981','#F59E0B','#EC4899'][pi],
    borderWidth: 1, borderRadius: 4,
  }));
  propTypeChart.update();

  // Ranking
  const prices = regs.map(r => { const i = LATEST_LABELS.indexOf(r); return i>=0 ? LATEST_PRICES[i] : 0; });
  const sorted = regs.map((r,i) => ({r, p: prices[i]})).sort((a,b) => a.p - b.p);
  rankingChart.data.labels = sorted.map(x => x.r);
  rankingChart.data.datasets[0].data = sorted.map(x => x.p);
  rankingChart.data.datasets[0].backgroundColor = sorted.map(x => COLORS[x.r] + 'bb');
  rankingChart.data.datasets[0].borderColor = sorted.map(x => COLORS[x.r]);
  rankingChart.update();
}

regionSel.addEventListener('change', updateCharts);

initCharts();
</script>
</body>
</html>"""

# Fix hardcoded year references before writing
def _fix_years(s):
    return (s
        .replace('id="yearFrom" min="1995" max="2016" value="1995"',
                 f'id="yearFrom" min="{first_year}" max="{latest_year}" value="{first_year}"')
        .replace('id="yearTo" min="1995" max="2016" value="2016"',
                 f'id="yearTo" min="{first_year}" max="{latest_year}" value="{latest_year}"')
        .replace('id="yearToDisplay">2016<', f'id="yearToDisplay">{latest_year}<')
        .replace('Jan 1995 – Dec 2016', f'Jan {first_year} – {date_label_s}')
        .replace('1995–2016', f'{first_year}–{latest_year}')
        .replace('1995 – Dec 2016', f'{first_year} – {date_label_s}')
    )

# Create output folder
date_str = os.environ.get('HPI_OUTPUT_DATE', datetime.date.today().strftime('%Y%m%d'))
folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), date_str)
os.makedirs(folder, exist_ok=True)
out_html = os.path.join(folder, 'HPIAdmins.html')
with open(out_html, 'w', encoding='utf-8') as f:
    f.write(_fix_years(html).replace('</body>\n</html>', _qa_widget() + '\n</body>\n</html>'))

print('Dashboard written to:', out_html)
print('Folder:', folder)
