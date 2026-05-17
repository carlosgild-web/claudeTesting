"""
Generates:
  {YYYYMMDD}/Presentation/HPI_Dashboard_Presentation.pptx
  {YYYYMMDD}/Presentation/HPI_Dashboard_UserGuide.pdf
"""
import os, sys, textwrap, datetime
import pandas as pd
import numpy as np

generate_pdf = '--pdf' in sys.argv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
date_str = os.environ.get('HPI_OUTPUT_DATE', datetime.date.today().strftime('%Y%m%d'))

# ── Data ─────────────────────────────────────────────────────────────────────
df = pd.read_excel(os.path.join(BASE_DIR, 'HPIAdmins_v1.xlsx'))
df['Date'] = pd.to_datetime(df['Date'])
dv = df[df['DataGapFlag'] == False].copy()

regions   = sorted(dv['RegionName'].unique())
latest_dt = dv['Date'].max()
first_dt  = dv['Date'].min()
latest    = dv[dv['Date'] == latest_dt]
first_yr  = dv[dv['Date'] == first_dt]

latest_avg  = round(latest['AveragePrice'].mean(), 0)
first_avg   = round(first_yr['AveragePrice'].mean(), 0)
total_growth = round((latest_avg / first_avg - 1) * 100, 1)
total_vol   = int(latest['SalesVolume'].sum())
date_lbl      = latest_dt.strftime('%B %Y')       # e.g. "December 2017"
date_lbl_s    = latest_dt.strftime('%b %Y')       # e.g. "Dec 2017"
first_year    = first_dt.year                     # e.g. 1995
latest_year   = latest_dt.year                    # e.g. 2017
data_period   = f"January {first_year} — {date_lbl}"   # em-dash

reg_growth = sorted([{
    'region': r,
    'start':  round(float(first_yr[first_yr['RegionName']==r]['AveragePrice'].values[0]),0),
    'end':    round(float(latest[latest['RegionName']==r]['AveragePrice'].values[0]),0),
    'growth': round((float(latest[latest['RegionName']==r]['AveragePrice'].values[0]) /
                     float(first_yr[first_yr['RegionName']==r]['AveragePrice'].values[0])-1)*100,1),
    'yoy':    round(float(latest[latest['RegionName']==r]['12mPctChange'].values[0]),1)
              if pd.notna(latest[latest['RegionName']==r]['12mPctChange'].values[0]) else 0,
    'vol':    int(latest[latest['RegionName']==r]['SalesVolume'].values[0]),
    'price':  round(float(latest[latest['RegionName']==r]['AveragePrice'].values[0]),0),
} for r in regions], key=lambda x: x['growth'], reverse=True)

fastest  = reg_growth[0]
highest  = sorted(reg_growth, key=lambda x: x['price'], reverse=True)[0]
lowest   = sorted(reg_growth, key=lambda x: x['price'])[0]

# Property type stats (market avg latest)
prop_avgs = {
    'Detached':      round(latest['DetachedPrice'].mean(),0),
    'Semi-Detached': round(latest['SemiDetachedPrice'].mean(),0),
    'Terraced':      round(latest['TerracedPrice'].mean(),0),
    'Flat':          round(latest['FlatPrice'].mean(),0),
}
det_flat_premium = round((prop_avgs['Detached']/prop_avgs['Flat']-1)*100,1)

price_spread = round(float(latest['AveragePrice'].max() - latest['AveragePrice'].min()),0)

FOLDER = os.path.join(BASE_DIR, date_str)

def fmt(v):  return f"£{v:,.0f}"
def fmtK(v): return f"£{v/1000:.0f}k"
def fmtp(v): return f"+{v:.1f}%" if v >= 0 else f"{v:.1f}%"

print("KPIs computed:")
print(f"  Latest avg price : {fmt(latest_avg)} ({date_lbl})")
print(f"  Total growth     : {fmtp(total_growth)} ({first_year}->{latest_year})")
print(f"  Fastest region   : {fastest['region']} ({fmtp(fastest['growth'])})")
print(f"  Highest price    : {highest['region']} ({fmt(highest['price'])})")
print(f"  Price spread     : {fmt(price_spread)}")
print(f"  Det/Flat premium : {fmtp(det_flat_premium)}")
print(f"  Total volume     : {total_vol:,}")


# ═════════════════════════════════════════════════════════════════════════════
# POWERPOINT
# ═════════════════════════════════════════════════════════════════════════════
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

# ── Color palette ────────────────────────────────────────────────────────────
BG      = RGBColor(0x0f, 0x14, 0x23)
SURFACE = RGBColor(0x17, 0x1e, 0x31)
GOLD    = RGBColor(0xc9, 0xa8, 0x4c)
WHITE   = RGBColor(0xff, 0xff, 0xff)
MUTED   = RGBColor(0x8a, 0x9b, 0xb4)
GREEN   = RGBColor(0x34, 0xd3, 0x99)
RED     = RGBColor(0xf8, 0x71, 0x71)
BLUE    = RGBColor(0x58, 0xa6, 0xff)

AUD_COLORS = {
    'CEO':        (RGBColor(0xc9,0xa8,0x4c), RGBColor(0x1c,0x14,0x04)),
    'ExCo':       (RGBColor(0x58,0xa6,0xff), RGBColor(0x04,0x12,0x24)),
    'Managers':   (RGBColor(0x1d,0x8a,0x6f), RGBColor(0x04,0x1a,0x12)),
    'Operations': (RGBColor(0x89,0xdc,0xeb), RGBColor(0x04,0x18,0x1e)),
}

def add_rect(slide, x, y, w, h, color, alpha=None):
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape

def add_text(slide, text, x, y, w, h, size=14, bold=False, color=WHITE,
             align=PP_ALIGN.LEFT, wrap=True, italic=False):
    txb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf  = txb.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.color.rgb = color
    run.font.italic = italic
    return txb

def set_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_bullet_box(slide, items, x, y, w, h, size=11, color=WHITE, dot_color=None, spacing=0.35):
    """Add a list of bullet items as stacked text boxes."""
    for i, item in enumerate(items):
        dot_color_use = dot_color or GOLD
        # bullet dot
        d = slide.shapes.add_shape(1, Inches(x), Inches(y + i*spacing + 0.06),
                                    Inches(0.08), Inches(0.08))
        d.fill.solid(); d.fill.fore_color.rgb = dot_color_use
        d.line.fill.background()
        # text
        txb = slide.shapes.add_textbox(Inches(x+0.15), Inches(y + i*spacing),
                                        Inches(w-0.15), Inches(spacing+0.02))
        tf = txb.text_frame; tf.word_wrap = True
        p  = tf.paragraphs[0]
        run = p.add_run(); run.text = item
        run.font.size = Pt(size); run.font.color.rgb = color

def add_kpi_box(slide, label, value, x, y, w=2.2, h=1.0,
                bg=SURFACE, label_color=MUTED, val_color=WHITE, val_size=22):
    add_rect(slide, x, y, w, h, bg)
    # left accent bar
    bar = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(0.04), Inches(h))
    bar.fill.solid(); bar.fill.fore_color.rgb = GOLD; bar.line.fill.background()
    add_text(slide, label, x+0.1, y+0.08, w-0.15, 0.25, size=7.5,
             color=label_color, bold=True)
    add_text(slide, value, x+0.1, y+0.28, w-0.15, h-0.3, size=val_size,
             bold=True, color=val_color)

# ── Blank slide layout ───────────────────────────────────────────────────────
blank_layout = prs.slide_layouts[6]

def new_slide():
    sl = prs.slides.add_slide(blank_layout)
    set_bg(sl, BG)
    return sl

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 1 — Title
# ─────────────────────────────────────────────────────────────────────────────
sl = new_slide()
# Top accent bar
add_rect(sl, 0, 0, 13.33, 0.06, GOLD)
# Bottom accent bar
add_rect(sl, 0, 7.44, 13.33, 0.06, GOLD)
# Left decorative block
add_rect(sl, 0, 0.06, 0.5, 7.38, SURFACE)
add_rect(sl, 0.5, 0.06, 0.04, 7.38, GOLD)

# Audience badges
badges = [('CEO', AUD_COLORS['CEO'][0]), ('Executive Committee', AUD_COLORS['ExCo'][0]),
          ('Managers', AUD_COLORS['Managers'][0]), ('Operations', AUD_COLORS['Operations'][0])]
for i,(lbl,clr) in enumerate(badges):
    bx = 1.0 + i*3.0; by = 5.5
    add_rect(sl, bx, by, 2.5, 0.35, clr)
    add_text(sl, lbl, bx+0.1, by+0.04, 2.3, 0.28, size=9, bold=True,
             color=RGBColor(0x0f,0x14,0x23))

add_text(sl, "UK HOUSE PRICE INDEX", 1.0, 1.4, 11.0, 0.6,
         size=13, bold=True, color=GOLD)
add_text(sl, "Dashboard Suite", 1.0, 1.95, 11.0, 1.5,
         size=38, bold=True, color=WHITE)
add_text(sl, "Data-driven insights for every organizational level",
         1.0, 3.4, 10.0, 0.5, size=16, color=MUTED, italic=True)

# Divider line
add_rect(sl, 1.0, 4.0, 8.0, 0.02, MUTED)

add_text(sl, f"6 Metropolitan Regions  ·  Jan {first_year} – {date_lbl_s}  ·  {len(dv):,} records",
         1.0, 4.2, 10.0, 0.4, size=11, color=MUTED)
add_text(sl, "May 2026", 1.0, 4.7, 10.0, 0.4, size=10, color=MUTED)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 2 — The Data
# ─────────────────────────────────────────────────────────────────────────────
sl = new_slide()
add_rect(sl, 0, 0, 13.33, 0.06, GOLD)
add_rect(sl, 0, 0.06, 3.8, 7.44, SURFACE)

add_text(sl, "THE DATA", 0.2, 0.25, 3.2, 0.35, size=8, bold=True, color=GOLD)
add_text(sl, "What we\nare working with", 0.2, 0.6, 3.2, 1.0,
         size=20, bold=True, color=WHITE)
add_rect(sl, 0.2, 1.65, 3.0, 0.02, GOLD)

stats = [
    ("Dataset", "HPIAdmins · HPIAdmins_v1.xlsx"),
    ("Records", f"{len(dv):,} monthly observations"),
    ("Period", data_period),
    ("Regions", "6 Metropolitan Areas"),
    ("Columns", "26 (after cleaning)"),
    ("Property types", "Detached · Semi-Det · Terraced · Flat"),
    ("Metrics", "Avg price · Index · Volume · YoY% · MoM%"),
    ("Data gaps", "12 flagged rows (Nov–Dec 2016 missing vol.)"),
]
for i,(lbl,val) in enumerate(stats):
    y = 1.85 + i*0.65
    add_text(sl, lbl.upper(), 0.2, y, 3.4, 0.22, size=7, bold=True, color=GOLD)
    add_text(sl, val, 0.2, y+0.2, 3.4, 0.38, size=9.5, color=WHITE)

# Right: region cards
add_text(sl, "6 METROPOLITAN REGIONS", 4.2, 0.35, 8.5, 0.3, size=8, bold=True, color=GOLD)
rcolors_list = ['#3B82F6','#10B981','#F59E0B','#EF4444','#8B5CF6','#EC4899']
rarea = {'Greater Manchester':'Northwest','Merseyside':'Northwest',
         'South Yorkshire':'North','Tyne and Wear':'North',
         'West Midlands':'Midlands','West Yorkshire':'North'}
for i,r in enumerate(reg_growth):
    col = i % 3; row = i // 3
    x = 4.2 + col * 2.9; y = 0.85 + row * 1.7
    rgb_hex = rcolors_list[regions.index(r['region'])]
    rgb = RGBColor(int(rgb_hex[1:3],16), int(rgb_hex[3:5],16), int(rgb_hex[5:7],16))
    add_rect(sl, x, y, 2.7, 1.5, SURFACE)
    bar = sl.shapes.add_shape(1, Inches(x), Inches(y), Inches(2.7), Inches(0.04))
    bar.fill.solid(); bar.fill.fore_color.rgb = rgb; bar.line.fill.background()
    add_text(sl, r['region'], x+0.12, y+0.1, 2.5, 0.35, size=9, bold=True, color=WHITE)
    add_text(sl, rarea[r['region']], x+0.12, y+0.42, 2.5, 0.22, size=7.5, color=MUTED)
    add_text(sl, fmt(r['price']), x+0.12, y+0.68, 1.5, 0.35, size=13, bold=True, color=WHITE)
    add_text(sl, f"Growth: {fmtp(r['growth'])}", x+0.12, y+1.05, 2.5, 0.28,
             size=8, color=GREEN, bold=True)
    add_text(sl, f"Vol: {r['vol']:,}", x+1.4, y+1.05, 1.2, 0.28, size=8, color=MUTED)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 3 — Four Dashboards Overview
# ─────────────────────────────────────────────────────────────────────────────
sl = new_slide()
add_rect(sl, 0, 0, 13.33, 0.06, BLUE)
add_text(sl, "FOUR DASHBOARDS · FOUR AUDIENCES", 0.5, 0.2, 12.0, 0.4,
         size=8, bold=True, color=BLUE)
add_text(sl, "One dataset. Tailored for every level.", 0.5, 0.6, 12.0, 0.5,
         size=22, bold=True, color=WHITE)
add_rect(sl, 0.5, 1.15, 12.0, 0.02, MUTED)

cards = [
    ('CEO', 'CEO_HPIAdmins.html', GOLD, RGBColor(0x1c,0x14,0x04),
     'C-Suite / Board',
     ['21-year market trajectory', 'Regional growth ranking', 'Headline KPIs only'],
     'Strategic · Minimal · High impact'),
    ('Executive Committee', 'ExCo_HPIAdmins.html', BLUE, RGBColor(0x04,0x12,0x24),
     'Senior Leadership',
     ['4-tab strategic view', 'Property segment analysis', 'Growth dynamics & benchmarks'],
     'Analytical · Comparative · Tabbed'),
    ('Managers', 'Mgr_HPIAdmins.html', RGBColor(0x1d,0x8a,0x6f), RGBColor(0x04,0x1a,0x12),
     'Regional / Dept Managers',
     ['Region status cards', 'Filterable trend charts', 'MoM + YoY signals + ranking table'],
     'Operational · Interactive · Filterable'),
    ('Operations', 'Ops_HPIAdmins.html', RGBColor(0x89,0xdc,0xeb), RGBColor(0x04,0x18,0x1e),
     'Analysts / Data Teams',
     ['Full 1,572-row data console', 'Real-time search & filter', 'Sortable table + data gap flags'],
     'Granular · Exploratory · Complete'),
]
for i,(name,fname,clr,dark,aud,bullets,tag) in enumerate(cards):
    x = 0.35 + i*3.2; y = 1.4
    add_rect(sl, x, y, 2.9, 5.7, SURFACE)
    add_rect(sl, x, y, 2.9, 0.06, clr)
    add_rect(sl, x, y+0.06, 2.9, 0.6, dark)
    add_text(sl, name.upper(), x+0.12, y+0.1, 2.7, 0.25, size=8, bold=True, color=clr)
    add_text(sl, aud, x+0.12, y+0.35, 2.7, 0.3, size=7.5, color=MUTED)
    add_rect(sl, x, y+0.66, 2.9, 0.02, clr)
    for j,b in enumerate(bullets):
        bx = x+0.1; by = y+0.85+j*0.55
        dot = sl.shapes.add_shape(1, Inches(bx), Inches(by+0.07),
                                   Inches(0.07), Inches(0.07))
        dot.fill.solid(); dot.fill.fore_color.rgb = clr; dot.line.fill.background()
        add_text(sl, b, bx+0.15, by, 2.6, 0.48, size=9, color=WHITE)
    add_rect(sl, x, y+5.35, 2.9, 0.35, dark)
    add_text(sl, fname, x+0.1, y+5.37, 2.7, 0.3, size=7.5, color=clr, italic=True)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDES 4-7 — Individual dashboards
# ─────────────────────────────────────────────────────────────────────────────
dash_slides = [
    {
        'title':    'CEO Dashboard',
        'file':     'CEO_HPIAdmins.html',
        'audience': 'C-Suite / Board',
        'purpose':  'Answers the single question: How is the market performing and where is it heading?',
        'aud_key':  'CEO',
        'features': [
            ('MARKET AVERAGE PRICE', fmt(latest_avg), 'All regions combined · ' + date_lbl),
            ('21-YEAR GROWTH',       fmtp(total_growth), f'From {fmt(first_avg)} in 1995'),
            ('LEADING REGION',       fastest['region'], f"Highest total appreciation {fmtp(fastest['growth'])}"),
            ('MONTHLY VOLUME',       f"{total_vol:,}", f"Transactions · {date_lbl}"),
        ],
        'highlights': [
            f"Market avg grew {fmtp(total_growth)} over 21 years — from {fmtK(first_avg)} to {fmtK(latest_avg)}",
            f"{fastest['region']} leads with {fmtp(fastest['growth'])} total price appreciation",
            f"Price spread across regions: {fmtK(price_spread)} — significant regional divergence",
            "Clean annual trend chart with no noise — built for boardroom screens",
        ],
        'design': 'Dark gold · Playfair serif · No filters needed · Scroll-free',
    },
    {
        'title':    'Executive Committee Dashboard',
        'file':     'ExCo_HPIAdmins.html',
        'audience': 'Senior Leadership',
        'purpose':  'Strategic landscape: regional dynamics, property segments, and growth benchmarks.',
        'aud_key':  'ExCo',
        'features': [
            ('TABS',           '4 Views', 'Market · Regional · Segments · Growth'),
            ('REGIONAL CARDS', '6 cards', 'YoY%, price, volume, property types per region'),
            ('SEGMENT PREMIUM', fmtp(det_flat_premium), 'Detached vs Flat price gap'),
            ('GROWTH TABLE',   'Full ranking', f'{first_year}→{latest_year} start/end/growth per region'),
        ],
        'highlights': [
            "Tab 1 — Market Overview: aggregate trends and YoY summary across all regions",
            f"Tab 2 — Regions: interactive cards showing each region's price, YoY, and property breakdown",
            f"Tab 3 — Segments: Detached commands {fmtp(det_flat_premium)} premium over flats on average",
            "Tab 4 — Growth Dynamics: rolling YoY chart reveals 2007-2008 downturn and recovery",
        ],
        'design': 'Dark blue · Tabbed navigation · Click to explore · Interactive legend',
    },
    {
        'title':    'Managers Dashboard',
        'file':     'Mgr_HPIAdmins.html',
        'audience': 'Regional / Department Managers',
        'purpose':  'Drill into your region. Filter by area, time range, and monitor MoM/YoY signals.',
        'aud_key':  'Managers',
        'features': [
            ('REGION CARDS',   '6 live cards', 'Price, YoY badge, volume — at a glance'),
            ('FILTERS',        '3 controls', 'Region · From year · To year'),
            ('TREND CHARTS',   '4 charts', 'Price · Volume · YoY% · MoM%'),
            ('RANKING TABLE',  'Sortable', 'All regions: price, vol, YoY, Detached, Flat'),
        ],
        'highlights': [
            "Select your region in the top filter bar — all charts update instantly",
            "MoM % chart reveals seasonal patterns and short-term market shocks",
            f"{highest['region']} consistently carries the highest average price ({fmt(highest['price'])})",
            f"Ranking table lets managers benchmark their region against peers at a glance",
        ],
        'design': 'Light/corporate · IBM Plex Sans · Familiar business UI · No technical knowledge needed',
    },
    {
        'title':    'Operations Dashboard',
        'file':     'Ops_HPIAdmins.html',
        'audience': 'Analysts / Data Teams',
        'purpose':  'Full data console. Every metric, every row — filterable, searchable, sortable.',
        'aud_key':  'Operations',
        'features': [
            ('RAW ROWS',      f"{len(dv):,}", 'Fully paginated & searchable data table'),
            ('LIVE FILTERS',  '5 controls', 'Region · Year range · Search · Page size · Metric'),
            ('CHARTS',        '5 live',     'Trend · YoY · MoM · Volume · Property type'),
            ('DATA FLAGS',    '12 flagged', 'Nov–Dec 2016 zero-volume rows highlighted'),
        ],
        'highlights': [
            "Search bar accepts region name, date (YYYY-MM), or area — table filters instantly",
            "Click any column header to sort — ascending/descending toggle",
            "7-stat status bar updates to reflect exactly what is in the current filter scope",
            "Data gap rows (vol = 0) display a ⚠ warning — no surprises in analysis",
        ],
        'design': 'Dark terminal · JetBrains Mono · Built for analysts · Full transparency',
    },
]

for sd in dash_slides:
    sl = new_slide()
    ac, dark = AUD_COLORS[sd['aud_key']]
    add_rect(sl, 0, 0, 13.33, 0.06, ac)

    # Left panel
    add_rect(sl, 0, 0.06, 3.5, 7.44, dark)
    add_text(sl, sd['aud_key'].upper(), 0.2, 0.25, 3.0, 0.3, size=7.5, bold=True, color=ac)
    add_text(sl, sd['title'], 0.2, 0.55, 3.0, 1.0, size=18, bold=True, color=WHITE)
    add_rect(sl, 0.2, 1.6, 2.8, 0.02, ac)
    add_text(sl, "AUDIENCE", 0.2, 1.75, 3.0, 0.22, size=6.5, bold=True, color=MUTED)
    add_text(sl, sd['audience'], 0.2, 1.95, 3.0, 0.3, size=9.5, color=WHITE)
    add_text(sl, "PURPOSE", 0.2, 2.38, 3.0, 0.22, size=6.5, bold=True, color=MUTED)
    add_text(sl, sd['purpose'], 0.2, 2.58, 3.0, 0.9, size=9, color=WHITE, italic=True)
    add_text(sl, "FILE", 0.2, 3.62, 3.0, 0.22, size=6.5, bold=True, color=MUTED)
    add_text(sl, sd['file'], 0.2, 3.82, 3.0, 0.3, size=9, color=ac, italic=True)
    add_rect(sl, 0.2, 4.25, 2.8, 0.02, ac)
    add_text(sl, "DESIGN", 0.2, 4.38, 3.0, 0.22, size=6.5, bold=True, color=MUTED)
    add_text(sl, sd['design'], 0.2, 4.58, 3.0, 0.8, size=8.5, color=MUTED, italic=True)

    # KPI boxes top right
    for i,(lbl,val,sub) in enumerate(sd['features']):
        x = 3.8 + i*2.35; y = 0.2
        add_rect(sl, x, y, 2.2, 1.1, SURFACE)
        bar2 = sl.shapes.add_shape(1, Inches(x), Inches(y),
                                    Inches(2.2), Inches(0.04))
        bar2.fill.solid(); bar2.fill.fore_color.rgb = ac; bar2.line.fill.background()
        add_text(sl, lbl, x+0.1, y+0.08, 2.0, 0.22, size=6.5, bold=True, color=MUTED)
        add_text(sl, val, x+0.1, y+0.28, 2.0, 0.42, size=16, bold=True, color=WHITE)
        add_text(sl, sub, x+0.1, y+0.72, 2.0, 0.3, size=7, color=MUTED)

    # Highlights
    add_text(sl, "KEY OBSERVATIONS", 3.8, 1.55, 9.0, 0.28, size=7.5, bold=True, color=ac)
    add_rect(sl, 3.8, 1.82, 9.0, 0.02, SURFACE)
    for i,h in enumerate(sd['highlights']):
        x = 3.8; y = 2.0 + i * 1.1
        add_rect(sl, x, y, 8.9, 0.95, SURFACE)
        n = sl.shapes.add_shape(1, Inches(x), Inches(y),
                                 Inches(0.5), Inches(0.95))
        n.fill.solid(); n.fill.fore_color.rgb = ac; n.line.fill.background()
        add_text(sl, str(i+1), x+0.12, y+0.22, 0.3, 0.5, size=18, bold=True,
                 color=RGBColor(0x0f,0x14,0x23), align=PP_ALIGN.CENTER)
        add_text(sl, h, x+0.62, y+0.12, 8.1, 0.72, size=10.5, color=WHITE)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 8 — Key Takeaways
# ─────────────────────────────────────────────────────────────────────────────
sl = new_slide()
add_rect(sl, 0, 0, 13.33, 0.06, GREEN)
add_rect(sl, 0, 0.06, 4.2, 7.44, SURFACE)
add_text(sl, "KEY TAKEAWAYS", 0.3, 0.25, 3.5, 0.3, size=8, bold=True, color=GREEN)
add_text(sl, "What the data\ntells us", 0.3, 0.6, 3.5, 0.9,
         size=22, bold=True, color=WHITE)
add_rect(sl, 0.3, 1.6, 3.4, 0.02, GREEN)
add_text(sl, "Across 4 dashboards,\none coherent story:", 0.3, 1.75, 3.5, 0.7,
         size=10, color=MUTED, italic=True)

takeaways = [
    (fmtp(total_growth), "Market Growth",
     f"Average house price rose from {fmtK(first_avg)} ({first_year}) to {fmtK(latest_avg)} ({latest_year})\nacross all 6 metropolitan regions"),
    (fastest['region'].split()[0], "Fastest Grower",
     f"{fastest['region']} led appreciation at {fmtp(fastest['growth'])},\nfollowed by {reg_growth[1]['region']} at {fmtp(reg_growth[1]['growth'])}"),
    (fmtp(det_flat_premium), "Detached Premium",
     f"Detached properties carry a {fmtp(det_flat_premium)} price premium\nover flats — consistent across all regions"),
    (fmt(price_spread), "Regional Spread",
     f"Up to {fmtK(price_spread)} difference between highest\nand lowest priced region in the same month"),
    ("2007–09", "Market Dip",
     "YoY chart (ExCo / Managers) clearly shows\nthe financial crisis impact — and the recovery"),
]
for i,(val,lbl,desc) in enumerate(takeaways):
    x = 4.6; y = 0.35 + i*1.37
    add_rect(sl, x, y, 8.5, 1.2, SURFACE)
    add_rect(sl, x, y, 0.04, 1.2, GREEN)
    add_text(sl, val, x+0.15, y+0.08, 2.0, 0.55, size=20, bold=True, color=GREEN)
    add_text(sl, lbl.upper(), x+0.15, y+0.62, 2.0, 0.28, size=7, bold=True, color=MUTED)
    add_text(sl, desc, x+2.3, y+0.12, 6.0, 0.95, size=10, color=WHITE)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 9 — Next Steps
# ─────────────────────────────────────────────────────────────────────────────
sl = new_slide()
add_rect(sl, 0, 0, 13.33, 0.06, GOLD)
add_rect(sl, 0, 0.06, 13.33, 7.44, SURFACE)
add_text(sl, "NEXT STEPS & RESOURCES", 0.7, 0.3, 11.5, 0.4, size=8, bold=True, color=GOLD)
add_text(sl, "Open your dashboard and explore.", 0.7, 0.7, 10.0, 0.8,
         size=28, bold=True, color=WHITE)
add_rect(sl, 0.7, 1.55, 11.6, 0.02, GOLD)

steps = [
    ("1", "Open in browser", "All dashboards are self-contained HTML files — no server needed. Double-click to open."),
    ("2", "Use the filters",  "Region selector and year sliders narrow all charts simultaneously. Start broad, then drill down."),
    ("3", "Read the User Guide", "The HPI_Dashboard_UserGuide.pdf covers every chart, filter, and metric in detail."),
    ("4", "Share your region",  "Managers: bookmark your filtered URL view or screenshot for team updates."),
]
for i,(num,title,desc) in enumerate(steps):
    x = 0.7 + i*3.15; y = 1.85
    add_rect(sl, x, y, 2.9, 2.8, RGBColor(0x0f,0x14,0x23))
    add_rect(sl, x, y, 2.9, 0.04, GOLD)
    add_text(sl, num, x+0.15, y+0.12, 0.5, 0.6, size=28, bold=True, color=GOLD)
    add_text(sl, title, x+0.15, y+0.75, 2.6, 0.4, size=11, bold=True, color=WHITE)
    add_text(sl, desc, x+0.15, y+1.2, 2.65, 1.4, size=9.5, color=MUTED)

add_text(sl, "FILES IN: 20260515/", 0.7, 5.0, 11.5, 0.3, size=9, color=MUTED)
files_row = [
    ('CEO_HPIAdmins.html', 'CEO', GOLD),
    ('ExCo_HPIAdmins.html', 'ExCo', BLUE),
    ('Mgr_HPIAdmins.html', 'Managers', RGBColor(0x1d,0x8a,0x6f)),
    ('Ops_HPIAdmins.html', 'Operations', RGBColor(0x89,0xdc,0xeb)),
    ('HPI_Dashboard_UserGuide.pdf', 'User Guide', MUTED),
]
for i,(fn,lbl,clr) in enumerate(files_row):
    x = 0.7 + i*2.5; y = 5.4
    add_rect(sl, x, y, 2.35, 0.5, RGBColor(0x0f,0x14,0x23))
    add_rect(sl, x, y, 2.35, 0.03, clr)
    add_text(sl, fn, x+0.08, y+0.08, 2.2, 0.18, size=7, color=clr, italic=True)
    add_text(sl, lbl, x+0.08, y+0.28, 2.2, 0.18, size=7.5, bold=True, color=WHITE)

add_text(sl, "Confidential · UK House Price Index · HPI Admins · Generated 2026-05-15",
         0.7, 6.8, 11.5, 0.3, size=7.5, color=MUTED, align=PP_ALIGN.CENTER)

# Save PPT
ppt_path = os.path.join(FOLDER, 'HPI_Dashboard_Presentation.pptx')
prs.save(ppt_path)
size_kb = os.path.getsize(ppt_path)/1024
print(f"PPT saved: {ppt_path} ({size_kb:.0f} KB)")

if not generate_pdf:
    print("PDF skipped. Pass --pdf to also generate HPI_Dashboard_UserGuide.pdf")
    print("Done.")
    sys.exit(0)

# ═════════════════════════════════════════════════════════════════════════════
# PDF USER GUIDE
# ═════════════════════════════════════════════════════════════════════════════
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, HRFlowable, PageBreak, KeepTogether)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate
from reportlab.pdfgen import canvas as rl_canvas

W, H = A4

# ── Colors ───────────────────────────────────────────────────────────────────
C_NAVY    = colors.HexColor('#0f1423')
C_GOLD2   = colors.HexColor('#c9a84c')
C_BLUE2   = colors.HexColor('#1d6fa4')
C_GREEN2  = colors.HexColor('#166534')
C_SURFACE2= colors.HexColor('#f4f6f9')
C_BORDER2 = colors.HexColor('#dde1e7')
C_TEXT2   = colors.HexColor('#1d2329')
C_MUTED2  = colors.HexColor('#6b7280')
C_CEO2    = colors.HexColor('#c9a84c')
C_EXEC2   = colors.HexColor('#1d6fa4')
C_MGR2    = colors.HexColor('#166534')
C_OPS2    = colors.HexColor('#0e7490')

styles = getSampleStyleSheet()

def sty(name='Normal', **kwargs):
    base = styles.get(name, styles['Normal'])
    return ParagraphStyle(
        name+'_custom_'+str(id(kwargs)),
        parent=base, **kwargs
    )

S_TITLE   = sty('Normal', fontSize=28, textColor=C_NAVY, fontName='Helvetica-Bold',
                 spaceAfter=6, leading=34)
S_SUBTITLE= sty('Normal', fontSize=14, textColor=C_MUTED2, fontName='Helvetica',
                 spaceAfter=12, leading=18)
S_H1      = sty('Normal', fontSize=16, textColor=C_NAVY, fontName='Helvetica-Bold',
                 spaceBefore=18, spaceAfter=6, leading=20)
S_H2      = sty('Normal', fontSize=12, textColor=C_BLUE2, fontName='Helvetica-Bold',
                 spaceBefore=12, spaceAfter=4, leading=15)
S_H3      = sty('Normal', fontSize=10, textColor=C_MUTED2, fontName='Helvetica-Bold',
                 spaceBefore=8, spaceAfter=3, leading=13, textTransform='uppercase')
S_BODY    = sty('Normal', fontSize=9.5, textColor=C_TEXT2, fontName='Helvetica',
                 spaceAfter=5, leading=14, alignment=TA_JUSTIFY)
S_BULLET  = sty('Normal', fontSize=9, textColor=C_TEXT2, fontName='Helvetica',
                 spaceAfter=3, leading=13, leftIndent=14, bulletIndent=0)
S_CODE    = sty('Normal', fontSize=8.5, textColor=C_BLUE2, fontName='Courier',
                 spaceAfter=4, leading=12, backColor=C_SURFACE2)
S_CAPTION = sty('Normal', fontSize=8, textColor=C_MUTED2, fontName='Helvetica-Oblique',
                 spaceAfter=6, leading=10, alignment=TA_CENTER)
S_CALLOUT = sty('Normal', fontSize=9, textColor=C_NAVY, fontName='Helvetica',
                 spaceAfter=4, leading=13, leftIndent=10, backColor=colors.HexColor('#fef9ec'),
                 borderPadding=(6,8,6,8))

def H1(t): return Paragraph(t, S_H1)
def H2(t): return Paragraph(t, S_H2)
def H3(t): return Paragraph(t, S_H3)
def P(t):  return Paragraph(t, S_BODY)
def B(t):  return Paragraph(f"<bullet>&bull;</bullet> {t}", S_BULLET)
def Code(t): return Paragraph(t, S_CODE)
def SP(n=6): return Spacer(1, n)
def HR(clr=C_BORDER2, w=0.5): return HRFlowable(width='100%', thickness=w, color=clr,
                                                   spaceAfter=8, spaceBefore=4)
def Callout(t): return Paragraph(f"<b>Tip:</b> {t}", S_CALLOUT)

def kpi_table(rows):
    data = [[Paragraph(f"<b>{k}</b>", sty('Normal', fontSize=9, textColor=C_MUTED2,
                                           fontName='Helvetica-Bold')),
             Paragraph(v, sty('Normal', fontSize=11, textColor=C_TEXT2,
                               fontName='Helvetica-Bold', leading=14))]
            for k,v in rows]
    t = Table(data, colWidths=[5.5*cm, 10.5*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0),(-1,-1), C_SURFACE2),
        ('GRID',       (0,0),(-1,-1), 0.5, C_BORDER2),
        ('ROWBACKGROUNDS', (0,0),(-1,-1), [colors.white, C_SURFACE2]),
        ('VALIGN',     (0,0),(-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0),(-1,-1), 5),
        ('BOTTOMPADDING', (0,0),(-1,-1), 5),
        ('LEFTPADDING', (0,0),(-1,-1), 8),
    ]))
    return t

def step_table(steps):
    data = [[Paragraph(f"<b>{i+1}</b>", sty('Normal', fontSize=11, textColor=C_GOLD2,
                                              fontName='Helvetica-Bold', alignment=TA_CENTER)),
             Paragraph(f"<b>{title}</b><br/><font size='8' color='#6b7280'>{desc}</font>",
                       sty('Normal', fontSize=9.5, leading=14))]
            for i,(title,desc) in enumerate(steps)]
    t = Table(data, colWidths=[1.0*cm, 15.0*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0),(0,-1), C_NAVY),
        ('BACKGROUND', (1,0),(1,-1), colors.white),
        ('GRID',       (0,0),(-1,-1), 0.5, C_BORDER2),
        ('VALIGN',     (0,0),(-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0),(-1,-1), 7),
        ('BOTTOMPADDING', (0,0),(-1,-1), 7),
        ('LEFTPADDING', (1,0),(1,-1), 10),
    ]))
    return t

def aud_header(name, file, color_hex, desc):
    clr = colors.HexColor(color_hex)
    data = [[Paragraph(f"<b>{name}</b>", sty('Normal', fontSize=14, textColor=colors.white,
                                              fontName='Helvetica-Bold')),
             Paragraph(f"<i>{file}</i>", sty('Normal', fontSize=9, textColor=clr,
                                              fontName='Helvetica-Oblique', alignment=TA_RIGHT))],
            [Paragraph(desc, sty('Normal', fontSize=9, textColor=colors.HexColor('#dde1e7'),
                                  leading=13)), '']]
    t = Table(data, colWidths=[11*cm, 5*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0),(-1,-1), C_NAVY),
        ('SPAN', (0,1),(1,1)),
        ('VALIGN', (0,0),(-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0),(-1,-1), 8),
        ('BOTTOMPADDING', (0,0),(-1,-1), 8),
        ('LEFTPADDING', (0,0),(-1,-1), 12),
        ('RIGHTPADDING', (0,0),(-1,-1), 12),
        ('LINEBELOW', (0,0),(-1,0), 2, clr),
    ]))
    return t

# ── Page template ─────────────────────────────────────────────────────────────
pdf_path = os.path.join(FOLDER, 'HPI_Dashboard_UserGuide.pdf')

def on_page(canvas, doc):
    canvas.saveState()
    # Header bar
    canvas.setFillColor(C_NAVY)
    canvas.rect(0, H-1.2*cm, W, 1.2*cm, fill=1, stroke=0)
    canvas.setFillColor(C_GOLD2)
    canvas.rect(0, H-1.2*cm, W, 0.15*cm, fill=1, stroke=0)
    canvas.setFont('Helvetica-Bold', 8)
    canvas.setFillColor(colors.white)
    canvas.drawString(1.5*cm, H-0.9*cm, "UK HOUSE PRICE INDEX  ·  Dashboard User Guide")
    canvas.setFont('Helvetica', 7.5)
    canvas.setFillColor(colors.HexColor('#8a9bb4'))
    canvas.drawRightString(W-1.5*cm, H-0.9*cm, "Confidential  ·  2026-05-15")
    # Footer
    canvas.setFillColor(C_BORDER2)
    canvas.rect(0, 0, W, 0.8*cm, fill=1, stroke=0)
    canvas.setFont('Helvetica', 7.5)
    canvas.setFillColor(C_MUTED2)
    canvas.drawCentredString(W/2, 0.25*cm, f"Page {doc.page}")
    canvas.restoreState()

doc = SimpleDocTemplate(
    pdf_path, pagesize=A4,
    leftMargin=1.5*cm, rightMargin=1.5*cm,
    topMargin=1.8*cm, bottomMargin=1.2*cm,
    title="HPI Dashboard User Guide",
    author="Data Analytics Team",
)

story = []

# ── COVER ────────────────────────────────────────────────────────────────────
story.append(SP(80))
story.append(Paragraph("UK House Price Index", sty('Normal', fontSize=32, textColor=C_NAVY,
                        fontName='Helvetica-Bold', alignment=TA_CENTER, leading=38)))
story.append(Paragraph("Dashboard Suite — User Guide", sty('Normal', fontSize=20,
                        textColor=C_GOLD2, fontName='Helvetica', alignment=TA_CENTER,
                        spaceAfter=16, leading=26)))
story.append(HR(C_GOLD2, 1.5))
story.append(SP(10))
story.append(Paragraph("Comprehensive instructions for navigating and interpreting<br/>"
                        "the four organizational-level dashboards",
                        sty('Normal', fontSize=11, textColor=C_MUTED2, fontName='Helvetica-Oblique',
                            alignment=TA_CENTER, leading=16)))
story.append(SP(30))

cover_data = [
    ["Dataset", "HPIAdmins · UK Metropolitan House Prices"],
    ["Period",  data_period],
    ["Regions", "6 Metropolitan Areas (Greater Manchester, Merseyside, South Yorkshire,\nTyne and Wear, West Midlands, West Yorkshire)"],
    ["Records", f"{len(dv):,} cleaned observations"],
    ["Dashboards","4 HTML files — CEO · Executive Committee · Managers · Operations"],
    ["Generated", "15 May 2026"],
]
ct = Table([[Paragraph(f"<b>{k}</b>", sty('Normal', fontSize=9, textColor=C_MUTED2,
                                           fontName='Helvetica-Bold')),
             Paragraph(v, sty('Normal', fontSize=9, leading=13))]
            for k,v in cover_data], colWidths=[4*cm, 12*cm])
ct.setStyle(TableStyle([
    ('BACKGROUND', (0,0),(-1,-1), C_SURFACE2),
    ('GRID', (0,0),(-1,-1), 0.5, C_BORDER2),
    ('ROWBACKGROUNDS', (0,0),(-1,-1), [colors.white, C_SURFACE2]),
    ('VALIGN', (0,0),(-1,-1), 'MIDDLE'),
    ('TOPPADDING', (0,0),(-1,-1), 6), ('BOTTOMPADDING', (0,0),(-1,-1), 6),
    ('LEFTPADDING', (0,0),(-1,-1), 8),
]))
story.append(ct)
story.append(PageBreak())

# ── 1. INTRODUCTION ──────────────────────────────────────────────────────────
story.append(H1("1. Introduction"))
story.append(HR())
story.append(P(
    "This user guide provides detailed instructions for all four HPI Admins dashboards. "
    "Each dashboard was designed for a specific organizational level, presenting the same "
    "underlying dataset through a lens suited to that audience's decisions and questions."
))
story.append(P(
    "The dashboards are standalone HTML files. No internet connection is required after the "
    "initial load (Chart.js is loaded from a CDN on first open — subsequent offline use "
    "works if the browser has cached the library)."
))
story.append(SP(8))
story.append(H2("Dashboard Overview"))
dash_overview = [
    ["File", "Audience", "Purpose", "Key Feature"],
    ["CEO_HPIAdmins.html",  "CEO / Board",        "Strategic market overview",    "Headline KPIs + 21-yr trend"],
    ["ExCo_HPIAdmins.html", "Executive Committee","Regional & segment strategy",  "4 navigable tabs"],
    ["Mgr_HPIAdmins.html",  "Managers",           "Regional operational view",    "Filterable by region & year"],
    ["Ops_HPIAdmins.html",  "Operations / Analysts","Full data console",          "1,572-row searchable table"],
]
ot = Table(dash_overview, colWidths=[5.2*cm, 3.5*cm, 4.5*cm, 4.0*cm])
ot.setStyle(TableStyle([
    ('BACKGROUND', (0,0),(-1,0), C_NAVY),
    ('TEXTCOLOR', (0,0),(-1,0), colors.white),
    ('FONTNAME', (0,0),(-1,0), 'Helvetica-Bold'),
    ('FONTSIZE', (0,0),(-1,-1), 8.5),
    ('GRID', (0,0),(-1,-1), 0.5, C_BORDER2),
    ('ROWBACKGROUNDS', (0,1),(-1,-1), [colors.white, C_SURFACE2]),
    ('VALIGN', (0,0),(-1,-1), 'MIDDLE'),
    ('TOPPADDING', (0,0),(-1,-1), 5), ('BOTTOMPADDING', (0,0),(-1,-1), 5),
    ('LEFTPADDING', (0,0),(-1,-1), 7),
]))
story.append(ot)
story.append(SP(8))
story.append(Callout("All four dashboards open directly in any modern browser (Chrome, Edge, Firefox, Safari). "
                      "No installation or plugins required."))

# ── 2. DATA OVERVIEW ─────────────────────────────────────────────────────────
story.append(PageBreak())
story.append(H1("2. Data Overview"))
story.append(HR())
story.append(P(
    "The source data is the UK House Price Index (HPI), covering six metropolitan areas in England. "
    "The data was sourced from ONS/Land Registry records and cleaned into <b>HPIAdmins_v1.xlsx</b>."
))
story.append(H2("Regions Covered"))
reg_tbl = [["Region", "Area", f"Price (Jan {first_year})", f"Price ({date_lbl_s})", "Total Growth"]]
for r in reg_growth:
    reg_tbl.append([r['region'], rarea[r['region']], fmt(r['start']), fmt(r['end']), fmtp(r['growth'])])
rt = Table(reg_tbl, colWidths=[5*cm, 2.5*cm, 3*cm, 3*cm, 2.5*cm])
rt.setStyle(TableStyle([
    ('BACKGROUND', (0,0),(-1,0), C_NAVY),
    ('TEXTCOLOR', (0,0),(-1,0), colors.white),
    ('FONTNAME', (0,0),(-1,0), 'Helvetica-Bold'),
    ('FONTSIZE', (0,0),(-1,-1), 8.5),
    ('GRID', (0,0),(-1,-1), 0.5, C_BORDER2),
    ('ROWBACKGROUNDS', (0,1),(-1,-1), [colors.white, C_SURFACE2]),
    ('VALIGN', (0,0),(-1,-1), 'MIDDLE'),
    ('TOPPADDING', (0,0),(-1,-1), 5), ('BOTTOMPADDING', (0,0),(-1,-1), 5),
    ('LEFTPADDING', (0,0),(-1,-1), 7),
    ('ALIGN', (2,0),(-1,-1), 'RIGHT'),
]))
story.append(rt)
story.append(SP(8))
story.append(H2("Data Cleaning Notes"))
story.append(B("<b>Dropped columns:</b> IndexSA and AveragePriceSA — 100% null across all records."))
story.append(B("<b>Column renaming:</b> % characters removed from column names (e.g. 12m%Change → 12mPctChange) for compatibility."))
story.append(B("<b>DataGapFlag column added:</b> 12 rows (Nov–Dec 2016, all 6 regions) have SalesVolume = 0. These are flagged True and visible in the Operations dashboard."))
story.append(B("<b>Partial nulls retained:</b> 1m and 12m % change columns are null for the first 1 and 12 months respectively — this is expected (no prior period to compare against)."))
story.append(SP(8))
story.append(H2("Key Market Statistics"))
story.append(kpi_table([
    (f"Average price ({date_lbl_s})", f"{fmt(latest_avg)} across all 6 regions"),
    (f"Average price (Jan {first_year})", f"{fmt(first_avg)} across all 6 regions"),
    ("Overall market growth",    f"{fmtp(total_growth)} over {latest_year - first_year} years"),
    ("Fastest growing region",   f"{fastest['region']} at {fmtp(fastest['growth'])} total"),
    (f"Highest priced region",   f"{highest['region']} at {fmt(highest['price'])} ({date_lbl_s})"),
    ("Detached vs Flat premium", f"{fmtp(det_flat_premium)} on average (latest month)"),
    ("Monthly transaction volume", f"{total_vol:,} combined transactions ({date_lbl_s})"),
    ("Price spread (region gap)", f"{fmt(price_spread)} between highest and lowest region"),
]))

# ── 3. CEO DASHBOARD ─────────────────────────────────────────────────────────
story.append(PageBreak())
story.append(aud_header("CEO Dashboard", "CEO_HPIAdmins.html", "#c9a84c",
    "Designed for C-Suite executives. Provides a definitive strategic overview with no operational noise."))
story.append(SP(10))
story.append(H2("Purpose"))
story.append(P(
    "This dashboard answers one question: <i>How is the market performing and where is it heading?</i> "
    "It presents the minimum information needed for a strategic decision or board-level discussion. "
    "No filters, no tabs — just the signal."
))
story.append(H2("Sections at a Glance"))
story.append(B("<b>4 Headline KPIs (top row):</b> Market average price, 21-year total growth, leading region by appreciation, and combined monthly transaction volume."))
story.append(B(f"<b>21-Year Annual Trend chart:</b> The market-wide average price per year from {first_year} to {latest_year}. Gold area chart highlights the sustained growth trajectory and the 2008 dip."))
story.append(B("<b>Regional Growth Ranking:</b> A horizontal bar visual showing which region appreciated most. Length of bar = relative growth magnitude."))
story.append(B("<b>Regional Price Trajectories:</b> All 6 regions on one chart — showing convergence, divergence, and relative positioning over time."))
story.append(B("<b>Current Price Spread bar:</b> Latest average price per region sorted from highest to lowest — instant competitive ranking."))
story.append(SP(8))
story.append(H2("How to Use"))
story.append(step_table([
    ("Open the file", "Double-click CEO_HPIAdmins.html. It opens in your default browser."),
    ("Read the KPIs", f"Top row shows the four headline numbers. All reference {date_lbl} (latest full data month)."),
    ("Scan the trend", "The gold area chart shows annual averages. The slope steepness tells the growth story at a glance."),
    ("Check regional ranking", "The growth bar table on the right ranks regions by total appreciation — useful for portfolio discussions."),
    ("Compare current prices", "The sorted bar chart (bottom right) shows today's price hierarchy across regions."),
]))
story.append(SP(6))
story.append(Callout(f"Key talking point: The market average grew {fmtp(total_growth)} in 21 years — "
                      f"from {fmtK(first_avg)} to {fmtK(latest_avg)}. "
                      f"{fastest['region']} leads all regions at {fmtp(fastest['growth'])} total appreciation."))

# ── 4. EXCO DASHBOARD ────────────────────────────────────────────────────────
story.append(PageBreak())
story.append(aud_header("Executive Committee Dashboard", "ExCo_HPIAdmins.html", "#58a6ff",
    "For senior leadership. Delivers strategic depth via four navigable tabs covering market, regions, segments, and growth."))
story.append(SP(10))
story.append(H2("Purpose"))
story.append(P(
    "Provides the analytical depth that executive committees need: regional benchmarking, "
    "property segment analysis, and multi-year growth dynamics. The tabbed interface lets "
    "each committee member navigate to the view most relevant to their portfolio."
))
story.append(H2("Tab Guide"))
tabs = [
    ("Tab 1 — Market Overview",
     "Aggregate market view. Contains: 5 KPI indicators (avg price, growth, top region, volume, price spread); "
     "interactive price trend by region (click legends to toggle regions); YoY % change bar chart; "
     "annual area chart; total volume trend."),
    ("Tab 2 — Regional Analysis",
     "Six region cards showing: current avg price, YoY badge (green = positive, red = negative), "
     "monthly volume, and property type sub-prices. Below the cards: full historical price chart "
     "and a start-vs-end grouped bar comparison."),
    ("Tab 3 — Property Segments",
     "Grouped bar chart: Detached, Semi-Detached, Terraced, and Flat prices side-by-side per region. "
     f"Horizontal bar shows Detached-Flat premium per region (market average: {fmtp(det_flat_premium)}). "
     f"Annual trend lines show each segment's trajectory {first_year}–{latest_year}."),
    ("Tab 4 — Growth Dynamics",
     "Rolling YoY % change chart for all 6 regions — this is where the 2007–2009 downturn and "
     f"subsequent recovery are most visible. Summary table shows each region's {first_year} start price, "
     f"{latest_year} end price, and total % appreciation ranked by growth."),
]
for title, desc in tabs:
    story.append(H3(title))
    story.append(P(desc))
story.append(H2("How to Use"))
story.append(step_table([
    ("Open the file", "Double-click ExCo_HPIAdmins.html. The Market Overview tab loads by default."),
    ("Navigate tabs", "Click the tab labels at the top (Market Overview / Regional Analysis / Property Segments / Growth Dynamics)."),
    ("Toggle regions (Tab 1)", "In the price trend chart, click a region name in the legend below the chart to show/hide that region."),
    ("Read region cards (Tab 2)", "Each card shows the YoY % badge — green means price rose year-on-year, red means it fell."),
    ("Assess segment gap (Tab 3)", "The horizontal bar chart ranks regions by their Detached-vs-Flat price gap — useful for investment targeting."),
    ("Spot the crisis (Tab 4)", "In the YoY chart, look for the dip around 2008–2009. This is the financial crisis impact — followed by the recovery."),
]))

# ── 5. MANAGERS DASHBOARD ────────────────────────────────────────────────────
story.append(PageBreak())
story.append(aud_header("Managers Dashboard", "Mgr_HPIAdmins.html", "#1d8a6f",
    "For regional and department managers. Filter by your region and time range to monitor local performance."))
story.append(SP(10))
story.append(H2("Purpose"))
story.append(P(
    "Gives managers the ability to zoom into their own region's performance without being "
    "overwhelmed by data from other areas. The filter bar at the top drives all charts simultaneously, "
    "turning a market-wide view into a focused regional briefing in seconds."
))
story.append(H2("Filter Bar Controls"))
filters = [
    ("Region selector", "Drop-down list. Select 'All Regions' for a market-wide view, or pick one region to isolate all charts to that area."),
    ("From year slider", "Drag left/right to set the start of the analysis window. The year value updates live next to the slider."),
    ("To year slider", f"Drag left/right to set the end of the analysis window. Use a narrow range (e.g. 2010–{latest_year}) for recent-period focus."),
]
ft = Table([[Paragraph(f"<b>{n}</b>", sty('Normal', fontSize=9, fontName='Helvetica-Bold')),
             Paragraph(d, sty('Normal', fontSize=9, leading=13))]
            for n,d in filters], colWidths=[4*cm, 12*cm])
ft.setStyle(TableStyle([
    ('BACKGROUND', (0,0),(-1,-1), C_SURFACE2),
    ('GRID', (0,0),(-1,-1), 0.5, C_BORDER2),
    ('VALIGN', (0,0),(-1,-1), 'MIDDLE'),
    ('TOPPADDING', (0,0),(-1,-1), 6), ('BOTTOMPADDING', (0,0),(-1,-1), 6),
    ('LEFTPADDING', (0,0),(-1,-1), 8),
]))
story.append(ft)
story.append(SP(8))
story.append(H2("Dashboard Sections"))
story.append(B("<b>Region Status Cards (top row):</b> One card per region. Shows current avg price, colour-coded YoY badge, and monthly volume. The coloured top bar matches the region's chart colour."))
story.append(B("<b>6 KPI indicators:</b> Market avg price, overall growth, highest/lowest priced region, fastest-growing region, and total monthly volume."))
story.append(B("<b>Average Price Trend:</b> Monthly prices filtered by region and year range. Multiple regions display as coloured lines."))
story.append(B("<b>Monthly Volume chart:</b> Transactions per month — useful for spotting seasonal dips and market activity patterns."))
story.append(B("<b>YoY % Change:</b> 12-month rolling change — shows whether your market is in growth or decline phase."))
story.append(B("<b>MoM % Change:</b> 1-month change — early warning signal for short-term shifts."))
story.append(B("<b>Property Type chart:</b> All four property types per region at the latest month — compare your regional mix."))
story.append(B("<b>Regional Ranking Table:</b> Sortable. Click any column header to rank regions by that metric."))
story.append(SP(8))
story.append(H2("How to Use"))
story.append(step_table([
    ("Open the file", "Double-click Mgr_HPIAdmins.html in a browser."),
    ("Select your region", "Use the Region drop-down at the top. All 4 trend charts and KPIs update immediately."),
    ("Set your time window", "Use From/To year sliders to focus on a specific period (e.g. post-2010 for recent trends)."),
    ("Read region cards", "Find your region's card at the top — note YoY badge colour (green = growth, red = decline)."),
    ("Monitor MoM signals", "The MoM % chart is the earliest indicator. Negative readings for 2+ months signal a softening market."),
    ("Use the ranking table", "Click 'YoY' column to rank regions by year-on-year growth. Click 'Avg Price' to rank by price level."),
]))

# ── 6. OPERATIONS DASHBOARD ──────────────────────────────────────────────────
story.append(PageBreak())
story.append(aud_header("Operations Dashboard", "Ops_HPIAdmins.html", "#0e7490",
    "For analysts and data teams. Full data console with every row, every metric, and real-time filtering."))
story.append(SP(10))
story.append(H2("Purpose"))
story.append(P(
    "The Operations dashboard gives analysts complete transparency into the underlying data. "
    "Every metric from the cleaned dataset is accessible — filterable, searchable, and sortable. "
    "It is designed for data validation, deep-dive analysis, and quality control tasks."
))
story.append(H2("Control Bar"))
ops_controls = [
    ("Region",       "Filter all charts and the data table to a specific region, or keep 'ALL' for full dataset."),
    ("From / To year", "Year range sliders. Narrow the window to focus analysis on a specific period."),
    ("Search box",   "Type a region name, date (YYYY-MM format), or area name. The table filters instantly as you type."),
    ("Rows/page",    "Set how many records appear per table page: 20, 50, or 100."),
    ("Metric selector", "Switch the primary trend chart between: Avg Price, Volume, YoY %, or MoM %."),
]
oc = Table([[Paragraph(f"<b>{n}</b>", sty('Normal', fontSize=9, fontName='Helvetica-Bold')),
             Paragraph(d, sty('Normal', fontSize=9, leading=13))]
            for n,d in ops_controls], colWidths=[3.5*cm, 12.5*cm])
oc.setStyle(TableStyle([
    ('BACKGROUND', (0,0),(0,-1), C_NAVY),
    ('TEXTCOLOR', (0,0),(0,-1), colors.HexColor('#89dceb')),
    ('FONTNAME', (0,0),(0,-1), 'Helvetica-Bold'),
    ('BACKGROUND', (1,0),(1,-1), C_SURFACE2),
    ('GRID', (0,0),(-1,-1), 0.5, C_BORDER2),
    ('VALIGN', (0,0),(-1,-1), 'MIDDLE'),
    ('TOPPADDING', (0,0),(-1,-1), 6), ('BOTTOMPADDING', (0,0),(-1,-1), 6),
    ('LEFTPADDING', (0,0),(-1,-1), 8),
]))
story.append(oc)
story.append(SP(8))
story.append(H2("Status Bar"))
story.append(P(
    "Seven live counters at the top of the page update with every filter change:"
))
story.append(B("<b>Total Rows:</b> Always 1,572 (full dataset minus flagged gap rows)."))
story.append(B("<b>Filtered Rows:</b> Number of records matching the current filter combination."))
story.append(B("<b>Data Gaps:</b> Count of rows with SalesVolume = 0 in the current filter scope."))
story.append(B("<b>Avg Price / Avg YoY / Total Vol:</b> Aggregated from filtered rows only."))
story.append(B("<b>Date Range:</b> Shows the earliest and latest date in the current filter."))
story.append(SP(6))
story.append(H2("Data Table"))
story.append(B("Click any column header to sort ascending. Click again to sort descending."))
story.append(B("The search box filters across region name, area, and date simultaneously."))
story.append(B("Rows where SalesVolume = 0 display a <b>⚠</b> warning icon — these are the 12 data gap rows (Nov–Dec 2016)."))
story.append(B("Use Rows/page = 100 and search a specific region to review its full history on one page."))
story.append(SP(8))
story.append(H2("How to Use"))
story.append(step_table([
    ("Open the file", "Double-click Ops_HPIAdmins.html. All 1,572 valid rows load by default."),
    ("Apply filters", "Select region and year range — the status bar, all 5 charts, and the table all update together."),
    ("Search the table", "Type 'Greater Manchester' or '2008-' to quickly isolate specific records in the table."),
    ("Sort columns", "Click any table header to sort. Useful: sort by YoY % to find the highest/lowest growth months."),
    ("Switch metrics", "Use the Metric selector to switch the primary trend chart from price to volume, YoY %, or MoM %."),
    ("Spot data gaps", "Sort by Volume ascending — the 12 flagged gap rows (Nov–Dec 2016) will surface at the top with ⚠ warnings."),
    ("Paginate", "Use Prev/Next buttons to navigate pages. The page count and row range update at the bottom of the table."),
]))

# ── 7. DATA DICTIONARY ───────────────────────────────────────────────────────
story.append(PageBreak())
story.append(H1("7. Data Dictionary"))
story.append(HR())
story.append(P("All fields available in HPIAdmins_v1.xlsx and reflected in the dashboards:"))
story.append(SP(6))
dict_data = [
    ["Field", "Type", "Description"],
    ["Date",                  "Date",   "First day of the month (e.g. 2005-03-01)"],
    ["RegionName",            "Text",   "Name of the metropolitan region"],
    ["Area",                  "Text",   "Broad area (North, Northwest, Midlands)"],
    ["AreaCode",              "Text",   "ONS geographic code for the region"],
    ["AveragePrice",          "£",      "Mean house price across all property types"],
    ["Index",                 "Float",  "HPI index value (base = 100 at a reference point)"],
    ["1mPctChange",           "%",      "Month-on-month % price change (null for Jan 1995)"],
    ["12mPctChange",          "%",      "Year-on-year % price change (null for first 12 months)"],
    ["SalesVolume",           "Int",    "Number of transactions recorded that month"],
    ["DetachedPrice",         "£",      "Average price for detached properties"],
    ["DetachedIndex",         "Float",  "HPI index for detached properties"],
    ["Detached1mPctChange",   "%",      "MoM change for detached properties"],
    ["Detached12mPctChange",  "%",      "YoY change for detached properties"],
    ["SemiDetachedPrice",     "£",      "Average price for semi-detached properties"],
    ["SemiDetachedIndex",     "Float",  "HPI index for semi-detached properties"],
    ["SemiDetached1mPctChange","%",     "MoM change for semi-detached"],
    ["SemiDetached12mPctChange","%",    "YoY change for semi-detached"],
    ["TerracedPrice",         "£",      "Average price for terraced properties"],
    ["TerracedIndex",         "Float",  "HPI index for terraced properties"],
    ["Terraced1mPctChange",   "%",      "MoM change for terraced"],
    ["Terraced12mPctChange",  "%",      "YoY change for terraced"],
    ["FlatPrice",             "£",      "Average price for flats/apartments"],
    ["FlatIndex",             "Float",  "HPI index for flats"],
    ["Flat1mPctChange",       "%",      "MoM change for flats"],
    ["Flat12mPctChange",      "%",      "YoY change for flats"],
    ["DataGapFlag",           "Bool",   "True = SalesVolume is 0 (data not reported). Affects Nov–Dec 2016 only."],
]
dd = Table(dict_data, colWidths=[5.0*cm, 1.5*cm, 9.5*cm])
dd.setStyle(TableStyle([
    ('BACKGROUND', (0,0),(-1,0), C_NAVY),
    ('TEXTCOLOR', (0,0),(-1,0), colors.white),
    ('FONTNAME', (0,0),(-1,0), 'Helvetica-Bold'),
    ('FONTSIZE', (0,0),(-1,-1), 8),
    ('GRID', (0,0),(-1,-1), 0.5, C_BORDER2),
    ('ROWBACKGROUNDS', (0,1),(-1,-1), [colors.white, C_SURFACE2]),
    ('VALIGN', (0,0),(-1,-1), 'MIDDLE'),
    ('TOPPADDING', (0,0),(-1,-1), 4), ('BOTTOMPADDING', (0,0),(-1,-1), 4),
    ('LEFTPADDING', (0,0),(-1,-1), 6),
]))
story.append(dd)
story.append(SP(12))
story.append(Callout(
    "All dashboards read from pre-computed data embedded at generation time. "
    "To refresh with updated data, re-run the Python generation scripts with the new HPIAdmins_v1.xlsx."
))

# Build
doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
size_kb = os.path.getsize(pdf_path)/1024
print(f"PDF saved: {pdf_path} ({size_kb:.0f} KB)")
print("Done.")
