"""
Reads HPIAdmins_v1.xlsx and outputs a JSON highlights dict to stdout.
Used by the refresh-hpi-dashboards skill to populate the email draft.
"""
import json, os, sys
import pandas as pd
import numpy as np

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(BASE_DIR, 'HPIAdmins_v1.xlsx')


def main():
    if not os.path.exists(XLSX_PATH):
        print(json.dumps({'error': 'HPIAdmins_v1.xlsx not found — run clean.py first'}))
        sys.exit(1)

    df = pd.read_excel(XLSX_PATH)
    df['Date'] = pd.to_datetime(df['Date'])
    dv = df[df['DataGapFlag'] == False].copy()

    latest_dt  = dv['Date'].max()
    prev_dt    = dv[dv['Date'] < latest_dt]['Date'].max()
    first_dt   = dv['Date'].min()

    latest = dv[dv['Date'] == latest_dt]
    prev   = dv[dv['Date'] == prev_dt]
    first  = dv[dv['Date'] == first_dt]

    avg_price      = round(float(latest['AveragePrice'].mean()), 0)
    avg_price_prev = round(float(prev['AveragePrice'].mean()), 0)
    price_mom_pct  = round((avg_price / avg_price_prev - 1) * 100, 2)

    first_avg      = round(float(first['AveragePrice'].mean()), 0)
    total_growth   = round((avg_price / first_avg - 1) * 100, 1)

    total_vol      = int(latest['SalesVolume'].sum())
    price_spread   = round(float(latest['AveragePrice'].max() - latest['AveragePrice'].min()), 0)

    # Top YoY region (latest month)
    yoy_vals = latest[['RegionName', '12mPctChange']].dropna()
    if not yoy_vals.empty:
        top_row        = yoy_vals.loc[yoy_vals['12mPctChange'].idxmax()]
        top_yoy_region = top_row['RegionName']
        top_yoy_val    = round(float(top_row['12mPctChange']), 1)
    else:
        top_yoy_region, top_yoy_val = 'N/A', 0.0

    # Fastest 21-year growth region
    growth_list = []
    for r in dv['RegionName'].unique():
        p0 = float(first[first['RegionName'] == r]['AveragePrice'].values[0])
        p1 = float(latest[latest['RegionName'] == r]['AveragePrice'].values[0])
        growth_list.append({'region': r, 'growth': round((p1 / p0 - 1) * 100, 1),
                             'start': round(p0, 0), 'end': round(p1, 0)})
    growth_list.sort(key=lambda x: x['growth'], reverse=True)
    fastest = growth_list[0]

    # Highest / lowest priced region latest
    price_by_region = latest.set_index('RegionName')['AveragePrice'].round(0).to_dict()
    highest_region  = max(price_by_region, key=price_by_region.get)
    lowest_region   = min(price_by_region, key=price_by_region.get)

    output = {
        'latest_date':       latest_dt.strftime('%B %Y'),
        'first_date':        first_dt.strftime('%B %Y'),
        'avg_price':         avg_price,
        'avg_price_prev':    avg_price_prev,
        'price_mom_pct':     price_mom_pct,
        'total_growth_pct':  total_growth,
        'total_volume':      total_vol,
        'price_spread':      price_spread,
        'top_yoy_region':    top_yoy_region,
        'top_yoy_val':       top_yoy_val,
        'fastest_21yr':      fastest,
        'highest_region':    highest_region,
        'highest_price':     price_by_region[highest_region],
        'lowest_region':     lowest_region,
        'lowest_price':      price_by_region[lowest_region],
        'price_by_region':   {k: float(v) for k, v in price_by_region.items()},
        'all_growth':        growth_list,
    }

    print(json.dumps(output, indent=2))


if __name__ == '__main__':
    main()
