"""
Appends one year of synthetic data to PowerBIData.xlsx (HPIAdmins tab).

Logic:
  - Finds the latest year Y in the data.
  - If Y+1 already exists (idempotent), exits cleanly with skipped=true.
  - Copies all 12 months of Y per region into Y+1, applying:
      • A per-region random annual growth rate (3–7 %).
      • Small per-month random noise (±0.5 %).
  - Recomputes all 1m% and 12m% change columns from the new series.
  - Appends and saves back to PowerBIData.xlsx.

Prints a single JSON result to stdout.
"""
import json, os, sys
import numpy as np
import pandas as pd

rng = np.random.default_rng()

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(BASE_DIR, 'PowerBIData.xlsx')
SHEET     = 'HPIAdmins'

# (price_col, 1m_col, 12m_col)
PRICE_TRIPLES = [
    ('AveragePrice',      '1m%Change',             '12m%Change'),
    ('DetachedPrice',     'Detached1m%Change',      'Detached12m%Change'),
    ('SemiDetachedPrice', 'SemiDetached1m%Change',  'SemiDetached12m%Change'),
    ('TerracedPrice',     'Terraced1m%Change',      'Terraced12m%Change'),
    ('FlatPrice',         'Flat1m%Change',          'Flat12m%Change'),
]

# (index_col, matching_price_col) — index scales proportionally with price
INDEX_PAIRS = [
    ('Index',             'AveragePrice'),
    ('DetachedIndex',     'DetachedPrice'),
    ('SemiDetachedIndex', 'SemiDetachedPrice'),
    ('TerracedIndex',     'TerracedPrice'),
    ('FlatIndex',         'FlatPrice'),
]


def pct_chg(new_val, old_val):
    if pd.isna(old_val) or old_val == 0:
        return np.nan
    return round((new_val / old_val - 1) * 100, 6)


def main():
    if not os.path.exists(XLSX_PATH):
        print(json.dumps({'error': 'PowerBIData.xlsx not found'}))
        sys.exit(1)

    df = pd.read_excel(XLSX_PATH, sheet_name=SHEET)
    df['Date'] = pd.to_datetime(df['Date'])

    max_year  = int(df['Date'].dt.year.max())
    next_year = max_year + 1

    if (df['Date'].dt.year == next_year).any():
        print(json.dumps({
            'skipped':   True,
            'reason':    f'Data for {next_year} already present',
            'max_year':  max_year,
            'next_year': next_year,
        }))
        sys.exit(0)

    regions = sorted(df['RegionName'].unique())

    # Region metadata (Area, AreaCode) — from the most-recent row per region
    region_meta = (
        df.sort_values('Date')
          .groupby('RegionName')[['Area', 'AreaCode']]
          .last()
          .to_dict('index')
    )

    # Per-region random annual growth 3–7 %
    region_growth = {r: rng.uniform(0.03, 0.07) for r in regions}

    prev_yr = df[df['Date'].dt.year == max_year].sort_values(['RegionName', 'Date'])

    new_rows = []
    for region in regions:
        base = prev_yr[prev_yr['RegionName'] == region].sort_values('Date')
        g    = region_growth[region]

        for _, src in base.iterrows():
            month_noise = rng.uniform(-0.005, 0.005)
            scale = (1 + g) * (1 + month_noise)
            new_date = src['Date'].replace(year=next_year)

            row = {
                'Date':          new_date,
                'RegionName':    region,
                'Area':          region_meta[region]['Area'],
                'AreaCode':      region_meta[region]['AreaCode'],
                'IndexSA':       np.nan,
                'AveragePriceSA': np.nan,
            }

            for price_col, _, _ in PRICE_TRIPLES:
                v = src[price_col]
                row[price_col] = (v * scale) if not pd.isna(v) else np.nan

            for idx_col, ref_price in INDEX_PAIRS:
                old_idx = src[idx_col]
                old_p   = src[ref_price]
                new_p   = row[ref_price]
                if not pd.isna(old_idx) and not pd.isna(old_p) and old_p != 0:
                    row[idx_col] = old_idx * (new_p / old_p)
                else:
                    row[idx_col] = np.nan

            sv = src['SalesVolume']
            row['SalesVolume'] = (
                max(0, int(sv * rng.uniform(0.85, 1.15)))
                if (not pd.isna(sv) and sv > 0) else 0
            )

            # Placeholder % changes — computed below
            for _, mom_col, yoy_col in PRICE_TRIPLES:
                row[mom_col] = np.nan
                row[yoy_col] = np.nan

            new_rows.append(row)

    new_df = pd.DataFrame(new_rows)[df.columns]

    # Combine original + new to compute % changes across the year boundary
    combined = pd.concat([df, new_df], ignore_index=True)
    combined['Date'] = pd.to_datetime(combined['Date'])
    combined = combined.sort_values(['RegionName', 'Date']).reset_index(drop=True)
    new_mask = combined['Date'].dt.year == next_year

    for price_col, mom_col, yoy_col in PRICE_TRIPLES:
        combined[mom_col] = combined[mom_col].astype(float)
        combined[yoy_col] = combined[yoy_col].astype(float)

        for region in regions:
            r_mask = combined['RegionName'] == region
            r_data = combined[r_mask].sort_values('Date')

            for idx in combined[r_mask & new_mask].index:
                cur_price = combined.at[idx, price_col]
                cur_date  = combined.at[idx, 'Date']

                # 1m: closest earlier date in same region
                prev_rows = r_data[r_data['Date'] < cur_date]
                if not prev_rows.empty:
                    combined.at[idx, mom_col] = pct_chg(cur_price, prev_rows.iloc[-1][price_col])

                # 12m: same month one year prior (from original data only)
                prev_yr_date = cur_date.replace(year=cur_date.year - 1)
                old_row = df[(df['RegionName'] == region) & (df['Date'] == prev_yr_date)]
                if not old_row.empty:
                    combined.at[idx, yoy_col] = pct_chg(cur_price, float(old_row.iloc[0][price_col]))

    final_new = combined[new_mask][df.columns]
    result    = pd.concat([df, final_new], ignore_index=True)
    result    = result.sort_values(['Date', 'RegionName']).reset_index(drop=True)
    result.to_excel(XLSX_PATH, sheet_name=SHEET, index=False)

    print(json.dumps({
        'appended':    True,
        'year_added':  next_year,
        'rows_added':  len(final_new),
        'new_total':   len(result),
        'latest_date': result['Date'].max().strftime('%B %Y'),
    }))


if __name__ == '__main__':
    main()
