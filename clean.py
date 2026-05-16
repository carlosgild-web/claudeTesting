import pandas as pd
import numpy as np

df = pd.read_excel(r'C:\MyData\ClaudeProjects\DashboardTest\PowerBIData.xlsx', sheet_name='HPIAdmins')

print('Original shape:', df.shape)

# 1. Drop 100% null columns
df = df.drop(columns=['IndexSA', 'AveragePriceSA'])
print('After dropping null columns:', df.shape)

# 2. Ensure Date is date only (no time)
df['Date'] = pd.to_datetime(df['Date']).dt.normalize()

# 3. Flag zero SalesVolume rows (last 2 months: Nov-Dec 2016)
df['DataGapFlag'] = df['SalesVolume'] == 0
flagged = df['DataGapFlag'].sum()
print(f'Flagged {flagged} rows with SalesVolume=0 (DataGapFlag=True)')

# 4. Round price and index columns to 2 decimals for cleanliness
price_cols = ['AveragePrice','DetachedPrice','DetachedIndex',
              'SemiDetachedPrice','SemiDetachedIndex',
              'TerracedPrice','TerracedIndex',
              'FlatPrice','FlatIndex','Index']
df[price_cols] = df[price_cols].round(2)

# 5. Rename columns: remove special chars for clarity
rename = {
    '1m%Change': '1mPctChange',
    '12m%Change': '12mPctChange',
    'Detached1m%Change': 'Detached1mPctChange',
    'Detached12m%Change': 'Detached12mPctChange',
    'SemiDetached1m%Change': 'SemiDetached1mPctChange',
    'SemiDetached12m%Change': 'SemiDetached12mPctChange',
    'Terraced1m%Change': 'Terraced1mPctChange',
    'Terraced12m%Change': 'Terraced12mPctChange',
    'Flat1m%Change': 'Flat1mPctChange',
    'Flat12m%Change': 'Flat12mPctChange',
}
df = df.rename(columns=rename)

# 6. Sort by Date then RegionName
df = df.sort_values(['Date', 'RegionName']).reset_index(drop=True)

print('Final shape:', df.shape)
print('Columns:', list(df.columns))
print()
print('Sample cleaned data:')
print(df.head(6).to_string())

# Save
out_path = r'C:\MyData\ClaudeProjects\DashboardTest\HPIAdmins_v1.xlsx'
df.to_excel(out_path, index=False, sheet_name='HPIAdmins')
print()
print('Saved to:', out_path)
