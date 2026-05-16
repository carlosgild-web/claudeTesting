import pandas as pd
import numpy as np

df = pd.read_excel(r'C:\MyData\ClaudeProjects\DashboardTest\PowerBIData.xlsx', sheet_name='HPIAdmins')

print('=== VALIDATION REPORT ===')
print()

issues = []

# 1. Fully null columns
fully_null = [c for c in df.columns if df[c].isnull().all()]
print('ISSUE 1 - Columns 100% null (will DROP):')
for c in fully_null:
    print('  -', c)
    issues.append(f'DROP column "{c}" (100% null)')

# 2. Partial nulls
print()
print('ISSUE 2 - Partial nulls (expected for first months - KEEP):')
partial_null = {c: df[c].isnull().sum() for c in df.columns if 0 < df[c].isnull().sum() < len(df)}
for col, cnt in partial_null.items():
    print(f'  {col}: {cnt} nulls ({cnt/len(df)*100:.1f}%) - structural, first-period gaps')

# 3. Zero SalesVolume
zero_vol = df[df['SalesVolume'] == 0]
print()
print('ISSUE 3 - Zero SalesVolume records:', len(zero_vol))
print(zero_vol[['Date','RegionName','Area','SalesVolume']].to_string())
issues.append(f'SalesVolume=0 in {len(zero_vol)} rows - flag as data gaps')

# 4. Date time component
has_time = any(d.time().hour != 0 or d.time().minute != 0 for d in pd.to_datetime(df['Date']))
print()
print('ISSUE 4 - Date has time component:', has_time)
if has_time:
    issues.append('Strip time from Date column')

# 5. Consistency check: Area vs RegionName
region_area_map = df.groupby('RegionName')['Area'].nunique()
inconsistent = region_area_map[region_area_map > 1]
print()
print('ISSUE 5 - Regions mapped to multiple Areas:', len(inconsistent))
if len(inconsistent) == 0:
    print('  OK - All regions map to exactly one area')

# 6. AreaCode consistency
region_code_map = df.groupby('RegionName')['AreaCode'].nunique()
bad_codes = region_code_map[region_code_map > 1]
print()
print('ISSUE 6 - Regions with multiple AreaCodes:', len(bad_codes))
if len(bad_codes) == 0:
    print('  OK - All regions have one consistent AreaCode')

# 7. Price sanity
print()
print('ISSUE 7 - Price range sanity:')
print('  AveragePrice min:', round(df['AveragePrice'].min(), 2), '| max:', round(df['AveragePrice'].max(), 2))
print('  DetachedPrice min:', round(df['DetachedPrice'].min(), 2), '| max:', round(df['DetachedPrice'].max(), 2))
print('  All prices positive:', (df[['AveragePrice','DetachedPrice','SemiDetachedPrice','TerracedPrice','FlatPrice']] > 0).all().all())

print()
print('=== SUMMARY: Issues to fix ===')
for i, iss in enumerate(issues, 1):
    print(f'  {i}. {iss}')
