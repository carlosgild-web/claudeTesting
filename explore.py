import pandas as pd
import numpy as np

df = pd.read_excel(r'C:\MyData\ClaudeProjects\DashboardTest\PowerBIData.xlsx', sheet_name='HPIAdmins')

print('=== SHAPE ===')
print('Rows:', df.shape[0], '| Columns:', df.shape[1])

print('\n=== DATE RANGE ===')
print('From:', df['Date'].min(), '| To:', df['Date'].max())
print('Unique dates:', df['Date'].nunique())

print('\n=== UNIQUE VALUES ===')
print('Regions:', df['RegionName'].nunique(), '=>', sorted(df['RegionName'].unique()))
print('Areas:', df['Area'].nunique(), '=>', sorted(df['Area'].unique()))
print('AreaCodes:', df['AreaCode'].nunique())

print('\n=== NULL COUNTS ===')
nulls = df.isnull().sum()
print(nulls[nulls > 0].to_string())

print('\n=== NULL % ===')
null_pct = (df.isnull().sum() / len(df) * 100).round(1)
print(null_pct[null_pct > 0].to_string())

print('\n=== DUPLICATES ===')
print('Full duplicates:', df.duplicated().sum())
print('Date+AreaCode duplicates:', df.duplicated(subset=['Date','AreaCode']).sum())

print('\n=== NUMERIC SUMMARY ===')
print(df[['AveragePrice','SalesVolume','Index']].describe().round(2).to_string())

print('\n=== NEGATIVE VALUES ===')
num_cols = ['AveragePrice','SalesVolume','DetachedPrice','SemiDetachedPrice','TerracedPrice','FlatPrice']
found_neg = False
for col in num_cols:
    neg = (df[col] < 0).sum()
    if neg > 0:
        print(col + ':', neg, 'negative values')
        found_neg = True
if not found_neg:
    print('No negative values found in price/volume columns.')

print('\n=== SALES VOLUME ZEROS ===')
print('SalesVolume == 0:', (df['SalesVolume'] == 0).sum())
print('SalesVolume < 0:', (df['SalesVolume'] < 0).sum())

print('\n=== OUTLIER CHECK (AveragePrice) ===')
q1 = df['AveragePrice'].quantile(0.25)
q3 = df['AveragePrice'].quantile(0.75)
iqr = q3 - q1
low = q1 - 1.5 * iqr
high = q3 + 1.5 * iqr
outliers = df[(df['AveragePrice'] < low) | (df['AveragePrice'] > high)]
print('Price outliers (IQR method):', len(outliers))
if len(outliers) > 0:
    print(outliers[['Date','RegionName','AveragePrice']].to_string())
