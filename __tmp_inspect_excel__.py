import pandas as pd
import os
files=[r'c:\\Users\\MSH\\Desktop\\DAILY REPORT\\Install Plan.xlsx', r'c:\\Users\\MSH\\Desktop\\DAILY REPORT\\Survey Plan.xlsx']
for name in files:
    print('FILE', name, os.path.exists(name))
    if os.path.exists(name):
        df=pd.read_excel(name)
        print('  rows', len(df), 'columns', list(df.columns))
