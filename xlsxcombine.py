import os
import pandas as pd
cwd = os.path.abspath('')
files = os.listdir(cwd)
df = pd.DataFrame()
for file in files:
    if file.endswith('.xlsx'):
        dfs = pd.read_excel(file)
        dfs['name'] = file
        df = df.append(dfs, ignore_index=True)
df.head()
df.to_excel('combined2.xlsx')
