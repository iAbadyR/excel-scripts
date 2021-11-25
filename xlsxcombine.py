import os
import pandas as pd
generated_filename = 'combined3.xlsx'
cwd = os.path.abspath('')
files = os.listdir(cwd)
df = pd.DataFrame()
for file in files:
    if file.endswith('.xlsx'):
        filedata = pd.read_excel(file)
        x_row = filedata.columns[1]
        y_row = filedata.columns[2]
        no_row = filedata.columns[0]
        dfs = pd.read_excel('template.xlsx')
        dfs['x'] = filedata[x_row].values
        dfs['y'] = filedata[y_row].values
        dfs['#'] = filedata[no_row].values
        dfs['name'] = file
        df = df.append(dfs, ignore_index=True)
df.head()
if os.path.exists(generated_filename):
  os.remove(generated_filename)
df.to_excel(generated_filename)
