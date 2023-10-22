import pandas as pd
import glob as g

filepaths = g.glob('invoices/*.xlsx')


for filepath in filepaths:
    df = pd.read_excel(filepath ,sheet_name= 'Sheet 1')
    print(df)