import pandas as pd
import glob
files = glob.glob("*.xls*")
main = pd.DataFrame()
for file in files:
    df = pd.concat(pd.read_excel(file, sheet_name=None),ignore_index=True,sort=False)
    main = main.append(df,ignore_index=True)
finale = main.set_index('control_id')
for i in range(1 , 100):
    i = 1;
    finale.to_excel(i&'.xlsx')
    i = i+1