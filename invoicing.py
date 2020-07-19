import pandas as pd

df = pd.read_excel(r"E:\Nikhil\automation\vms_dump.xlsx")
print (df)
df.weekending = df.weekending + '/2020'
df.weekending = pd.to_datetime(df.weekending, format='%m/%d/%Y')
print (df.set_index('weekending').resample('D').ffill().reset_index())
#print (df)
