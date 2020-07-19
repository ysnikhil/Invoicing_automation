import pandas as pd
pd.set_option('display.max_columns',None)
pd.set_option('display.expand_frame_repr', False)
df = pd.read_excel(r"E:\Nikhil\automation\vms_intermediate_sheet.xlsx", header=0)
#df.loc[len(df)]=0
df = df.set_index('Calculation')
df = df.append(pd.Series(name='final_output'))
df = df.T
df = df.rename_axis('Days')
df = df.rename_axis('',axis=1)
df['vms_pending_hours'] = df['vms_hours']
df.loc[ (df['vms_working_days'] > df['leave_working_days']) & (df['weekday_flag'] == 1) & (df['leave_hours'] > 0), ['final_output'] ] = df['leave_hours'] / df['leave_working_days']
df['vms_pending_hours'] = df['vms_pending_hours'] - df.groupby('vms_week').final_output.transform('sum')
# df.loc[]
# print(df.groupby('vms week').final_output.transform('sum'))
print (df)
