import pandas as pd
import numpy as np
pd.set_option('display.max_columns',None)
pd.set_option('display.expand_frame_repr', False)
df = pd.read_excel(r"E:\Nikhil\automation\Invoicing_automation\vms_intermediate_sheet.xlsx", header=0)

df = df.set_index('Calculation')
df = df.append(pd.Series(name='final_output'))
df = df.T
df = df.rename_axis('Days')
df = df.rename_axis('',axis=1)
a=df['leave_hours']
b=df['leave_working_days']

# This is to handle VMS > Leave Tracker days
# First set handles the weekdays and Second set handles the weekend.
df.loc[ (df['vms_working_days'] > df['leave_working_days']) & (df['weekday_flag'] == 1) & (df['leave_hours'] > 0) & (df['leave_working_wkdays'] > 0), 'final_output' ] = np.divide(a, b,out=np.zeros_like(a), where=b!=0)
df['vms_pending_hours'] = df['vms_hours'] - df.groupby('vms_week').final_output.transform('sum')

df.loc[ (df['vms_working_days'] > df['leave_working_days']) & (df['weekend_flag'] == 1) & (df['leave_hours'] > 0) & (df['leave_working_wkenddays'] > 0) & (df['vms_pending_hours'] > 0), 'final_output' ] = np.divide(a, b,out=np.zeros_like(a), where=b!=0)
df['vms_pending_hours'] = df['vms_hours'] - df.groupby('vms_week').final_output.transform('sum')

# This is to handle if VMS days = Leave Tracker days
# First set handles the weekdays and Second set handles the weekend.
df.loc[(df['vms_working_days'] == df['leave_working_days']) &  (df['weekday_flag'] == 1) & (df['leave_hours'] > 0) & (df['leave_working_wkdays'] > 0), 'final_output'] = np.divide(a, b, out=np.zeros_like(a), where=b!=0)
df['vms_pending_hours'] = df['vms_hours'] - df.groupby('vms_week').final_output.transform('sum')

df.loc[(df['vms_working_days'] == df['leave_working_days']) &  (df['weekend_flag'] == 1) & (df['leave_hours'] > 0) & (df['leave_working_wkenddays'] > 0) & (df['vms_pending_hours'] > 0), 'final_output'] = np.divide(a, b, out=np.zeros_like(a), where=b!=0)
df['vms_pending_hours'] = df['vms_hours'] - df.groupby('vms_week').final_output.transform('sum')

# This is to handle VMS < Leave Tracker days
# First set handles the weekdays and Second set handles the weekend.
# Here we are assuring that the leave dates are mapped correctly and rest will be autofilled and highlighted.
df.loc[ (df['vms_working_days'] < df['leave_working_days']) & (df['weekday_flag'] == 1) & (df['leave_hours'] == 0) & (df['leave_working_wkdays'] > 0), 'final_output' ] = np.divide(a, b,out=np.zeros_like(a), where=b!=0)
df['vms_pending_hours'] = df['vms_hours'] - df.groupby('vms_week').final_output.transform('sum')

df.loc[ (df['vms_working_days'] < df['leave_working_days']) & (df['weekend_flag'] == 1) & (df['leave_hours'] == 0) & (df['leave_working_wkenddays'] == 0), 'final_output' ] = 0
df['vms_pending_hours'] = df['vms_hours'] - df.groupby('vms_week').final_output.transform('sum')

# To make the final_output equal to 0, if the vms_pending_hours is Zero
df.loc[(df['vms_pending_hours'] == 0) & (df['final_output'].isna()), 'final_output'] = 0


# Below logic is to fill the values for those dates where we don't have the clarity on the VMS hours
# The function takes the groups on VMS WEEK and then distributes the vms_pending_hours.
# If the pending hours are reduced to 0 and still days are left, then those days will be get 0 hours.
def fill_missing(x):
    hours_counter = x['vms_pending_hours'].mean()
    working_hrs_per_day = 8
    # To check if VMS itself is filled as 0 hours
    if hours_counter == 0:
        x['highlight_flag'] = 1
    else:
        x['highlight_flag'] = 0

    # Find the NaN values and assign VMS hours from top to bottom. Assign 0, if the vms_pending_hours reaches 0
    for var in x.index:
        if ( pd.isna(x.loc[var,'final_output']) & (hours_counter != 0) ):
            x.loc[var,'final_output'] = working_hrs_per_day
            x.loc[var,'highlight_flag'] = 1
            hours_counter = hours_counter - working_hrs_per_day
        elif ( pd.isna(x.loc[var,'final_output']) & (hours_counter == 0) ):
            x.loc[var,'final_output'] = 0
            x.loc[var,'highlight_flag'] = 1
    return x


df=df.groupby('vms_week').apply(fill_missing)

# Below will make the vms_pending_hours equal to 0. Note running it, as it gives how much hours were not
# calculated correctly
# df['vms_pending_hours'] = df['vms_hours'] - df.groupby('vms_week').final_output.transform('sum')
print (df[df['vms_week'] == '4/5'])
# print(df)
