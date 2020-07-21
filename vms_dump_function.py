import pandas as pd

pd.set_option('display.max_columns',None)
pd.set_option('display.expand_frame_repr', False)
pd.set_option('display.max_rows', None)

vmsdump_df = pd.read_excel(r"E:\Nikhil\automation\Invoicing_automation\vms_dump.xlsx", header=0, sheet_name='Sheet2')
vmsdump_df = vmsdump_df.set_index('RACF ID')    #Setting index to Racf id to make calculations easier
working_hrs_per_day = 8
print(vmsdump_df)

# ===================================================================================================================
# Reads the user Input and then generates the calender days expected for the current month Invoicing.
# This is needed as some teammates might take leave or submit 0 hours in the VMS and we don't get VMS data for those
# To handle those cases we are first creating a perfect calender and then fill in details from VMS Dump.
# In case any week is not available in the VMS Dump sheet, then that week would become 0.
# ===================================================================================================================

def create_default_calender(racf_id,start_date,end_date):
    date_index=pd.date_range(start=start_date, end=end_date,freq='W')
    vms_generated_calndr_df=pd.DataFrame(date_index, columns=['WeekEnding'])
    vms_generated_calndr_df['RACF ID'] = racf_id
    print (vms_generated_calndr_df)

    return vms_generated_calndr_df


# Reads the VMS Dump and create a pandas DataFrame with needed columns
def generate_vms_sheet(racf_id,vms_generated_calndr_df):    #(leave_tracker_index,racf_id,start_date,end_date):
    # Merge the VMS generated DataFrame which has the correct start and end Date with
    # the input VMS dump DF. This might have more or less weeks as compared to required dates.
    vmsdump_user_df = pd.merge(vms_generated_calndr_df,vmsdump_df.loc[racf_id,['WeekEnding', 'Reg Hours', 'OT Hours']],how='left',on=['WeekEnding'])
    print(vmsdump_user_df)

    # Add columns which would be required for the calculations of the VMS DUMP
    vmsdump_user_df['vms_WeekStarting'] = vmsdump_user_df['WeekEnding'] + pd.offsets.Day(-6)
    vmsdump_user_df[['Reg Hours','OT Hours']] = vmsdump_user_df[['Reg Hours','OT Hours']].fillna(0).astype(int) #Replace NaN with 0
    vmsdump_user_df['vms_pending_hours'] = vmsdump_user_df['Reg Hours'] + vmsdump_user_df['OT Hours']  #Created a new column for keeping VMS hours counter
    vmsdump_user_df['vms_working_days'] = ((vmsdump_user_df['Reg Hours'] + vmsdump_user_df['OT Hours']) / working_hrs_per_day).astype(int)

    # Below code resample the VMS Weekly data into Daily data.
    # As the resample method doesn't expand the last entry till the end, so we have to add another duplicate last row for the same.
    vmsdump_user_df = vmsdump_user_df.append(vmsdump_user_df.iloc[-1])  #appends the last row again
    vmsdump_user_df.iloc[-1, vmsdump_user_df.columns.get_loc('vms_WeekStarting')] = vmsdump_user_df.iloc[-1, vmsdump_user_df.columns.get_loc('WeekEnding')]
    vmsdump_user_df = vmsdump_user_df.set_index('vms_WeekStarting').resample('D').ffill().reset_index().set_index('RACF ID')
    vmsdump_user_df['weekday_flag'] = vmsdump_user_df['vms_WeekStarting'].apply(lambda x: x.date().weekday()<=4 if 1 else 0).astype(int)
    vmsdump_user_df['weekend_flag'] = vmsdump_user_df['vms_WeekStarting'].apply(lambda x: x.date().weekday()>4 if 1 else 0).astype(int)

    # print(vmsdump_user_df[])


    print(vmsdump_user_df)
    # print(vmsdump_user_df.to_csv())

vms_generated_calndr_df=create_default_calender('abcd1','2020-03-30','2020-04-26')
generate_vms_sheet('abcd1',vms_generated_calndr_df)
