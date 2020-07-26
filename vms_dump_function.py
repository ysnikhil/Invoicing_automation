import numpy as np
import datetime
import pandas as pd
import re

pd.set_option('display.max_columns',None)
pd.set_option('display.expand_frame_repr', False)
pd.set_option('display.max_rows', None)

start_index = 18
end_index = 37
main_range_actual_start = datetime.datetime(2020, 3, 30, 0, 0)
main_range_actual_end = datetime.datetime(2020, 4, 26, 0, 0)
working_hrs_per_day = 8
working_day_strings = 'working|w-b'

def read_leave_tracker(start_index,end_index,main_range_actual_start,main_range_actual_end):
    # Taking the data from the main code and reading only the required data from the leave tracker.
    leave_tracker_df=pd.read_excel(r'E:\Nikhil\automation\Invoicing_automation\Leave_Tracker_Marketing_Finance_2020.xlsx',
    sheet_name='Tracker',skiprows=start_index-2, nrows=(end_index-start_index+1))
    leave_tracker_df=leave_tracker_df.set_index('RACF ID')
    leave_tracker_df=leave_tracker_df.loc[:,'Resource Names ':] #The space in the name is due to the space in the leave t
    leave_tracker_df=leave_tracker_df.reset_index()
    del leave_tracker_df['Resource Names '] #Removing the resource name as it is not needed. If required, we can put this back.
    leave_tracker_df=leave_tracker_df.set_index('RACF ID')
    leave_tracker_df=leave_tracker_df.T #Transposing data, as we need the calendar days as columns for easier calculations.
    leave_tracker_df.index = pd.to_datetime(leave_tracker_df.index) #Convert the date value from leave tracker to datetime for accessing all the ready made datetime functions in pandas.
    leave_tracker_df=leave_tracker_df.loc[main_range_actual_start:main_range_actual_end]    #Take only the required data from the leave tracker.
    leave_tracker_df['WeekEnding'] = leave_tracker_df.index.week
    print (leave_tracker_df)
    return leave_tracker_df

print(read_leave_tracker(start_index,end_index,main_range_actual_start,main_range_actual_end))

vmsdump_df = pd.read_excel(r"E:\Nikhil\automation\Invoicing_automation\vms_dump.xlsx", header=0, sheet_name='Sheet2')
vmsdump_df = vmsdump_df.set_index('RACF ID')    #Setting index to Racf id to make calculations easier
# print(vmsdump_df)

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
    # print (vms_generated_calndr_df)

    return vms_generated_calndr_df


# Reads the VMS Dump and create a pandas DataFrame with needed columns
def generate_vms_sheet(racf_id,vms_generated_calndr_df):    #(leave_tracker_index,racf_id,start_date,end_date):
    # Merge the VMS generated DataFrame which has the correct start and end Date with
    # the input VMS dump DF. This might have more or less weeks as compared to required dates.
    vmsdump_user_df = pd.merge(vms_generated_calndr_df,vmsdump_df.loc[racf_id,['WeekEnding', 'Reg Hours', 'OT Hours']],how='left',on=['WeekEnding'])
    # print(vmsdump_user_df)

    # Add columns which would be required for the calculations of the VMS DUMP
    vmsdump_user_df['vms_WeekStarting'] = vmsdump_user_df['WeekEnding'] + pd.offsets.Day(-6)
    vmsdump_user_df[['Reg Hours','OT Hours']] = vmsdump_user_df[['Reg Hours','OT Hours']].fillna(0).astype(int) #Replace NaN with 0
    vmsdump_user_df['vms_pending_hours'] = vmsdump_user_df['Reg Hours'] + vmsdump_user_df['OT Hours']  #Created a new column for keeping VMS hours counter
    vmsdump_user_df['vms_working_days'] = ((vmsdump_user_df['Reg Hours'] + vmsdump_user_df['OT Hours']) / working_hrs_per_day).astype(int)

    # Below code resample the VMS Weekly data into Daily data.
    # As the resample method doesn't expand the last entry till the end, so we have to add another duplicate last row for the same.
    vmsdump_user_df = vmsdump_user_df.append(vmsdump_user_df.iloc[-1])  #appends the last row again
    vmsdump_user_df.iloc[-1, vmsdump_user_df.columns.get_loc('vms_WeekStarting')] = vmsdump_user_df.iloc[-1, vmsdump_user_df.columns.get_loc('WeekEnding')]

    vmsdump_user_df = vmsdump_user_df.set_index('vms_WeekStarting').resample('D').ffill().reset_index()
    vmsdump_user_df['weekday_flag'] = vmsdump_user_df['vms_WeekStarting'].apply(lambda x: x.date().weekday()<=4 if 1 else 0).astype(int)
    vmsdump_user_df['weekend_flag'] = vmsdump_user_df['vms_WeekStarting'].apply(lambda x: x.date().weekday()>4 if 1 else 0).astype(int)

    leave_tracker_df = read_leave_tracker(start_index,end_index,main_range_actual_start,main_range_actual_end)
    leave_tracker_df = leave_tracker_df.loc[:,[racf_id,'WeekEnding']]
    leave_tracker_df[racf_id].fillna(value='working',inplace=True) #Converting all the blank values as working, as if anyday is blank then it is assumed that it is a normal working day.
    leave_tracker_df[racf_id] = leave_tracker_df[racf_id].transform(lambda x : x.str.contains(working_day_strings,flags=re.IGNORECASE,regex=True)).astype(float)  #Used the pattern match to count the working|w-b as working days, rest all the comments are assumed as 0 working hours.
    leave_tracker_df[racf_id] = leave_tracker_df[racf_id].replace(to_replace=1, value=working_hrs_per_day)    #Replace the boolean value True/1 as the normal working hours as others are converted as 0 in the previous step.
    leave_tracker_df[racf_id] = leave_tracker_df.groupby('WeekEnding')[racf_id].transform(lambda x: x.replace(working_hrs_per_day,x.sum()) )  #Sums up the working hours for each day and replaces it with the sum of the values at per week level.
    print (leave_tracker_df)

    def calc_leave_working_days(df,racf_id):
        df['leave_working_wkdays'] = df[df['weekday_flag']==1][racf_id].gt(0).sum()
        return df

    def calc_leave_weekend_days(df,racf_id):
        df['leave_working_wkenddays'] = df[df['weekend_flag']==1][racf_id].gt(0).sum()
        return df

    vmsdump_leave_merged_df = pd.merge(left=vmsdump_user_df,right=leave_tracker_df.loc[:,racf_id], how='left', left_on=['vms_WeekStarting'], right_on=leave_tracker_df.index)
    vmsdump_leave_merged_df = vmsdump_leave_merged_df.groupby('WeekEnding').apply(calc_leave_working_days, racf_id)
    vmsdump_leave_merged_df = vmsdump_leave_merged_df.groupby('WeekEnding').apply(calc_leave_weekend_days, racf_id)
    vmsdump_leave_merged_df['leave_working_days'] = vmsdump_leave_merged_df.groupby('WeekEnding')[racf_id].transform(lambda x: x.gt(0).sum())
    vmsdump_leave_merged_df['vms_leave_diff'] = vmsdump_leave_merged_df['vms_working_days']-vmsdump_leave_merged_df['leave_working_days']
    # print(vmsdump_user_df)
    print(vmsdump_leave_merged_df)

vms_generated_calndr_df=create_default_calender('a131','2020-03-30','2020-04-26')
vmsdump_leave_merged_df = generate_vms_sheet('a131',vms_generated_calndr_df)
