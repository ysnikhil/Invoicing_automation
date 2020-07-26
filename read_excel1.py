import numpy as np
import datetime
import pandas as pd
import re

pd.set_option('display.max_columns', None)
pd.set_option('display.expand_frame_repr', False)
pd.set_option('display.max_rows', None)

start_index = 18
end_index = 37
main_range_actual_start = datetime.datetime(2020, 6, 1, 0, 0)
main_range_actual_end = datetime.datetime(2020, 6, 28, 0, 0)
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

racf_id='a131'

leave_tracker_df = read_leave_tracker(start_index,end_index,main_range_actual_start,main_range_actual_end)
leave_tracker_df = leave_tracker_df.loc[:,[racf_id,'WeekEnding']]
leave_tracker_df[racf_id].fillna(value='working',inplace=True) #Converting all the blank values as working, as if anyday is blank then it is assumed that it is a normal working day.
leave_tracker_df[racf_id] = leave_tracker_df[racf_id].transform(lambda x : x.str.contains(working_day_strings,flags=re.IGNORECASE,regex=True)).astype(float)  #Used the pattern match to count the working|w-b as working days, rest all the comments are assumed as 0 working hours.
leave_tracker_df[racf_id] = leave_tracker_df[racf_id].replace(to_replace=1, value=working_hrs_per_day)    #Replace the boolean value True/1 as the normal working hours as others are converted as 0 in the previous step.
leave_tracker_df[racf_id] = leave_tracker_df.groupby('WeekEnding')[racf_id].transform(lambda x: x.replace(working_hrs_per_day,x.sum()) )  #Sums up the working hours for each day and replaces it with the sum of the values at per week level.
print (leave_tracker_df)
