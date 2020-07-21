import numpy as np
import datetime
import pandas as pd

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
    leave_tracker_df=pd.read_excel(r'E:\Nikhil\automation\Invoicing_automation\Leave_Tracker_Marketing_Finance_2020.xlsx',
    sheet_name='Tracker',skiprows=start_index-2, nrows=(end_index-start_index+1))
    leave_tracker_df=leave_tracker_df.set_index('RACF ID')
    leave_tracker_df=leave_tracker_df.loc[:,'Resource Names ':]
    leave_tracker_df=leave_tracker_df.reset_index()
    del leave_tracker_df['Resource Names ']
    leave_tracker_df=leave_tracker_df.set_index('RACF ID')
    leave_tracker_df=leave_tracker_df.T
    leave_tracker_df.index = pd.to_datetime(leave_tracker_df.index)
    leave_tracker_df=leave_tracker_df.loc[main_range_actual_start:main_range_actual_end]
    leave_tracker_df = leave_tracker_df.fillna(value='working')
    leave_tracker_df = leave_tracker_df.transform(lambda x : x.str.contains(working_day_strings,flags=re.IGNORECASE,regex=True) if 1 else 0).astype(int)
    leave_tracker_df = leave_tracker_df.replace(to_replace=1, value=working_hrs_per_day)
    leave_tracker_df['WeekEnding'] = leave_tracker_df.index.week
    leave_tracker_df = leave_tracker_df.reset_index()

    leave_tracker_df=leave_tracker_df.set_index('index')
    leave_tracker_df = leave_tracker_df.groupby('WeekEnding').transform(lambda x: x.replace(working_hrs_per_day,x.sum()) )
    return leave_tracker_df

print(read_leave_tracker(start_index,end_index,main_range_actual_start,main_range_actual_end))
# print(leave_tracker_leave_tracker_df)
