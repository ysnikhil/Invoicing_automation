import numpy as np
import datetime
import pandas as pd
import re

pd.set_option('display.max_columns',None)
pd.set_option('display.expand_frame_repr', False)
pd.set_option('display.max_rows', None)

#Define Global Variables
working_hrs_per_day = 8
working_day_strings = 'working|w-b'
# start_index = 18
# end_index = 37
# main_range_actual_start = datetime.datetime(2020, 3, 30, 0, 0)
# main_range_actual_end = datetime.datetime(2020, 4, 26, 0, 0)
# racf_id='a127'
# resource_name='Aparna Mohandas'

# ===================================================================================================================
# Read the entire VMS dump
# ===================================================================================================================
def read_vms_dump():
    # Reads the VMS Dump. Should be modified as per the user input values.
    vmsdump_df = pd.read_excel(r"E:\Nikhil\automation\Invoicing_automation\vms_dump.xlsx", header=0, sheet_name='Sheet2')
    vmsdump_df = vmsdump_df.set_index('RACF ID')    #Setting index to Racf id to make calculations easier

    return vmsdump_df
# ===================================================================================================================
# This function reads the Leave Tracker and creates a Dataframe which contains the data from the
# given start and end date range (main range). It contains data for all the users and using the function
# variables it can be pulled as 1 user leave details at a time.
# ===================================================================================================================
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
    # print (leave_tracker_df)
    return leave_tracker_df

# ===================================================================================================================
# Reads the user Input and then generates the calender days expected for the current month Invoicing.
# This is needed as some teammates might take leave or submit 0 hours in the VMS and we don't get VMS data for those
# To handle those cases we are first creating a perfect calender and then fill in details from VMS Dump.
# In case any week is not available in the VMS Dump sheet, then that week would become 0.
# ===================================================================================================================
def create_default_calender(racf_id,main_range_actual_start,main_range_actual_end):
    date_index=pd.date_range(start=main_range_actual_start, end=main_range_actual_end,freq='W')
    vms_generated_calndr_df=pd.DataFrame(date_index, columns=['WeekEnding'])
    vms_generated_calndr_df['RACF ID'] = racf_id
    # print (vms_generated_calndr_df)
    return vms_generated_calndr_df

# ===============================================================================================================================
# Merges the Default calender with the VMS dump data and create a pandas DataFrame with needed variable columns for calculations.
# This functions uses the read_leave_tracker function and merges the VMS DUMP df with it to bring all the values in
# a single dataframe.
# ===============================================================================================================================
def create_vms_sheet(racf_id,vms_generated_calndr_df,vmsdump_df,leave_tracker_df):    #(racf_id,vms_generated_calndr_df,VMS_dump_sheet)
    # Merge the VMS generated DataFrame which has the correct start and end Date with
    # the input VMS dump DF. This might have more or less weeks as compared to required dates.
    vmsdump_user_df = pd.merge(vms_generated_calndr_df,vmsdump_df.loc[racf_id,['WeekEnding', 'Reg Hours', 'OT Hours']],how='left',on=['WeekEnding'])

    # Add columns which would be required for the calculations of the VMS DUMP
    vmsdump_user_df['vms_WeekStarting'] = vmsdump_user_df['WeekEnding'] + pd.offsets.Day(-6)
    vmsdump_user_df[['Reg Hours','OT Hours']] = vmsdump_user_df[['Reg Hours','OT Hours']].fillna(0).astype(int) #Replace NaN with 0
    vmsdump_user_df['vms_hours'] = vmsdump_user_df['Reg Hours'] + vmsdump_user_df['OT Hours']  #Created a new column for keeping the sum of VMS hours for easier calculations.
    vmsdump_user_df['vms_pending_hours'] = vmsdump_user_df['vms_hours']  #Created a new column for keeping VMS hours counter
    vmsdump_user_df['vms_working_days'] = (vmsdump_user_df['vms_hours'] / working_hrs_per_day).astype(float)

    # Below code resample the VMS Weekly data into Daily data.
    # As the resample method doesn't expand the last entry till the end, so we have to add another duplicate
    # last row for the same.
    vmsdump_user_df = vmsdump_user_df.append(vmsdump_user_df.iloc[-1])  #appends the last row again
    vmsdump_user_df.iloc[-1, vmsdump_user_df.columns.get_loc('vms_WeekStarting')] = vmsdump_user_df.iloc[-1, vmsdump_user_df.columns.get_loc('WeekEnding')]
    vmsdump_user_df = vmsdump_user_df.set_index('vms_WeekStarting').resample('D').ffill().reset_index()
    vmsdump_user_df['weekday_flag'] = vmsdump_user_df['vms_WeekStarting'].apply(lambda x: x.date().weekday()<=4 if 1 else 0).astype(float)
    vmsdump_user_df['weekend_flag'] = vmsdump_user_df['vms_WeekStarting'].apply(lambda x: x.date().weekday()>4 if 1 else 0).astype(float)

    # Call the read_leave_tracker fuction to pull the leave tracker dataframe and then compute working hours
    # leave_tracker_df = read_leave_tracker(start_index,end_index,main_range_actual_start,main_range_actual_end)
    leave_tracker_df = leave_tracker_df[[racf_id,'WeekEnding']]
    leave_tracker_df[racf_id].fillna(value='working',inplace=True) #Converting all the blank values as working, as if anyday is blank then it is assumed that it is a normal working day.
    leave_tracker_df[racf_id] = leave_tracker_df[racf_id].transform(lambda x : x.str.contains(working_day_strings,flags=re.IGNORECASE,regex=True)).astype(float)  #Used the pattern match to count the working|w-b as working days, rest all the comments are assumed as 0 working hours.
    leave_tracker_df[racf_id] = leave_tracker_df[racf_id].replace(to_replace=1, value=working_hrs_per_day)    #Replace the boolean value True/1 as the normal working hours as others are converted as 0 in the previous step.
    leave_tracker_df[racf_id] = leave_tracker_df.groupby('WeekEnding')[racf_id].transform(lambda x: x.replace(working_hrs_per_day,x.sum()) )  #Sums up the working hours for each day and replaces it with the sum of the values at per week level.
    # print (leave_tracker_df)

    # Functions to calculate the leave working weekdays and weekends from the leave tracker df.
    def calc_leave_working_days(df,racf_id):
        df['leave_working_wkdays'] = df[df['weekday_flag']==1][racf_id].gt(0).sum()
        return df

    def calc_leave_weekend_days(df,racf_id):
        df['leave_working_wkenddays'] = df[df['weekend_flag']==1][racf_id].gt(0).sum()
        return df

    # Merges the VMS DUMP DF with the LEAVE TRACKER DF along with the last set of required variable columns.
    vmsdump_leave_merged_df = pd.merge(left=vmsdump_user_df,right=leave_tracker_df.loc[:,racf_id], how='left', left_on=['vms_WeekStarting'], right_on=leave_tracker_df.index)

    vmsdump_leave_merged_df = vmsdump_leave_merged_df.groupby('WeekEnding').apply(calc_leave_working_days, racf_id)
    vmsdump_leave_merged_df = vmsdump_leave_merged_df.groupby('WeekEnding').apply(calc_leave_weekend_days, racf_id)
    vmsdump_leave_merged_df['leave_working_days'] = vmsdump_leave_merged_df.groupby('WeekEnding')[racf_id].transform(lambda x: x.gt(0).sum())
    vmsdump_leave_merged_df['vms_leave_diff'] = vmsdump_leave_merged_df['vms_working_days']-vmsdump_leave_merged_df['leave_working_days']
    # print(vmsdump_user_df)
    vmsdump_leave_merged_df=vmsdump_leave_merged_df.rename(columns={racf_id:'leave_hours'})
    return vmsdump_leave_merged_df

# ===================================================================================================================
# This is the final function which would be called from the main function and this will inturn call all the
# intermediate functions. Runs for each user at a time.
# Input - start_index,end_index,main_range_actual_start,main_range_actual_end,racf_id,working_hrs_per_day,VMS_dump_sheet
# Output - VMS Output for each user.
# ===================================================================================================================
def generate_vms_sheet(racf_id,vms_generated_calndr_df,vmsdump_leave_merged_df,leave_tracker_df,resource_name,appended_final_df_for_styling):
    # print(vmsdump_leave_merged_df)

    a=vmsdump_leave_merged_df['leave_hours']
    b=vmsdump_leave_merged_df['leave_working_days']

    # This is to handle VMS > Leave Tracker days
    # First set handles the weekdays and Second set handles the weekend.
    if (vmsdump_leave_merged_df['vms_working_days'] > vmsdump_leave_merged_df['leave_working_days']).sum():
        vmsdump_leave_merged_df.loc[ (vmsdump_leave_merged_df['vms_working_days'] > vmsdump_leave_merged_df['leave_working_days']) & (vmsdump_leave_merged_df['weekday_flag'] == 1) & (vmsdump_leave_merged_df['leave_hours'] > 0) & (vmsdump_leave_merged_df['leave_working_wkdays'] > 0), 'final_output' ] = np.divide(a, b,out=np.zeros_like(a), where=b!=0)
        vmsdump_leave_merged_df['vms_pending_hours'] = vmsdump_leave_merged_df['vms_hours'] - vmsdump_leave_merged_df.groupby('WeekEnding').final_output.transform('sum')

        vmsdump_leave_merged_df.loc[ (vmsdump_leave_merged_df['vms_working_days'] > vmsdump_leave_merged_df['leave_working_days']) & (vmsdump_leave_merged_df['weekend_flag'] == 1) & (vmsdump_leave_merged_df['leave_hours'] > 0) & (vmsdump_leave_merged_df['leave_working_wkenddays'] > 0) & (vmsdump_leave_merged_df['vms_pending_hours'] > 0), 'final_output' ] = np.divide(a, b,out=np.zeros_like(a), where=b!=0)
        vmsdump_leave_merged_df['vms_pending_hours'] = vmsdump_leave_merged_df['vms_hours'] - vmsdump_leave_merged_df.groupby('WeekEnding').final_output.transform('sum')

        # vmsdump_leave_merged_df.loc[ (vmsdump_leave_merged_df['vms_working_days'] > vmsdump_leave_merged_df['leave_working_days']) & (vmsdump_leave_merged_df['weekend_flag'] == 1) & (vmsdump_leave_merged_df['leave_hours'] == 0) & (vmsdump_leave_merged_df['leave_working_wkenddays'] == 0), 'final_output' ] = 0
        # vmsdump_leave_merged_df['vms_pending_hours'] = vmsdump_leave_merged_df['vms_hours'] - vmsdump_leave_merged_df.groupby('WeekEnding').final_output.transform('sum')

    # This is to handle if VMS days = Leave Tracker days
    # First set handles the weekdays and Second set handles the weekend.
    if (vmsdump_leave_merged_df['vms_working_days'] == vmsdump_leave_merged_df['leave_working_days']).sum():

        vmsdump_leave_merged_df.loc[ (vmsdump_leave_merged_df['vms_working_days'] == vmsdump_leave_merged_df['leave_working_days']) &  (vmsdump_leave_merged_df['weekday_flag'] == 1) & (vmsdump_leave_merged_df['leave_hours'] > 0) & (vmsdump_leave_merged_df['leave_working_wkdays'] > 0), 'final_output'] = np.divide(a, b, out=np.zeros_like(a), where=b!=0)
        vmsdump_leave_merged_df['vms_pending_hours'] = vmsdump_leave_merged_df['vms_hours'] - vmsdump_leave_merged_df.groupby('WeekEnding').final_output.transform('sum')

        vmsdump_leave_merged_df.loc[(vmsdump_leave_merged_df['vms_working_days'] == vmsdump_leave_merged_df['leave_working_days']) &  (vmsdump_leave_merged_df['weekend_flag'] == 1) & (vmsdump_leave_merged_df['leave_hours'] > 0) & (vmsdump_leave_merged_df['leave_working_wkenddays'] > 0) & (vmsdump_leave_merged_df['vms_pending_hours'] > 0), 'final_output'] = np.divide(a, b, out=np.zeros_like(a), where=b!=0)
        vmsdump_leave_merged_df['vms_pending_hours'] = vmsdump_leave_merged_df['vms_hours'] - vmsdump_leave_merged_df.groupby('WeekEnding').final_output.transform('sum')

        vmsdump_leave_merged_df.loc[(vmsdump_leave_merged_df['vms_working_days'] == vmsdump_leave_merged_df['leave_working_days']) &  (vmsdump_leave_merged_df['weekend_flag'] == 1) & (vmsdump_leave_merged_df['leave_hours'] == 0) & (vmsdump_leave_merged_df['leave_working_wkenddays'] == 0) & (vmsdump_leave_merged_df['vms_pending_hours'] > 0), 'final_output'] = 0
        vmsdump_leave_merged_df['vms_pending_hours'] = vmsdump_leave_merged_df['vms_hours'] - vmsdump_leave_merged_df.groupby('WeekEnding').final_output.transform('sum')

    # This is to handle VMS < Leave Tracker days
    # First set handles the weekdays and Second set handles the weekend.
    # Here we are assuring that the leave dates are mapped correctly and rest will be autofilled and highlighted.
    if (vmsdump_leave_merged_df['vms_working_days'] < vmsdump_leave_merged_df['leave_working_days']).sum():
        vmsdump_leave_merged_df.loc[ (vmsdump_leave_merged_df['vms_working_days'] < vmsdump_leave_merged_df['leave_working_days']) & (vmsdump_leave_merged_df['weekday_flag'] == 1) & (vmsdump_leave_merged_df['leave_hours'] == 0) & (vmsdump_leave_merged_df['leave_working_wkdays'] > 0), 'final_output' ] = np.divide(a, b,out=np.zeros_like(a), where=b!=0)
        vmsdump_leave_merged_df['vms_pending_hours'] = vmsdump_leave_merged_df['vms_hours'] - vmsdump_leave_merged_df.groupby('WeekEnding').final_output.transform('sum')

        vmsdump_leave_merged_df.loc[ (vmsdump_leave_merged_df['vms_working_days'] < vmsdump_leave_merged_df['leave_working_days']) & (vmsdump_leave_merged_df['weekend_flag'] == 1) & (vmsdump_leave_merged_df['leave_hours'] == 0) & (vmsdump_leave_merged_df['leave_working_wkenddays'] == 0), 'final_output' ] = 0
        vmsdump_leave_merged_df['vms_pending_hours'] = vmsdump_leave_merged_df['vms_hours'] - vmsdump_leave_merged_df.groupby('WeekEnding').final_output.transform('sum')

    # To make the final_output equal to 0, if the vms_pending_hours is Zero
    vmsdump_leave_merged_df.loc[(vmsdump_leave_merged_df['vms_pending_hours'] == 0) & (vmsdump_leave_merged_df['final_output'].isna()), 'final_output'] = 0

    # Below logic is to fill the values for those dates where we don't have the clarity on the VMS hours
    # The function takes the groups on VMS WEEK and then distributes the vms_pending_hours.
    # If the pending hours are reduced to 0 and still days are left, then those days will be get 0 hours.
    def fill_missing(x):
        hours_counter = x['vms_pending_hours'].mean()
        working_hrs_per_day = 8
        # To check if VMS itself is filled as 0 hours but leave tracker had hours
        if ( (x['vms_working_days'].mean() == 0) & (x['vms_leave_diff'].mean() != 0) ):
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


    vmsdump_leave_merged_df=vmsdump_leave_merged_df.groupby('WeekEnding').apply(fill_missing)

    # The below module resets the columns for easier readability and writes the intermediate data
    # into a sheet. Each sheet will have all the details related to 1 teammate. This is not meant for the Users
    # instead it will help the developers to debug any issue that comes for any teammate.
    reorder_cols=['vms_WeekStarting', 'WeekEnding', 'RACF ID', 'weekday_flag','weekend_flag','Reg Hours', 'OT Hours',
       'leave_working_wkenddays', 'leave_working_wkdays', 'vms_hours', 'leave_hours', 'vms_working_days',
        'leave_working_days', 'vms_leave_diff', 'vms_pending_hours','final_output', 'highlight_flag']
    vmsdump_leave_merged_df = vmsdump_leave_merged_df[reorder_cols]
    vmsdump_leave_merged_df = vmsdump_leave_merged_df.set_index('vms_WeekStarting')
    vmsdump_leave_merged_df['WeekEnding'] = vmsdump_leave_merged_df['WeekEnding'].dt.date
    vmsdump_leave_merged_df.index = vmsdump_leave_merged_df.index.date
    print (vmsdump_leave_merged_df)
    (vmsdump_leave_merged_df
    .to_excel(r'E:\Nikhil\automation\Invoicing_automation\vms_dump_intermediate_generated.xlsx', sheet_name=racf_id,freeze_panes=(1,2)))

    # The below will create the final template that is required by the USER.
    # leave_tracker_df = read_leave_tracker(start_index,end_index,main_range_actual_start,main_range_actual_end)
    # Created a new data frame for the final values.
    vmsdump_leave_final_df = pd.merge(left=vmsdump_leave_merged_df.reset_index()[['index','final_output','highlight_flag']], right=leave_tracker_df[racf_id], how='left', left_on='index', right_on=leave_tracker_df.index.date)
    (vmsdump_leave_final_df.rename(columns={'index' : 'Days',
                                            'final_output' : 'VMS Hours',
                                            racf_id : 'Leave Hours'}, inplace=True))
    vmsdump_leave_final_df.set_index('Days', inplace=True)
    vmsdump_leave_final_df = vmsdump_leave_final_df[['VMS Hours','Leave Hours','highlight_flag']]
    vmsdump_leave_final_df = vmsdump_leave_final_df.T   #Transpose the results as per the required format.
    vmsdump_leave_final_df['RACF ID'] = racf_id
    vmsdump_leave_final_df['Teammate'] = resource_name
    vmsdump_leave_final_df.reset_index(inplace=True)
    vmsdump_leave_final_df.set_index('RACF ID',inplace=True)
    vmsdump_leave_final_df.rename(columns={'index': 'Leave Type'}, inplace=True)
    temp_cols= vmsdump_leave_final_df.columns.tolist()
    temp_cols=temp_cols[-1:] + temp_cols[:-1]
    vmsdump_leave_final_df = vmsdump_leave_final_df[temp_cols]
    appended_final_df_for_styling = appended_final_df_for_styling.append(vmsdump_leave_final_df)
    print (vmsdump_leave_final_df)
    print (appended_final_df_for_styling)
    return appended_final_df_for_styling

def gen_styled_sheet(appended_final_df_for_styling):
    # Highlighting the DataFrame. For styling we need the highlight_flag column, but we cannot keep it in
    # final df, as we don't want that in our output excel sheet. So we split our dataframe where we
    # take the last row in a seperate DF and pass it on to styler seperately only for highlights.
    # Do note, after using styling property, none of DF property will work. Only the styling methods are
    # available to the DF. So we have to manipulate all the data before styling.
    appended_final_df_for_styling=appended_final_df_for_styling.rename_axis('RACF ID').reset_index()  #Styler function can't run on non unique row index.
    print (appended_final_df_for_styling)
    vmsdump_leave_final_styling_df=appended_final_df_for_styling.loc[appended_final_df_for_styling['Leave Type']!='highlight_flag']     #Slicing the last row in a seperate DF.
    # vmsdump_leave_final_styling_df.reset_index(inplace=True)
    print (vmsdump_leave_final_styling_df)

    # Below function will highlight the VMS hours values as yellow wherever the highlight flag is 1.
    def styling(func_df,main_df,):
        y='background-color: yellow'    # Assigns the yellow color in the format of CSS which is needed by Pandas Styler.
        func_df = func_df.T # Transposing the DF created for this function for easier calculations.
        main_df = main_df.T # Transposing the original DF created for easier calculations.
        # print (main_df)
        func_df1=pd.DataFrame('',index=func_df.index, columns=func_df.columns)  #We create empty DF with same index and columns as our input, so that we can pass on just the styling and rest is blank.
        for var in range(len(main_df.columns)//3):      # Implement the styling for each user using loop.
            vms_hours_column = var*3    # Find the VMS column for each user.
            highlight_flag_column=(var*3) + 2   # Find the highlight flag column for each user.
            main_df.loc[(main_df[highlight_flag_column]==1),vms_hours_column] = y  # Use main DF, highlight column.
            main_df.loc[(main_df[highlight_flag_column]!=1),vms_hours_column] = '' # VMS row styler is set as blank string as we don't want any styling for rest.
            func_df1[vms_hours_column]=main_df[vms_hours_column]      #Assign the main df vms column styling  to func df vms column
        func_df1=func_df1.T     #Final transpose to bring it back to its original structure.
        print (func_df1)
        return func_df1

    styled = vmsdump_leave_final_styling_df.style.apply(styling, axis=None, main_df=appended_final_df_for_styling)
    styled.to_excel(r'E:\Nikhil\automation\Invoicing_automation\vms_dump_final_generated_styling.xlsx', engine='openpyxl', sheet_name='VMS_DATA - Generated', index=False, freeze_panes=(1,3))
