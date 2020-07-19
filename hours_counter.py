import pandas as pd
pd.set_option('display.max_columns', None)
pd.set_option('display.expand_frame_repr', False)

df = pd.read_excel(r"E:\Nikhil\automation\invoicing\vms_intermediate_sheet.xlsx", header=0, sheet_name='group')

def find_missing(x, col):
    print(type(x))
    print ('\n')
    print(x)
    print(col)
    print(x.index)
    return x

def fill_missing(x):
    hours_counter = x['vms_pending_hours'].mean()
    working_hrs_per_day = 8
    x['highlight_flag'] = 0
    for var in x.index:
        if ( pd.isna(x.loc[var,'final_output']) & (hours_counter != 0) ):
            x.loc[var,'final_output'] = working_hrs_per_day
            x.loc[var,'highlight_flag'] = 1
            hours_counter = hours_counter - working_hrs_per_day
        elif ( pd.isna(x.loc[var,'final_output']) & (hours_counter == 0) ):
            x.loc[var,'final_output'] = 0
            x.loc[var,'highlight_flag'] = 1
    return x


print(df.groupby('vms_week').apply(fill_missing))

# print(df)
