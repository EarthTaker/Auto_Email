import pandas as pd
import datetime as dt

df = pd.read_excel('D:\Collegia_Terranus\CoffeeTree\TimeReport_GarrisonGeho (Template).xlsx').copy()

today = dt.date.today()

first_day_of_week = today - dt.timedelta(days=today.weekday())

print(first_day_of_week)

print(df)

# Function to Fill Excelsheet with this week's default hours
def fill_Excelsheet_Template(df, first_day_of_week, last_day_of_week):
    
    # Edit lines within timesheet
    # Add the first day of the week to column A, row 2, incrementing up 4 days until you reach column A, row 6
    for i in range(5):
        df.at[i+1, 'Date'] = first_day_of_week
        first_day_of_week += dt.timedelta(days=1)
        
        
    df.to_Excel('TimeReport_GarrisonGeho ({} -> {}).xlsx'.format(first_day_of_week, last_day_of_week))