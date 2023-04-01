import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
import datetime as DTM

# Function to generate date range for email body, i.e., MM-DD-YYYY -> MM-DD-YYYY
def generate_Date_Range():
    
    #Get the current date 
    today = DTM.date.today()
    
    # Get the first day of the week based on the current date
    first_day_of_week = today - DTM.timedelta(days=today.weekday())
    
    #Get the date of the friday of each week
    last_day_of_week = first_day_of_week + DTM.timedelta(days=4)
    
    #Get date range between first and last day of week
    dateRange = pd.date_range(start=first_day_of_week, end=last_day_of_week, freq='D')
    
    #Return populated range of dates and the first and last day of week.
    return dateRange, first_day_of_week, last_day_of_week

# Main Function - Application Starting Point
def prompt_user():
    
    while True:
        print()
        print("<============================================================>")
        print("Timesheet Submission App:")
        print("<============================================================>")
        print()
        print("Option 1: Auto Populate")
        print()
        print("Option 2: Manual Populate")
        print()
        print("Option 3: Add to Online Excel Workbook & Exit")
        
        #Set up while loop condition
        choice = input("Awaiting Input: ")
        
        #Regardless of timesheet generation, grab first and last day of week.
        dateRange, first_day, last_day = generate_Date_Range()
        
        #Create empty data Dictionary
        data = {
            'Date': [],
            'Customer:Team': [],
            'Name': [],
            'Start Time': [],
            'End Time': [],
            'Hours Worked': []
        }
        
        #Create Data frame from data Dictionary
        dataFrame = pd.DataFrame(data, index=None).fillna('')
        
        #Format data frame
        dataFrame = dataFrame[['Date', 'Customer:Team', 'Name', 'Start Time', 'End Time', 'Hours Worked']]
        
        #Begin auto generation
        if choice == "1":
            dataFrame = populate_data_auto(dataFrame, dateRange)
            
            print(dataFrame.to_string(index=False))
        
        #Begin manual generation
        elif choice == "2":
            
            dataFrame = populate_data_manual(dataFrame, dateRange)
            
            print(dataFrame.to_string(index=False))
        
        #Add Worksheet to online excel workbook
        elif choice == "3":
            
            print()
            print("<====================Confirm Timesheet=======================>")
            print()
            print(dataFrame.to_string(index=False))
            print()
            print("<============================================================>")
            print()
            
            submitNow = input("Y / N").lower()
            
            if submitNow == True:
                print("Generate Total for all hours")
            
            #Call to function that adds sheet to online workbook
            return False
        
        else:
            print("Stop wasting time. This was meant for automation, not procrastination.")
            
# Function to populate week with default values, i.e., 8 Hours Worked, 8:00:00 AM -> 4:00:00 PM 
def populate_data_auto(dataFrame, dateRange):

    # For each date in the date range passed
    for i, date in enumerate(dateRange):
        
        #Populate Data frame with default info.
        dataFrame.loc[i, "Date"] = date.strftime('%m/%d/%Y')
        dataFrame.loc[i, "Customer:Team"] = "LCE: CV Support L2"
        dataFrame.loc[i, "Name"] = "Garrison Geho"
        dataFrame.loc[i, "Start Time"] = "8:00:00 AM"
        dataFrame.loc[i, "End Time"] = "4:00:00 PM"
        dataFrame.loc[i, "Hours Worked"] = 8
        
    # Calculate total hours worked
    dataFrame.loc[6, 'Date'] = 'Total:'
    dataFrame.loc[6, 'Hours Worked'] = dataFrame['Hours Worked'].sum()
    
    print()
    print("<=====================================================>")
    print()
    print("Affected Date Range: {} -> {}.".format(dataFrame.loc[0, "Date"], dataFrame.loc[4, "Date"]))
    print()
    print("<=====================================================>")
    print()
    
    return dataFrame

# Function to populate week with custom values
def populate_data_manual(dataFrame, dateRange):
    
    # For each date in the date range
    for i, date in enumerate(dateRange):
        
        print()
        print("<=====================================================>")
        print()
        print("Affecting Date: {}.".format(date))
        print()
        print("<=====================================================>")
        print()

        # For each key in the data dictionary
        for key in dataFrame.keys():
            
            #Create Boolean to control while loop
            populatingDataFrame = True
        
            if key == "Date":
            
                dataFrame.loc[i, key] = date.strftime('%m/%d/%Y')
                continue
            
            elif key == "Customer:Team":
                
                dataFrame.loc[i, key] = 'LCE: CV Support L2'
                continue
                
            elif key == "Name":
                
                dataFrame.loc[i, key] = 'Garrison Geho'
                continue
                
            # Keep asking for input until a valid value is entered
            while populatingDataFrame:
                    
                # Prompt the user for input for the current key
                if key != "Hours Worked":
                    value = input(f"Input value for {key}: ")
                
                try:                                        
                    if key == "Start Time" or key == "End Time":
                            
                        #Attempt to format the value to 12-Hour format with am/pm
                        value = DTM.datetime.strptime(value, "%I %p")
                        
                        #Format to 12-Hour format with am/pm
                        value = value.strftime("%I:%M:%S %p")
                        
                        dataFrame.loc[i, key] = value
                        
                        break
                    
                    elif key == "Hours Worked":
                
                        dataFrame.loc[i, key] = (pd.to_datetime(dataFrame.loc[i, 'End Time']) - pd.to_datetime(dataFrame.loc[i, 'Start Time'])) / pd.Timedelta(hours=1)
                        
                        populatingDataFrame = False
                
                #Catch any invalid entries.
                except ValueError:
                    print(f"Key: {key} does not take {value}.")
                    continue
    
    # Calculate total hours worked
    dataFrame.loc[6, 'Date'] = 'Total:'
    dataFrame.loc[6, 'Hours Worked'] = dataFrame['Hours Worked'].sum()
    
    # Return the updated data frame
    return dataFrame

if __name__ == '__main__':
    prompt_user()