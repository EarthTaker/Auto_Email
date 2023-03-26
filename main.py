# Import libraries
import pandas as pd
import smtplib
import os
import glob
import ssl
import datetime as dt

# Import functionality from classes text, multipart, and application all from superclass email.
# Import specific classes from the email library
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Function to Fill Excelsheet with this week's default hours
def fill_Excelsheet_Template(first_day_of_week, last_day_of_week):
    
    # Copy existing empty template from the directory.
    timesheetCopy = pd.read_excel('D:\Collegia_Terranus\CoffeeTree\TimeReport_GarrisonGeho (Template).xlsx').copy()
    
    # To prevent the display date within the file name from incrementing as well.
    first_day_of_week_increment = first_day_of_week
    
    # Edit lines within timesheet
    # Add the first day of the week to column A, row 2, incrementing up 4 days until you reach column A, row 6
    for i in range(0, 5):
        timesheetCopy.at[i, 'Date'] = first_day_of_week_increment.strftime('%m-%d-%Y') 
        
        first_day_of_week_increment += dt.timedelta(days=1)
        
    # Set the Start Time, i.e., column D, row 2, to 8:00 AM. Read back row value to make sure it stores as 8:00:00 AM. Only apply this value to Column D, rows 2 through 6.
    timesheetCopy.loc[0:4, 'Start Time'] = pd.to_datetime('8:00:00 AM', format='%I:%M:%S %p').strftime('%I:%M:%S %p')

    # Set the End Time, i.e., column E, row 2, to 4:00 PM. Read back row value to make sure it stores as 4:00:00 PM. Only apply this value to Column E, rows 2 through 6.
    timesheetCopy.loc[0:4, 'End Time'] = pd.to_datetime('4:00:00 PM', format='%I:%M:%S %p').strftime('%I:%M:%S %p')

    # Calculate the difference between 'Start Time' and 'End Time' in seconds and convert it to hours
    timesheetCopy['Hours Worked'] = (pd.to_datetime(timesheetCopy['End Time'], format='%I:%M:%S %p') - pd.to_datetime(timesheetCopy['Start Time'], format='%I:%M:%S %p')) / pd.Timedelta(hours=1)
    
    #Get total number of hours worked. 
    timesheetCopy.loc[timesheetCopy.index[-1], 'Date'] = 'Total'
    
    #Grab the sum of the Hours Worked Column, add it to the associated row.
    timesheetCopy.loc[timesheetCopy.index[-1], 'Hours Worked'] = timesheetCopy['Hours Worked'].sum()
    
    # Create writer object to generate the Excelsheet.
    writer = pd.ExcelWriter('D:\\Collegia_Terranus\\CoffeeTree\\TimeReport_GarrisonGeho ({} - {}).xlsx'.format(first_day_of_week, last_day_of_week), engine='xlsxwriter')
    
    # Copy the current timesheet template
    timesheetCopy.to_excel(writer, index=False)
    
    # Auto-fit column width
    for i, col in enumerate(timesheetCopy.columns):
        writer.sheets['Sheet1'].set_column(i, i, max(timesheetCopy[col].astype(str).str.len() + 10) + 1)
    
    # Close writer
    writer.close()

# Function to generate date range for email body, i.e., MM-DD-YYYY -> MM-DD-YYYY
def generate_Date_Range():
    
    #Get the current date
    today = dt.date.today()
    
    # Get the first day of the week based on the current date
    # Subtract today's date from the duration of time since today's date
    first_day_of_week = today - dt.timedelta(days=today.weekday())
    
    #Get the date of the friday of each week
    last_day_of_week = first_day_of_week + dt.timedelta(days=4)
    
    return first_day_of_week, last_day_of_week

# Function to send email.
def send_email(to, subject, body):
    # Set up the email message
    # Build out the email template, assign it to the object emailMessage.
    emailMessage = MIMEMultipart()
    
    # Assign values to the associated variables within that email template.
    emailMessage['From'] = 'garrisongeho1992@gmail.com'
    emailMessage['To'] = to
    emailMessage['Subject'] = subject
    
    # Format the first day of the week date time to mm/dd/yyyy, concatinate -> between it and the formatted last day of the week (friday) of mm/dd/yyyy
    date_range = first_day_of_week.strftime('%m-%d-%Y') + ' -> ' + last_day_of_week.strftime('%m-%d-%Y')
    
    # Format the body of the email
    body = """
    Katie, 

    I've attached my timesheet above for the week: {}.

    Thank you for Taking Time out of Your Day to Invest in Mine, 
    Garrison G.

    LinkedIn: https://www.linkedin.com/in/garrison-geho-06246aa1/
    GitHub: https://github.com/EarthTaker
    ArtStation: https://earthtaker101.artstation.com
    Facebook: https://www.facebook.com/EarthTaker
    YouTube: https://www.youtube.com/channel/UCSxKU9Mg94CmXvu7jo8Wd6g

    """.format(date_range)

    # Add the body of the email
    emailMessage.attach(MIMEText(body, 'plain'))
    
    # Establish file pattern in which glob can look in the CoffeeTree directory 
    file_pattern = "TimeReport_GarrisonGeho (*).xlsx"
    attachment_path = glob.glob(os.path.join('D:\Collegia_Terranus\CoffeeTree', file_pattern))[0]

    # Add the timesheets file as an attachment
    with open(attachment_path, "rb") as f: # <- With opens the file at the file path. Closes the file once this block is executed too.
        attach = MIMEApplication(f.read(),_subtype="xlsx") 
        attach.add_header('Content-Disposition','attachment',filename="Timesheet: {} - {}".format(first_day_of_week, last_day_of_week))
        emailMessage.attach(attach)

    # Set up the SMTP server
    # smtp_server = 'smtp.office365.com' #Office 365 #Cannot set up unless CoffeeTreeGroup enables SMTP Auth OR Allows me the settings to generate an app password. 
    smtp_server = "smtp.gmail.com"
    port = 587
    sender_email = 'garrisongeho1992@gmail.com'
    password = input("Type your password and press enter: ") # evyxgvudqjvlxzkv <- Google App Password
    
    # Create secure SSL connection
    context = ssl.create_default_context()

    try: 
        # Start the SMTP server
        server = smtplib.SMTP(smtp_server, port)
        
        server.starttls(context=context)
        
        # Login to the email account
        server.login(sender_email, password) 
        
        # Send the email
        server.sendmail(sender_email, to, emailMessage.as_string())
        
    except Exception as e:
        print(e)

    # Close the SMTP server
    server.quit()
    
    print("Email successfully sent to: {}. Date Range: {} -> {}".format(to, first_day_of_week, last_day_of_week))

# Grab first and last day of week
first_day_of_week, last_day_of_week = generate_Date_Range()

# Fill Excelsheet with this week's default hours: Pass excelsheet and first day of week.
fill_Excelsheet_Template(first_day_of_week, last_day_of_week)

# Send the email
send_email('garrisongeh@coffeetreegroup.com', 'Timesheet: {} -> {} '.format(first_day_of_week, last_day_of_week))