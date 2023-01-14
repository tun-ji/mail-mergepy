import os
from dotenv import load_dotenv
from email.message import EmailMessage
import openpyxl
import smtplib

# Open the spreadsheet and get the sheet
workbook = openpyxl.load_workbook('Test.xlsx')
sheet = workbook.active

load_dotenv()
EMAIL = os.getenv('EMAIL')
PASSWORD = os.getenv('APP_PASSWORD')

# Set up the SMTP server
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(EMAIL, PASSWORD)

# Iterate through the rows of the spreadsheet
for row in sheet.iter_rows(values_only=True):
    name = row[1] + " " + row[2]
    grade = row[3]
    email = row[4]
    fName = row[2]

    msg = EmailMessage()
    text = f"Dear {name},\n\nYour grade for the ISM 305 2nd CA is {grade}.\n\n *DISCLAIMER:This message is automated on behalf of ***REMOVED**** \n\nBest regards!"
    msg.set_content(text)
    msg['Subject'] = f"{fName}\'s ISM 305 CA 2 Grade"
    msg['From'] = EMAIL
    msg['To'] = email
    # Compose and send the email
    
    server.send_message(msg)

# Close the SMTP server
server.quit()