import os
from dotenv import load_dotenv
from email.message import EmailMessage
import openpyxl
import smtplib


load_dotenv()


def sendGrades(EMAIL, PASSWORD, file, course_code):
    # Open the spreadsheet and get the sheet
    workbook = openpyxl.load_workbook(file)
    sheet = workbook.active


    # Set up the SMTP server
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(EMAIL, PASSWORD)

    # Iterate through the rows of the spreadsheet
    for row in sheet.iter_rows(values_only=True):
        name = row[1] + " " + row[2]
        grade = row[3]
        u_email = row[4]
        fName = row[2]

        # Compose and send the email
        msg = EmailMessage()
        text = f"Dear {name},\n\nYour grade for the {course_code} 2nd CA is {grade}.\n\n *DISCLAIMER:This message is automated* \n\nBest regards!"
        msg.set_content(text)
        msg['Subject'] = f"{fName}\'s {course_code} CA 2 Grade"
        msg['From'] = EMAIL
        msg['To'] = u_email

        server.send_message(msg)

    # Close the SMTP server
    server.quit()

sendGrades(os.getenv('EMAIL'), os.getenv('APP_PASSWORD'), 'Test.xlsx', 'ISM 333' )