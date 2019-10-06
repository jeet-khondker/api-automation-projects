"""
This Is For My Personal Business Project: PyTec LMS
1. Extracting Emails Of Students From Google Spreadsheet
2. Sending Email To The Students Accordingly By Entering Email Subject & Email Message.
"""

# Importing Google Spreadsheet & OAuth2Client Module
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Displaying Data Much More Nicer - Importing Pretty Print Module
from pprint import pprint

# Importing SMTP & Email Libraries
import smtplib
from email.mime.text import MIMEText

# Defining Scope & Creating Credentials Using That Scope & The Content Of JSON File
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

# Adding The Credentials From The Downloaded JSON File
credentials = ServiceAccountCredentials.from_json_keyfile_name("PyTec-CodingGround-151c219d664b.json", scope)

# Authorizing The Google Spreadsheet With The Credentials
client = gspread.authorize(credentials)

# Opening The First Sheet Of Google Spreadsheet File
sheet = client.open("PyTec Learning Ground Students List").sheet1

# Getting A List Of All Records
data = sheet.get_all_records()

# pprint(data)

# Finding A Cell With Exact String Value
cell = sheet.find("Email")
# Getting The Rows & Columns Of The Cell
# print("Found Something At R%sC%s" % (cell.row, cell.col))
# print(cell.row)
# print(cell.col)

# Getting All Emails
column_email = sheet.col_values(cell.col)

# Removing The Email Header To Extract Only Emails
column_header = column_email.pop(0)

print(len(column_email), "Email Addresses")
# pprint(column_email)

recipient_list = column_email
# pprint(recipient_list)

sender = "jeet.usolutionsgroup@gmail.com"
message_subject = input("Enter Your Email Subject: ")
message = input("Enter Your Email Message: ")

# Making The Message To MIMEText
msg = MIMEText(message)

# Taking The Subject
msg["Subject"] = message_subject

# Taking The Sender Email Address
msg["From"] = sender

# Taking The Receiver's Email From The Receipient List
msg["To"] = ", ".join(recipient_list)

try:
    # Using The Google SMTP
    server = smtplib.SMTP("smtp.gmail.com", 587)

    # Identifying To An ESMTP Server
    server.ehlo()

    # Starting The TLS
    server.starttls()

    server.ehlo()

    # Login With Email Adddress & Password
    server.login(sender, "pass123")

    # Sending Email
    server.sendmail(sender, recipient_list, msg.as_string())

    # Closing The Server Connection
    server.close()

    print("Email Sent Successfully!")

except smtplib.SMTPException:
    print("ERROR: Unable To Send Email.")