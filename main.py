import psycopg2
import pandas as pd
from datetime import datetime

import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart

#Establish Connection to database
database = ""
user = ""
password = ""
host = ""
port = ""

conn = psycopg2.connect(
    database=database,
    user=user,
    password=password,
    host=host,
    port=port
)

#create a cursor object 
cursor = conn.cursor()

# create sql query
query = """
select name, location_latitude, location_longitude from lpg_tracking_system_station;
"""

# fetch the data
cursor.execute(query)

#columns of fetched data
columns = [column[0] for column in cursor.description]

# list to store all fetched data, fetched data will be put on a dictionary
returned_data = []
for row in cursor.fetchall():
    returned_data.append(dict(zip(columns, row)))
#close cursor and connection.
cursor.close()
conn.close()

#convert the returned data to pandas dataframe.
returned_data_frame = pd.DataFrame(returned_data)

#create excel writer object
report_file_name = datetime.now().strftime('%d%m%Y_%H_%M')+'report.xlsx'
writer = pd.ExcelWriter(report_file_name)
returned_data_frame.to_excel(writer)
writer.save()

#begin sending email
#define details

sender = ""
password = ""
recipients = [''] #list all emails

outer = MIMEMultipart()
outer['Subject'] = 'Email Subject'
outer['To'] = ','.join(recipients)
outer['From'] = sender
outer.preamble = "I cannot see the attachment"

print(os.path.basename(report_file_name))

try:
    with open(os.path.basename(report_file_name), 'rb') as fp:
        msg = MIMEBase('application', "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        msg.set_payload(fp.read())
    encoders.encode_base64(msg)
    msg.add_header('Content-Description', 'attachment', filename=os.path.basename(report_file_name))
    outer.attach(msg)
except:
    print("Unable to open attachement")
    raise

composed = outer.as_string()

#try sending the email.
try:
    with smtplib.SMTP('mail host', 587) as s:
        s.ehlo()
        s.starttls()
        s.ehlo()
        s.login(sender, password)
        s.sendmail(sender, recipients, composed)
        s.close()
except:
    print("Unabale to send email")
    raise