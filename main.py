import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from dotenv import load_dotenv
import os
 
load_dotenv('my.env')


file_path = 'acc.xlsx'
excel_data = pd.read_excel(file_path)


sender_email = os.getenv("EMAIL")
password = os.getenv("PASSWORD")
subject = "Congratulations on Your Acceptance to the Node.js Course at DevZone!"


message_template = """Dear {name},

We are thrilled to inform you that you have been accepted into the Node.js course at DevZone. This is an exciting opportunity to deepen your skills in backend development and to join a community of like-minded learners.

To stay connected with your instructors and fellow students, please join our official WhatsApp group by clicking on the link below:

WhatsApp Group Link: https://chat.whatsapp.com/IW1WtmVbrWsAKdJNOyLt5X

Thank you for choosing DevZone as your learning partner. We look forward to seeing you on this journey and supporting your growth in Node.js development.

Best regards,
DevZone Team
"""
 

with smtplib.SMTP("smtp.gmail.com", 587) as server:
    server.starttls()
    server.login(sender_email, password)
 

    for index, row in excel_data.iterrows():
        try:
         
            personalized_message = message_template.format(name=row["NAME"])
            
           
            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = row["EMAIL"]
            msg["Subject"] = subject
            msg.attach(MIMEText(personalized_message, "plain"))
            
            
            server.sendmail(sender_email, row["EMAIL"], msg.as_string())
            print(f"done {row['NAME']} to {row['EMAIL']}")
        
        except Exception as e:
            print(f"not sent {row['NAME']} to {row['EMAIL']}: {e}")
