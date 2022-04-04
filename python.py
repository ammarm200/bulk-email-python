import pandas as pd
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

html = '''
    <html>

        <body>

            <h1>Heading of the Body</h1>
            <h2>Hello, welcome to everyone!</h2>
            <p>Testing para</p>
            <button1 style="padding=20px;background=Red;color=white;">Subcribe</button1>
            <button2 style="padding=20px;background=Black;color=white;">UnSubcribe</button2>

        </body>
      
    </html>
    '''

#Using pandas library to import emails from the file
data = pd.read_excel("Subscribers.xlsx")
email_column= data.get("email")
list_of_subscribers = list(email_column)
print(list_of_subscribers)

#Sender Gmail id adress
GMAIL_ID = "user@gmail.com"
#sender gmail id password
GMAIL_PSD = "userPassword"



def sendmail():
  server = smtplib.SMTP("smtp.gmail.com",587)
  server.ehlo()
  server.starttls
  try:
    server.login( GMAIL_ID , GMAIL_PSD )
    from_="user@gmail.com"
    to_="list_of_subscribers"
    message= MIMEMultipart()
    message['subject']= "Checking Code"
    message['from']= GMAIL_ID

    # Add HTML/plain text part to MIMEMultipart message
    text=MIMEText(html,"html")
    message.attach(text)
    server.sendmail(from_ ,to_ ,message.as_string())

    count +=1
    print("Message sent successfully to emails")

    if(count%80 == 0):
            server.quit()
            print("120 Seconds cooldown of server")
            time.sleep(120)
            server.ehlo()
            server.login( GMAIL_ID, GMAIL_PSD)


  except Expection as e:
    print(e)
