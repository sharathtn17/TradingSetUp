import smtplib
import sys
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

inputArg=sys.argv[1].strip()
file=open("properties.txt","r")
for line in file:
    method=line.split(":",1)
    #print(inputArg+"PARM")
    #print(method[0].strip()+"URL")
    m=method[0].strip()
    if ( inputArg == m):
      # print(method[1]+" URL")
       URL=method[1]
       #print(method[1]+"URL TUR####")
    else:
        print("NOT FOUND")
print(URL+" URL")
mail_content = inputArg+' Report. Please find the attachment. URL:'+URL
#The mail addresses and password
sender_address = 'sharathoffical73@gmail.com'
sender_pass = 'PRU$ha2021'
receiver_address = 'sharathtn17@gmail.com'
rec_list =  ['sharathtn17@gmail.com', 'suhassj47@gmail.com', 'sharathtn73@gmail.com']
rec =  ', '.join(rec_list)
#Setup the MIME
message = MIMEMultipart()
message['From'] = "sharathoffical73@gmail.com"
message['To'] = rec
message['Subject'] = 'Todays '+inputArg+' report '   #The subject line
#The body and the attachments for the mail
message.attach(MIMEText(mail_content, 'plain'))
#Create SMTP session for sending the mail
# open the file to be sent 
filename = inputArg+".xlsx"
attachment = open(filename , "rb")

# instance of MIMEBase and named as p
p = MIMEBase('application', 'octet-stream')
  
# To change the payload into encoded form
p.set_payload((attachment).read())
  
# encode into base64
encoders.encode_base64(p)
   
p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
  
# attach the instance 'p' to instance 'msg'
message.attach(p)
  
session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
session.starttls() #enable security
session.login(sender_address, sender_pass) #login with mail_id and password
text = message.as_string()
session.sendmail(sender_address, rec_list, text)
session.quit()
print('Mail Sent')