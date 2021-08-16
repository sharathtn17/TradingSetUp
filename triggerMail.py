import smtplib
import sys
import itertools
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

URL=""
inputArg=sys.argv[1].strip()
file=open("properties.txt","r")
#for inpu in sys.argv:
#      #print(inpu)
#      for line in file:
       # print(line)
#          method=line.split(":",1)
        #print(method)
      #print(inputArg+"PARM")
      #print(method[0].strip()+"URL")
#          m=method[0].strip()
#          print(m)
#          if ( inpu == m):
#              print(method[1]+" URL")
#              URL=method[1]
#              prnit(URL)
#   continue
       
for (a) in zip(sys.argv, file):
 print (a)    

print(URL+" URL")
mail_content = inputArg+' Report. Please find the attachment. URL:'+URL
#The mail addresses and password
sender_address = 'sharathoffical73@gmail.com'
sender_pass = 'PRU$ha2021'
receiver_address = 'sharathtn17@gmail.com'
rec_list =  ['sharathtn17@gmail.com', 'sharathtn73@gmail.com']
rec =  ', '.join(rec_list)
#Setup the MIME
message = MIMEMultipart()
message['From'] = "sharathoffical73@gmail.com"
message['To'] = rec
message['Subject'] = 'Todays '+inputArg+' report '   #The subject line
#The body and the attachments for the mail
message.attach(MIMEText(mail_content, 'plain'))
for file in sys.argv:
 #Create SMTP session for sending the mail
 # open the file to be sent 
 if file != "triggerMail.py" :
   filename = file+".xlsx"
   attachment = open("report/"+filename , "rb")
   print(filename)
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
#session.sendmail(sender_address, rec_list, text)
session.quit()
print('Mail Sent')