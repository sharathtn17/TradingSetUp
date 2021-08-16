import requests
import openpyxl
import time
from pathlib import Path
from datetime import date
import calendar
import smtplib
import random
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



today = date.today()
print(calendar.day_name[today.weekday()])
# dd/mm/YY

d1 = today.strftime("%Y-%m-%d")
print("d1 =", d1)


def triggerMail(daylow,dz1,stock1,cmp1,method1):
      mail_content = stock1+' Reached Demand Zone \n'+ "Day's Low: "+str(daylow)+"\n"+"Demand Zone: "+str(dz1)+"\n"+"CMP "+str(cmp1)+"\n"+"Method: "+str(method1)
      #The mail addresses and password
      sender_address = 'sharathoffical73@gmail.com'
      sender_pass = 'PRU$ha2021'
      receiver_address = 'sharathtn17@gmail.com'
      rec_list =  ['sharathtn17@gmail.com']
      rec =  ', '.join(rec_list)
      #Setup the MIME
      message = MIMEMultipart()
      message['From'] = "sharathoffical73@gmail.com"
      message['To'] = rec
      message['Subject'] = 'Watch:'+str(stock1)+ ' Touched Demand Zone'   #The subject line
      #The body and the attachments for the mail
      message.attach(MIMEText(mail_content, 'plain'))
      #Create SMTP session for sending the mail
      # open the file to be sent 
      #filename = "WIT.xlsx"
      #attachment = open("WIT.xlsx", "rb")

      # instance of MIMEBase and named as p
      p = MIMEBase('application', 'octet-stream')
  
      #   To change the payload into encoded form
      #p.set_payload((attachment).read())
  
      # encode into base64
     # encoders.encode_base64(p)
   
      #p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
  
      # attach the instance 'p' to instance 'msg'
      #message.attach(p)
  
      session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
      session.starttls() #enable security
      session.login(sender_address, sender_pass) #login with mail_id and password
      text = message.as_string()
      session.sendmail(sender_address, rec_list, text)
      session.quit()
      print('Mail Sent')


def execute(stock,dz,method,id):
    n = random.randint(0,22)
    print(n)
    apikey2="1YRMPHAGY0KFMGWE"
    apikey1="PQXOVVU3GNFV8DJ7"
    if((n%2)==0):
       url="https://www.alphavantage.co/query?function=TIME_SERIES_DAILY_ADJUSTED&symbol="+str(id)+".BSE&outputsize=full&apikey="+apikey1
       
    else:
       url="https://www.alphavantage.co/query?function=TIME_SERIES_DAILY_ADJUSTED&symbol="+str(id)+".BSE&outputsize=full&apikey="+apikey2
       
    print(url)
    result=requests.get(url).json()
    time.sleep(2)
    daysLow=result.get('Time Series (Daily)').get('2021-08-13').get('3. low')
    cmp=result.get('Time Series (Daily)').get('2021-08-13').get('4. close')
    print("Stock:"+str(stock)+" DZ:"+str(dz)+" day'sLow:"+str(daysLow)+" CMPBSE:"+str(cmp))
    intdz=float(dz)
    if(float(daysLow)<=intdz):
     print("DZ EQUAL TO CMP")
     triggerMail(daysLow,dz,stock,cmp,method)
    time.sleep(2) 

xlsx_file = Path('', 'Stockdz.xlsx')
wb_obj = openpyxl.load_workbook(xlsx_file) 
sheet = wb_obj.active
e_list=[]
i=2
stock="null"
BseID="null"
dz="null"
method="null"


for row in sheet.iter_rows(max_row=sheet.max_row-1):
    #for cell in row:
       stock=sheet.cell(row=i, column=1).value
       BseID=sheet.cell(row=i, column=2).value
       dz=sheet.cell(row=i, column=3).value
       method=sheet.cell(row=i, column=7).value
       time.sleep(3)
       
       
       try:
          execute(stock,dz,method,BseID)
       except Exception as e:
              print(e)
              time.sleep(3)
              print("###Exception####")
            #  url1="https://www.alphavantage.co/query?function=TIME_SERIES_DAILY_ADJUSTED&symbol="+str(stock)+".BSE&outputsize=full&apikey=PQXOVVU3GNFV8DJ7"
             # print(url1)
              
              try:
                  execute(stock,dz,method,stock)
              except Exception:
                     print(stock+" ***** "+str(BseID))
                     tempList=[]
                     tempList.append(stock)
                     tempList.append(BseID)
                     tempList.append(dz)
                     tempList.append(method)
                     e_list.append(tempList)
                     print(e_list)
                     
                     pass
       pass
       
      
       i+=1

print(e_list)
j=0
for item in e_list:
    try:
        execute(e_list[j][0],e_list[j][2],e_list[j][3],e_list[j][1])
    except Exception:
        pass
    j+=1