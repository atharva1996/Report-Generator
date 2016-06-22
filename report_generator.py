#-----------------------------------------------------------------------------------------------------
#packages

import MySQLdb
import xlsxwriter
import mysql.connector  
import datetime  
import logging
import sys
from mailer import Mailer
from mailer import Message
import smtplib
import base64
import time
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders
from xml.dom import minidom

#--------------------------------------------------------------------------------------

#saving in log file
log = open("report_log.log", "a")
sys.stdout = log

#---------------------------------------------------------------------------------------------------

ts = time.time()
st = datetime.datetime.fromtimestamp(ts).strftime('%d-%m-%Y %H:%M:%S')
#This marks the start of the code execution
print "--------------------------- Scheduled execution of code began : " + st + " -----------------------------------------------"


#----------------------------------------------------------------------------------------------------
#parsing the xml document
doc = minidom.parse("report_details.xml")
reports = doc.getElementsByTagName("report")

#getting platform details of system
get_platform = doc.getElementsByTagName("system")[0]
platform = get_platform.firstChild.data
platform = platform.lower()

#getting senders detail to establishment connection and send the mail
get_sender_addr= doc.getElementsByTagName("sender")[0]
get_sender_pwd= doc.getElementsByTagName("pwd")[0]
get_smtp = doc.getElementsByTagName("smtp")[0]
get_port= doc.getElementsByTagName("port")[0]
fromaddr = get_sender_addr.firstChild.data
msg_pwd = get_sender_pwd.firstChild.data
port = get_port.firstChild.data
msg_smtp = get_smtp.firstChild.data


#Establish connection in order to mail if we require mailing system
mail_flag = 0
for report in reports:
  get_send_mail = report.getElementsByTagName("send_mail")[0]
  send_mail = get_send_mail.firstChild.data
  send_mail = send_mail.lower()
  if send_mail == 'y':
    break

if send_mail == 'y':
  server = smtplib.SMTP(msg_smtp, int(port))
  server.starttls()
  server.login(fromaddr, msg_pwd)
  mail_flag = 1
#-----------------------------------------------------------------------------------------------


#------------------------------------------------------------------------------------------------

#mailer function to send mails with attachment
def mailing_system(reciepent,loc,heading):
  
  reciever = reciepent
  toaddr = reciever.split(',')
  location = loc
  if platform == "windows":
    larray = location.split('\\')
  else:
    larray = location.split('/')
  larray.reverse()
  filename = larray[0]
  msg = MIMEMultipart()
  
  msg['From'] = fromaddr
  msg['To'] = reciever
  ts = time.time()
  st = datetime.datetime.fromtimestamp(ts).strftime('%d-%m-%Y %H:%M:%S')
  msg['Sent'] = st
  msg['Subject'] = heading
  

  body = "Hi, please find the attachments below. Thanks :D"           #body text of the message
 
  msg.attach(MIMEText(body, 'plain'))
 
  attachment = open(location, "rb")
 
  part = MIMEBase('application', 'octet-stream')
  part.set_payload((attachment).read())
  encoders.encode_base64(part)
  part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
 
  msg.attach(part)
 
  
  text = msg.as_string()
  server.sendmail(fromaddr, toaddr, text)
  ts = time.time()
  st = datetime.datetime.fromtimestamp(ts).strftime('%d-%m-%Y %H:%M:%S')
  print "A mail with attachment " + filename + " has been sent at  " + st
  
 
#--------------------------------------------------------------------------------------

#--------------------------------------------------------------------------------------

#gets database connection details


get_user = doc.getElementsByTagName("user")[0]
get_pwd = doc.getElementsByTagName("password")[0]
get_db = doc.getElementsByTagName("dbname")[0]
get_host = doc.getElementsByTagName("host")[0]
user = get_user.firstChild.data
pwd = get_pwd.firstChild.data
dbn = get_db.firstChild.data
host = get_host.firstChild.data

#database connection 
db = mysql.connector.connect(user=user, password=pwd,
                              host=host,
                              database=dbn)
                              
                              

cursor = db.cursor()
#---------------------------------------------------------------------------------------


#--------------------------------------------------------------------------------------

#gets the report details and starts writing to excel sheet
for report in reports:
  get_fire_sql = report.getElementsByTagName("fire_sql")[0]
  fire_sql = get_fire_sql.firstChild.data
  fire_sql=fire_sql.lower()
  if fire_sql == "n":
    continue

  title = report.getElementsByTagName("title")[0]
  heading = title.firstChild.data
  sql = report.getElementsByTagName("sql")[0]
  location = report.getElementsByTagName("location")[0]
  loc = location.firstChild.data
  query=sql.firstChild.data 
  reciever = report.getElementsByTagName("reciever")[0]
  reciepent = reciever.firstChild.data
  get_send_mail = report.getElementsByTagName("send_mail")[0]
  send_mail = get_send_mail.firstChild.data
  send_mail = send_mail.lower()
  try:  

          
          cursor.execute(query)
          ts = time.time()
          st = datetime.datetime.fromtimestamp(ts).strftime('%d-%m-%Y %H:%M:%S')
          print query + "   " + st

          result = cursor.fetchall()
        
          field_names = [i[0] for i in cursor.description]
          
          #----------------------------------------------------------------------------------

          #creating the workbook at the given location and creating its sheet
          workbook = xlsxwriter.Workbook(loc)
          worksheet = workbook.add_worksheet()


          #defining the formnat of the excel sheet created
        
          date_format = workbook.add_format({'text_wrap':1,'align':'left','valign':'top','num_format': 'mmmm d yyyy'})
          time_format = workbook.add_format({'text_wrap':1,'align':'left','valign':'top','num_format': 'hh:mm:ss'})
          timestamp_format = workbook.add_format({'text_wrap':1,'align':'left','valign':'top','num_format': 'dd/mm/yy hh:mm:ss'})
          format = workbook.add_format({'text_wrap':1,'align':'left','valign':'top'})
          format_text = workbook.add_format({'text_wrap':1,'align':'left','valign':'top'})

          size = workbook.add_format()
          align = workbook.add_format()
          bold = workbook.add_format({'bold': True})          

          format.set_border()
          date_format.set_border()
          time_format.set_border()
          timestamp_format.set_border()
          format_text.set_border()
          align.set_border()

          size.set_font_size(20)

          format.set_bg_color('cyan')
            
          align.set_align('left')
          date_format.set_align('left')
          time_format.set_align('left')
          
          timestamp_format.set_align('left')
          
          format.set_bold()

          
          worksheet.write(0,0,heading,size)               #writing the sheet title to excel sheet

          
          worksheet.set_column(0,20,20)                    #adjusting the column size as required
          worksheet.set_default_row(45)                    #adjusting the row size for all the rows


          #writing the table headings to excel sheet with formatting

          row=2
          col=0
          j = 0
          for rows in field_names:

              worksheet.write(row,col,field_names[j],format)

              col = col + 1
              j = j + 1


          #writing the query result to excel sheet with formatting
          
          n=0
          for rows in result:
            col=0
            row = row + 1
            for cols in rows:
                if type(result[n][col]) is datetime.date:
                  worksheet.write(row,col,result[n][col],date_format)
                if type(result[n][col]) is datetime.timedelta :
                  worksheet.write(row,col,result[n][col],time_format)
                if type(result[n][col]) is datetime.datetime:
                  worksheet.write(row,col,result[n][col],timestamp_format)

                else:
                  
                  worksheet.write(row,col,result[n][col],format_text)

                col = col + 1
            n = n+1

          #---------------------------------------------------------------------------------

          #--------------------------------------------------------------------------------

          #calling the mailer function to send mail with attachment
          if send_mail == "y":
            mailing_system(reciepent,loc,heading)      

          #-----------------------------------------------------------------------------------
            
          
    
  except Exception as inst:
        print "database & workbook is closing due to Exception"

  print "\n"
        
#--------------------------------------------------------------------------------------------    
      
#This marks the end of the code execution and we close the mail connection, workbook and database connection
if int(mail_flag) == 1:
 server.quit()
workbook.close()
db.close()
ts = time.time()
st = datetime.datetime.fromtimestamp(ts).strftime('%d-%m-%Y %H:%M:%S')
print "--------------------------- Scheduled execution of code ended : " + st + " -----------------------------------------------\n\n"