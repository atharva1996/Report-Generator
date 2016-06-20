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

#---------------------------------------------------------------------------------------------------

#------------------------------------------------------------------------------------------------
#parsing the xml document
doc = minidom.parse("mykong.xml")

#getting platform details of system
get_platform = doc.getElementsByTagName("system")[0]
platform = get_platform.firstChild.data

#getting senders detail to send mail
get_sender_addr= doc.getElementsByTagName("sender")[0]
get_sender_pwd= doc.getElementsByTagName("pwd")[0]
get_smtp = doc.getElementsByTagName("smtp")[0]
get_port= doc.getElementsByTagName("port")[0]

#-----------------------------------------------------------------------------------------------


#------------------------------------------------------------------------------------------------

#mailer function to send mails with attachment
def mailing_system(reciepent,loc,heading):
  fromaddr = get_sender_addr.firstChild.data
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
  
  msg['Subject'] = heading
  msg_pwd = get_sender_pwd.firstChild.data
  port = get_port.firstChild.data
  msg_smtp = get_smtp.firstChild.data

  body = "Hi, please find the attachments below. Thanks :D"           #body text of the message
 
  msg.attach(MIMEText(body, 'plain'))
 
  attachment = open(location, "rb")
 
  part = MIMEBase('application', 'octet-stream')
  part.set_payload((attachment).read())
  encoders.encode_base64(part)
  part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
 
  msg.attach(part)
 
  server = smtplib.SMTP(msg_smtp, int(port))
  server.starttls()
  server.login(fromaddr, msg_pwd)
  text = msg.as_string()
  server.sendmail(fromaddr, toaddr, text)
  ts = time.time()
  st = datetime.datetime.fromtimestamp(ts).strftime('%d-%m-%Y %H:%M:%S')
  print "A mail with attachment " + filename + " has been sent at  " + st
  server.quit()
 
#--------------------------------------------------------------------------------------

#--------------------------------------------------------------------------------------

#gets database connection details
reports = doc.getElementsByTagName("report")

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

#saving in log file
log = open("myprog.log", "a")
sys.stdout = log

#--------------------------------------------------------------------------------------

#gets the report details and starts writing to excel sheet
for report in reports:

  title = report.getElementsByTagName("title")[0]
  heading = title.firstChild.data
  sql = report.getElementsByTagName("sql")[0]
  location = report.getElementsByTagName("location")[0]
  loc = location.firstChild.data
  query=sql.firstChild.data 
  reciever = report.getElementsByTagName("reciever")[0]
  reciepent = reciever.firstChild.data
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
        
          bold = workbook.add_format({'bold': True})
          date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})
          time_format = workbook.add_format({'num_format': 'hh:mm:ss'})
          timestamp_format = workbook.add_format({'num_format': 'dd/mm/yy hh:mm:ss'})
          format = workbook.add_format()
          size = workbook.add_format()
          align = workbook.add_format()
          format.set_border()
          date_format.set_border()
          time_format.set_border()
          timestamp_format.set_border()
          align.set_border()
          format.set_bg_color('cyan')
          size.set_font_size(20)  
          align.set_align('left')
          date_format.set_align('left')
          time_format.set_align('left')
          timestamp_format.set_align('left')
          format.set_bold()

          
          worksheet.write(0,0,heading,size)               #writing the sheet title to excel sheet

          
          worksheet.set_column(0,6,10)                    #adjusting the column size as required
          


          #writing the table headings to excel sheet with formatting

          row=1
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
                  
                  worksheet.write(row,col,result[n][col],align)

                col = col + 1
            n = n+1

          #---------------------------------------------------------------------------------

          #--------------------------------------------------------------------------------

          #calling the mailer function to send mail with attachment
          mailing_system(reciepent,loc,heading)      

          #-----------------------------------------------------------------------------------


            
          
    
  except Exception as inst:
        print "database & workbook is closing due to Exception"
        
#--------------------------------------------------------------------------------------------    
      

workbook.close()
db.close()
print "database closed"
print "fine\n\n"