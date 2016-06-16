import MySQLdb
import xlsxwriter
import mysql.connector  
import datetime  

db = mysql.connector.connect(user='root', password='',
                              host='localhost',
                              database='squareinch')

cursor = db.cursor()




from xml.dom import minidom

doc = minidom.parse("mykong.xml")

majors = doc.getElementsByTagName("major")

for major in majors:

  title = major.getElementsByTagName("title")[0]
  heading = title.firstChild.data
  sql = major.getElementsByTagName("sql")[0]
  location = major.getElementsByTagName("location")[0]
  loc = location.firstChild.data
  query=sql.firstChild.data 
  try:	
          print query
          
          cursor.execute(query)
          
          result = cursor.fetchall()
          num_fields = len(cursor.description)
        
          field_names = [i[0] for i in cursor.description]
          print result
          workbook = xlsxwriter.Workbook(loc)
          worksheet = workbook.add_worksheet()
             
          worksheet.set_column(0,6,5)
          bold = workbook.add_format({'bold': True})
          date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})
          worksheet.write(0,0,heading,bold)
      

          row=1
          col=0
          j = 0
          for rows in field_names:
              worksheet.write(row,col,field_names[j],bold)
              col = col + 1
              j = j + 1
          n=0
          for rows in result:
            col=0
            row = row + 1
            for cols in rows:
                if type(result[n][col]) is datetime.date:
                  worksheet.write(row,col,result[n][col],date_format)
                else:
                  worksheet.write(row,col,result[n][col])
                col = col + 1
            n = n+1
                

            
          
   	
  except Exception as inst:
	 		  print "database & workbook is closing due to Exception"
        
      
      

workbook.close()
db.close()
print "database closed"
print "fine"