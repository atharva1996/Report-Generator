import MySQLdb
import xlsxwriter
import mysql.connector    

db = mysql.connector.connect(user='root', password='',
                              host='localhost',
                              database='squareinch')

cursor = db.cursor()


workbook = xlsxwriter.Workbook('demo.xlsx')


from xml.dom import minidom

doc = minidom.parse("mykong.xml")

majors = doc.getElementsByTagName("major")

for major in majors:
	sql = major.getElementsByTagName("sql")[0]
	query=sql.firstChild.data 
		
	try:	
          print query
          
          cursor.execute(query)
          
          result = cursor.fetchall()
          print result
          worksheet = workbook.add_worksheet()
          worksheet.set_column('A:A', 20)
          row=0
          col=0
          for rows in result:
            col=0
            
            for cols in rows:

                worksheet.write(row,col,result[row][col])
                col = col + 1
                

            row = row + 1
          
   	
    	except Exception as inst:
	 		  print "database & workbook is closing due to Exception"
        
      
      

workbook.close()
db.close()
print "database closed"
print "fine"