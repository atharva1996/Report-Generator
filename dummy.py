import MySQLdb
import xlsxwriter
import mysql.connector    

db = mysql.connector.connect(user='root', password='',
                              host='localhost',
                              database='squareinch')

cursor = db.cursor()
row=0
workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()

worksheet.set_column('A:A', 20)

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
   		worksheet.write(row,0,result[0][0])
   		row = row + 1;
   		print result[0][0]

    	except Exception as inst:
	 		print "database & workbook is closing due to Exception"
	 		db.close()

workbook.close()
db.close()
print "database closed"
print "fine"