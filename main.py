import xlrd
import psycopg2
from psycopg2 import Error

#----Postgray connection establishment----
try:
    connection = psycopg2.connect(user="postgres", password="12345", host="127.0.0.1", port="5432", database="qtec")
    cur = connection.cursor()

#----Table Creation----
    create_table_query = '''CREATE TABLE IF NOT EXISTS saleHistory
          (year INT PRIMARY KEY     NOT NULL,
          amount           INT    NOT NULL); '''
    cur.execute(create_table_query)
    connection.commit()
    print("Table created successfully in PostgreSQL ")

except (Exception, psycopg2.DatabaseError) as error:
    print("Error while creating PostgreSQL table", error)

#----Excel File import----
excelLoc= ("sales.xlsx")
uploadedFile = xlrd.open_workbook(excelLoc)
sheet = uploadedFile.sheet_by_index(0)  #First sheet selection from excel sheet

#----Data Import from sheet----
query = """INSERT INTO saleHistory (year, amount) VALUES (%s, %s) ON CONFLICT DO NOTHING"""
for r in range(1, sheet.nrows):  #Range from 1 due to avoid Header
    year = sheet.cell(r,0).value
    amount = sheet.cell(r,1).value
    values = (year, amount)
    cur.execute(query, values)

connection.commit()
print("Successfully imported Excel into postgreSQL \n")

#----Show Data from postGray----
view_query="select * from saleHistory"
cur.execute(view_query)
sale_records = cur.fetchall()

print("Here is the values of Sale History Table\n")
print("Year \t Amount")
for row in sale_records:
    print(row[0],"\t", row[1])

#----closing database connection----
if (connection):
    cur.close()
    connection.close()
    print("\nPostgreSQL connection is closed")