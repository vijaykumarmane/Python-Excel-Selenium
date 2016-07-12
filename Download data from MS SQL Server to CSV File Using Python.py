"""Download data from MS SQL Server to CSV File"""

import adodbapi
import csv

conn_string = "Provider=; Persist Security Info=True; User ID=; Password=;Initial Catalog=; Data Source=,Port#"
connect = adodbapi.connect(conn_string)
curs = connect.cursor()
query = "SELECT * FROM Database WHERE Column > '7/01/15' and Column <= '8/01/15'"
curs.execute(query)
with open(r'File_Path\NewTest.csv', "a") as csv_file:
    csv_writer = csv.writer(csv_file)
    csv_writer.writerow([i[0] for i in curs.description])  # write headers
    csv_writer.writerows(curs)
