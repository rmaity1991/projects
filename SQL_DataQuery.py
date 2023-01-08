import pyodbc
import datetime

server = 'localhost\SQLEXPRESS'
database = 'IOT_Database'
username = 'sa'
password = 'Admin@1234'
connection = pyodbc.connect('DRIVER={ODBC Driver 11 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

if(connection):
    print("Connection Successful");
else:
    print("Connection Failed");

cursor=connection.cursor()

cursor.execute(
    """
    SELECT * FROM IOT_Database.dbo.Sensor_Data
    
    """
)

row = cursor.fetchone() 
#row = cursor.fetchall()

print(row)

while(row):
    print(row[0])
    print(row[1])
    print(row[2])
    #print(row[3])
    #break

    row=cursor.fetchone()

connection.close()


