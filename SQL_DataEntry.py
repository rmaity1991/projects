import pyodbc
import datetime

server = 'localhost\SQLEXPRESS'
database = 'IOT_Database'
username = 'sa'
password = 'Admin@1234'
connection = pyodbc.connect('DRIVER={ODBC Driver 11 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

print(connection)

cursor=connection.cursor()

date=datetime.datetime.now();

cursor.execute('''
INSERT INTO Sensor_Data (Sensor_ID,Sensor_1,Sensor_2,Sensor_3)
VALUES
(1,700,800,150),
(2,1700,1800,1940),
(3,1750,8050,6789)
''')

cursor.commit()

connection.close()

