import pyodbc

server = 'localhost\SQLEXPRESS'
database = 'IOT_Database'
username = 'sa'
password = 'Admin@1234'
connection = pyodbc.connect('DRIVER={ODBC Driver 11 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

print(connection)

cursor=connection.cursor()

cursor.execute("""
CREATE TABLE Sensor_Data
(
Sensor_ID int primary key,
Sensor_1 real,
Sensor_2 real,
Sensor_3 real
)
""")

connection.commit()


connection.close()