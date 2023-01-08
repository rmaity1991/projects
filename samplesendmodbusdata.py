from pyModbusTCP.client import ModbusClient
from tkinter import *
from azure.iot.device import IoTHubDeviceClient,Message
import time

conn_str="HostName=rht-dfr-iothub.azure-devices.net;DeviceId=sensor1;SharedAccessKey=8zfoS7qmbFmOHjlZzSq4PIwUCFbbsnjkSUPrXyoyTfw="

message_text='{{"temperature":{temperature},"pressure":{pressure},"dewpoint":{dewpoint},"humidity":{humidity}}}'
client=IoTHubDeviceClient.create_from_connection_string(conn_str)


c = ModbusClient(host="127.0.0.1", auto_open=True, auto_close=True,port=502)



# reading holding registers (40001 = 0) so start from 0
count=int(input("Enter the number of times the message needs to be sent? "))
while(count>0):
    regs = c.read_holding_registers(0, 4)
    formatted_text=message_text.format(temperature=regs[0],pressure=regs[1],dewpoint=regs[2],humidity=regs[3])
    message=Message(formatted_text)
    client.send_message(message)
    
    count=count-1
    time.sleep(2)



    