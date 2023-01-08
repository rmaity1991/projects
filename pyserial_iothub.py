from xmlrpc.client import boolean
import serial
import time
from tkinter import *
from azure.iot.device import IoTHubDeviceClient, Message
import re


#*******************************************************************#
arduino = serial.Serial(port='COM17', baudrate=9600, timeout=.1)
board_data=[]
#*******************************************************************#
# def write_data(x):
#     arduino.write(bytes(x, 'utf-8'))
#     time.sleep(0.05)
#     data = arduino.readline()
#     return data
#*******************************************************************#

def read_data():
    list_data=[]
    for var in arduino.read(10):
        print("Loop running ..")
        list_data.append(var)
    for var in list_data:
        board_data.append(var)
    print(board_data[1])
    return list_data
#*******************************************************************#
message_text = '{{"temperature": {temp},"regulation": {regSp},"defrost": {defSP},"compressor":{comp},"defrost":{defr}}}'
conn_str ="HostName=rht-dfr-iothub.azure-devices.net;DeviceId=arduino;SharedAccessKey=ro6mT87Q8YCeEs/IdkTvWAMgo6mY8SVxXAIa7pTipjY="
client = IoTHubDeviceClient.create_from_connection_string(conn_str)


#*******************************************************************#
while True:
    data=read_data()
    time.sleep(2)
    # current_temp=data[0]
    # regulationSP=data[1]
    # defrostSP=data[2]
    # compressor=data[3]
    # defrost=data[4]

    # message_text.format(temp=current_temp,regSp=regulationSP,defSP=defrostSP,comp=compressor,defr=defrost)
    # message = Message(message_text)
    # print(message)
    # client.send_message(message)
    # time.sleep(2)


    
