import tkinter
import serial
import time
from tkinter import *
from tkinter import ttk
from azure.iot.device import IoTHubDeviceClient, Message
import re

info=[]
message_text = '{{"temp": {temp},"regulationSP": {regulation},"compressor": {compressor}}}'
compressor=0

def read_info():
    comport=str(frame1_entry2.get())
    conn_str=frame1_entry1.get()
    baudrate=int(frame1_entry3.get())

    info.append(comport)
    info.append(conn_str)
    info.append(baudrate)

    return info

def connect(command):

    info=read_info()

    try:
        arduino = serial.Serial(port=info[0], baudrate=info[2], timeout=.1)
        conn_str =str(info[1])
        client = IoTHubDeviceClient.create_from_connection_string(conn_str)

        while(command):
            data=int.from_bytes(arduino.readall(),'big')
            regulationSP=13
            if(data>50):
               data=last_data
            if(data>regulationSP):
               compressor=1 
            else:
               compressor=0
            message_text_format=message_text.format(temp=data,regulation=regulationSP,compressor=compressor)
            message = Message(message_text_format)
            if data > regulationSP:
                message.custom_properties["temperatureAlert"] = "true"
            else:
                message.custom_properties["temperatureAlert"] = "false"
            frame1_textarea1.delete("1.0", END)
            frame1_textarea1.insert(END, message)
            client.send_message(message)
            last_data=data
            time.sleep(2)



    except:
        frame1_textarea1.delete("1.0", END)
        frame1_textarea1.insert(END, "Connection Issue: Check Info")


root=Tk()
root.title("Azure IOT Application")

frame1=LabelFrame(master=root,text="1.1 Communication Info")
frame1.grid(row=0,column=0)

frame1_label1=Label(master=frame1,text="Connection String :")
frame1_label1.grid(row=0,column=0)

frame1_entry1=Entry(master=frame1,width=40)
frame1_entry1.grid(row=0,column=1)
frame1_entry1.insert(0,"Connection String")


frame1_label2=Label(master=frame1,text="COM Port  For Micro Controller:")
frame1_label2.grid(row=1,column=0)

frame1_entry2=Entry(master=frame1,width=40)
frame1_entry2.grid(row=1,column=1)
frame1_entry2.insert(0,"COM10")

frame1_label3=Label(master=frame1,text="Baud Rate :")
frame1_label3.grid(row=2,column=0)

frame1_entry3=Entry(master=frame1,width=40)
frame1_entry3.grid(row=2,column=1)
frame1_entry3.insert(0,"9600")

frame1_button1=Button(master=frame1,text="Send To Cloud",command=lambda :connect(True))
frame1_button1.grid(row=3,column=0)


frame1_button1=Button(master=frame1,text="Stop Connection",command=lambda :connect(False))
frame1_button1.grid(row=3,column=2)
frame1_textarea1 = Text(master=frame1, height=3, width=50)
frame1_textarea1.grid(row=4, column=0)


root.mainloop()
    
