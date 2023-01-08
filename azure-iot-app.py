from tkinter import *
import os
import random
import time
from azure.iot.device import IoTHubDeviceClient, Message

sensor_data='{{"Current_Temperature":{temperature},"Current Pressure":{pressure}}}'

def connect():
    global client
    if(section1_Entry1.get()==""):
        section1_text1.delete("1.0",END)
        section1_text1.insert(END,"Enter a valid connection string")
    else:
        try:
            conn_str=str(section1_Entry1.get())
            client=IoTHubDeviceClient.create_from_connection_string(conn_str)
            section1_text1.delete("1.0",END)
            section1_text1.insert(END,str(client))

        except:
            section1_text1.delete("1.0",END)
            section1_text1.insert(END,"Cannot Establish Connection String")


def sendmessage():
    try:
        temperature=20+random.randint(4,8)
        pressure=2+random.randint(1,4)

        formatted_string=sensor_data.format(temperature=temperature,pressure=pressure)

        message=Message(formatted_string)

        if(temperature>27):
           message.custom_properties['temperature_alert']=True
        else:
            message.custom_properties['temperature_alert']=False

        if(pressure>5):
           message.custom_properties['pressure_alert']=True
        else:
            message.custom_properties['pressure_alert']=False
        
        client.send_message(message)

        client.shutdown()
    except:
        section1_text1.delete("1.0",END)
        section1_text1.insert(END,"Message Cannot Be Sent")


# GUI Area
root=Tk()

# Section 1 :

section1=LabelFrame(master=root,text="1: Connection String :")
section1.grid(row=0,column=0)
section1_Label1=Label(master=section1,text="Primary Connection String :")
section1_Label1.grid(row=0,column=0)
section1_Entry1=Entry(master=section1,width=50)
section1_Entry1.grid(row=0,column=1)
section1_Button1=Button(master=section1,text="Connect",command=connect)
section1_Button1.grid(row=0,column=3)
section1_text1=Text(master=section1,height=3,width=30)
section1_text1.grid(row=1,column=0,columnspan=3)



root.mainloop()


