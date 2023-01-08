"""
    Initial Release : Rohit Maity
    Date : 21/08/2022
"""
from pymodbus.client.sync import ModbusSerialClient as ModbusClient 
import time
from tkinter import *
import pandas as pd
import numpy as np
import os
import re
#*************************************************************************************#
# MpxOne Modbus Address List
#*************************************************************************************#
#*************************************************************************************#

def extractdata():
    Comm_Type = frame1_entry1.get()
    COM_Port = frame1_entry2.get()
    Baud_Rate = frame1_entry3.get()
    Bits =frame1_entry4.get()
    Parity = frame1_entry5.get()
    StopBits = frame1_entry6.get()
    Address =frame1_entry7.get()

    command ="--connection " + Comm_Type +','+ COM_Port+','+Baud_Rate + Bits + Parity +StopBits +','+Address

    return command
#*************************************************************************************#
def deviceconnect():

    connection=extractdata();

    command = "sparkly device read-data " + connection
    os.system(command)
#*************************************************************************************#
def senddata(parameter,value):
    connection=extractdata()
    
    message=""
    message="{0}={1}".format(parameter,value)

    command="sparkly parameters write --parameter-list "+ " " + message + " " +connection
    os.system(command)
#*************************************************************************************#
def download():
    connection=extractdata()

    path=frame2_entry1.get()

    command='sparkly app download --src ' + '"' + 'flash.bin' + '"' + ' ' + connection + ' '+'--working-directory' + ' '+ '"' + path + '"'

    os.system(command)
#*************************************************************************************#
def readdata(parameter):
    connection=extractdata()


    command="sparkly parameters read --parameter-list "+ " " + parameter + " " +connection+ " " +"--out-file"+ " " + "out.txt"
    os.system(command)
#*************************************************************************************#

def readTxt():
    parameters={}
    try:
        fileOBj=open('out.txt','r')

        for line in fileOBj:
            columns=line.split(";")
            pattern='[\n]'
            value=re.sub(pattern,"",columns[1])
            parameters[columns[0]]=value
    except:
        print("The File could not be read or not exists")

    fileOBj.close()
    print(parameters)

    return parameters

#************************************************************************************#

def EOLCDUConfig():
    """
         This Test check the status of CDU and turns it off
    """

    pass
        

def EOLEEV():
    """
         This Test check the status of EEV and closes the EEV Valve
    """
    pass

def EOLDrainHeaterTest():
    """
         This Test check the status of Drain Heater and turns it ON
    """
    pass
    

def EOLDefrostHeaterTest():
    """
         This Test check the status of Defrost Heater and turns it ON and effect on Coil Outlet Temp Probe
    """
    pass


def EOLLightTest():
    """
    This Test check the status of Light and turns it off/on based on status
    """
    readdata("Lht")
    parameters=readTxt()
    for key,value in parameters.items():
        if key=='Lht':
            if value=='TRUE':
                param=False
            elif value=='FALSE':
                param=True
            else:
                pass      
    print(param)
    senddata("Lht",param)


def EOLDefrostTermination():
    """
        Check for Doors and then Check for Change in EOL Defrost Termination Temp Probe
    """
    pass

def EOLFanTest():
    """
        Initialize the EOL Fan Test
    """
    readdata("Aux2")
    parameters=readTxt()
    for key,value in parameters.items():
        if key=='Aux2':
            if value=='1':
                param=2
            elif value=='2':
                param=1
            else:
                pass 
    senddata("Aux2",param)


 
def EOLBuzzerTest():
    """
    Check for Doors and then Check for Change in EOL Defrost Termination Temp Probe
    """

    pass

#*************************************************************************************#



def mpxone_run_test():
    """

    Initial Release : Rohit Maity
    Date : 21/08/2022

    Lists the Type of Tests to be performed  by the Program

    """
    EOLLightTest()

#****************************************************************************************#

root=Tk()
root.title("Mpx One Programming Tool")
# root.iconbitmap(r'logo.ico')

#****************************************************************************************#

frame1=LabelFrame(master=root,text="1.1 Communication Info")
frame1.grid(row=0,column=0)

frame1_label1=Label(master=frame1,text="COM Type :")
frame1_label1.grid(row=0,column=0)

frame1_entry1=Entry(master=frame1,width=40)
frame1_entry1.grid(row=0,column=1)
frame1_entry1.insert(0,"Serial")


frame1_label2=Label(master=frame1,text="COM Port :")
frame1_label2.grid(row=1,column=0)

frame1_entry2=Entry(master=frame1,width=40)
frame1_entry2.grid(row=1,column=1)
frame1_entry2.insert(0,"COM11")

frame1_label3=Label(master=frame1,text="Baud Rate :")
frame1_label3.grid(row=2,column=0)

frame1_entry3=Entry(master=frame1,width=40)
frame1_entry3.grid(row=2,column=1)
frame1_entry3.insert(0,"19200")

frame1_label4=Label(master=frame1,text="Bits :")
frame1_label4.grid(row=3,column=0)


frame1_entry4=Entry(master=frame1,width=40)
frame1_entry4.grid(row=3,column=1)
frame1_entry4.insert(0,"8")

frame1_label5=Label(master=frame1,text="Parity :")
frame1_label5.grid(row=4,column=0)

frame1_entry5=Entry(master=frame1,width=40)
frame1_entry5.grid(row=4,column=1)
frame1_entry5.insert(0,"N")

frame1_label6=Label(master=frame1,text="Stop Bits :")
frame1_label6.grid(row=5,column=0)

frame1_entry6=Entry(master=frame1,width=40)
frame1_entry6.grid(row=5,column=1)
frame1_entry6.insert(0,"2")

frame1_label7=Label(master=frame1,text="Device Address :")
frame1_label7.grid(row=6,column=0)

frame1_entry7=Entry(master=frame1,width=40)
frame1_entry7.grid(row=6,column=1)
frame1_entry7.insert(0,"199")

frame1_button1=Button(master=frame1,text="Test Connection",command=deviceconnect)
frame1_button1.grid(row=7,column=0)

#************************************************************************************* #

frame2=LabelFrame(master=root,text="1.2 Download To Controller")
frame2.grid(row=1,column=0)

frame2_label1=Label(master=frame2,text="Enter Path to Flash Bin :")
frame2_label1.grid(row=0,column=0)


frame2_entry1=Entry(master=frame2,width=40)
frame2_entry1.grid(row=0,column=1)
frame2_entry1.insert(0,"Enter Path to Project flash.bin file")

frame2_button1=Button(master=frame2,text="Download",command=download)
frame2_button1.grid(row=1,column=0)


# *********************************************************************************** #

frame3=LabelFrame(master=root,text="1.3 EOL Tests List")
frame3.grid(row=2,column=0)

frame3_button1=Button(master=frame3,text="Lights",command=EOLLightTest)
frame3_button1.grid(row=0,column=0,columnspan=2)
frame3_button1=Button(master=frame3,text="Fans",command=EOLFanTest)
frame3_button1.grid(row=0,column=3,columnspan=2)
frame3_button1=Button(master=frame3,text="AUX",command=EOLLightTest)
frame3_button1.grid(row=0,column=5,columnspan=2)
frame3_button1=Button(master=frame3,text="Defrost",command=EOLLightTest)
frame3_button1.grid(row=0,column=7,columnspan=2)

#*************************************************************************************#
# frame4=LabelFrame(master=root,text="1.4 Read Data From Controller")
# frame4.grid(row=3,column=0)

# frame4_label1=Label(master=frame4,text="Controller Data Abbreviatios :")
# frame4_label1.grid(row=0,column=0)

# frame4_entry1=Entry(master=frame4,width=40)
# frame4_entry1.grid(row=0,column=1)
# frame4_entry1.insert(0,"Enter Abbreviations Separated By Comma... Ex: H1,H2...")

# frame4_button1=Button(master=frame4,text="Read Data",command=EOLLightTest)
# frame4_button1.grid(row=1,column=0)


#*************************************************************************************#

root.mainloop()