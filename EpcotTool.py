# -*- coding: utf-8 -*-
"""
Created on Mon Feb  3 13:50:17 2020
@author: JOSHUA Meduoye
Version 2.2 created on Dec 2 2020
@author: Wendy Dinch
"""
#Running a windows UI
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import sys
import glob
import serial
from pymodbus.client.sync import ModbusSerialClient
import time

import requests

import json
import jsonpickle



import pandas as pd
#root.tk.call('wm', 'iconphoto', root._w, PhotoImage(file='logo-dover_main.png'))


# time on computer
import datetime
time = datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S") # format using the function as required


info = None

'''
This part of code connects to google API to get information and write information
'''
#library for connecting to google cloud
import gspread
from oauth2client.service_account import ServiceAccountCredentials

#used to get username
import getpass



#Use this to log information about the running of the program
import logging
import time
#logging.basicConfig(filename = "DFR_ToolName_{}.log".format(str(time)), level = logging.INFO)


'''
This function can be used to log events of any program
'''
'''
def my_logger(original_function):
    
    def wrapper(*args, **kwargs):
        logging.info(f"{original_function.__name__}: Var: {args}, {kwargs}" )
        
        return original_function(*args, **kwargs)
    return wrapper
'''



'''
This part of code defines the google interface and reads json information that contains authentication
'''
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("cred.json", scope) ### return a credential from the key file.
client = gspread.authorize(creds)
file = client.open("Solochill Troubleshooting")
sheet = file.get_worksheet(0)
sheet2= file.get_worksheet(1)
sheet3= file.get_worksheet(2)
sheet4= file.get_worksheet(3)



setpoint= None
def_term_time= None
def_term_temp = None
def_interval = None
fan_mngmnt = None
fan_start = None
fan_per_def = None




#gets a list of all first column
column_orig = sheet.col_values(1)

'''
Change Everything to Lower Case i want to make sure upper and lower case are found
'''
column= [vals.lower() for vals in column_orig]


case_type_dt= None


'''
Main class defines GUI
'''
class Main_Window():
    def __init__(self, size="650x570"):
        #Current Version, firmware is used in the solochill application
        #This is an initialization for the class 
        self.curr_sw_version = None
        self.curr_firmware = None
        
        
        self.type = 0
        
        #This xl_data object is used to store data from the xl spreadsheet that is 
        #used to write parameters to the controller on the solochill side
        self.xl_data= None
        
        #This connected object is used to store the status of the modbus connection
        self.connected =  False
        self.curr_info = None
        
        #this is used for solochill to write the status of each write
        self.err =[None, None, None, None, None, None, None]
        
        #this is used for dollartree to write the status of each write
        self.err_dt =[None]

        self.sw_major= None
        self.sw_minor= None
        self.sw_prog= None
        self.sw_bios_high= None
        self.sw_bios_low = None
        self.setpoint=None
        self.def_term_time= None
        self.def_term_temp = None
        self.def_interval = None
        self.fan_mngmnt = None
        self.fan_start = None
        self.fan_per_def = None
        self.client2= None

        self.currhh = int(datetime.datetime.now().strftime("%H"))
        self.currmm = int(datetime.datetime.now().strftime("%M"))


        self.dt_data = {"Serial": None,
                        "UserName":None, 
                        "Date": None, 
                        "Case Type": None,
                        "Temp Unit": None,
                        "Decimal pt": None,
                        "Control Diff": None,
                        "Minimum SetP": None,
                        "Maximum SetP": None,
                        "Temp Setpoint": None,
                        "Comp Min Start Time": None,
                        "Defrost Interval": None, 
                        "Defrost Term Temp": None,
                        "Max Defrost Time": None,
                        "Drip Time": None,
                        "Alarm Type": None, 
                        "Low Temp ALarm": None,
                        "High Temp Alarm": None,
                        "Fan Management": None,
                        "Fan Start Temp": None,
                        "Fan Cycle": None, 
                        "Defrost Day": None, 
                        "Defrost Hour": None,
                        "Defrost Display": None,
                        "Password": None,
                        "Time Hour": None,
                        "Time Minute": None
                        


                        
                                                                    }

        self.dt_addr = {"Temp Unit": [0x27,'coil'],
                        "Decimal pt": [0x28,'coil'],
                        "Control Diff": [0x10,'hr'],
                        "Minimum SetP": [0x11,'hr'],
                        "Maximum SetP": [0x12,'hr'],
                        "Temp Setpoint": [0x0F,'hr'],
                        "Comp Min Start Time":[0x7F,'hr'],
                        "Defrost Interval": [0x89,'hr'], 
                        "Defrost Term Temp": [0x16,'hr'],
                        "Max Defrost Time": [0x8A,'hr'],
                        "Drip Time": [0x8F,'hr'],
                        "Alarm Type": [0x2F,'coil'], 
                        "Low Temp ALarm": [0x1A,'hr'],
                        "High Temp Alarm": [0x1B,'hr'],
                        "Fan Management": [0x9D,'hr'],
                        "Fan Start Temp": [0x1E,'hr'],
                        "Fan Cycle": [0x32,'coil'], 
                        "Defrost Day": [0xCD,'hr'], 
                        "Defrost Hour": [0xCE,'hr'],
                        "Defrost Display": [0x8E, 'hr'],
                        "Password": [0x71, 'hr'],
                        "Time Hour": [0x68, 'hr'],
                        "Time Minute": [0x69, 'hr']
            
                }
        
        

        #Data structure here is Key: [1 door med temp, 2-5 door Med Temp, Low Temp]

        self.dt_vals = {"Temp Unit": [True,True,True],
                        "Decimal pt": [True,True,True],
                        "Control Diff": [60,60,40],
                        "Minimum SetP": [250,250,65286],
                        "Maximum SetP": [800,800,800], 
                        "Temp Setpoint": [360,300,65446],
                        "Comp Min Start Time":[2,2,2],
                        "Defrost Interval": [12,12,24], 
                        "Defrost Term Temp": [400,400,480],
                        "Max Defrost Time": [46,46,46],
                        "Drip Time": [0,0,0],
                        "Alarm Type": [False, False, False], 
                        "Low Temp ALarm": [60,60,60],
                        "High Temp Alarm": [300,300,690],
                        "Fan Management": [0,0,2],
                        "Fan Start Temp": [50,50,300],
                        "Fan Cycle": [False, False, False], 
                        "Defrost Day": [11,11,11], 
                        "Defrost Hour": [6,6,6],
                        "Defrost Display": [0,0,0],
                        "Password": [138,138,138],
                        "Time Hour": [self.currhh,self.currhh,self.currhh],
                        "Time Minute": [self.currmm,self.currmm,self.currmm]
                       }     
        
 
        
        self.completed= None
        
        self.bsize, self.lsize = size.split('x')
        self.main = Tk()
        self.main.iconbitmap('logo.ico')
        self.main.title("Epcot 2.2")
        
        self.main.geometry(size)
        self.program = None
        self.case_port1_menu = None
        self.case_text1 = None
        self.status= None
        '''
        Setting up foundation
        '''
        
        self.foundation = ttk.Notebook(self.main)
        self.foundation.grid(row=0, column=1, columnspan =int(self.bsize), rowspan=int(self.lsize), sticky='NESW')
        self.lay()

        '''
        Sets first tab of window
        '''
        self.Window( "SoloChill","Troubleshooting", "Get Alarm Codes", "Send Feedback", "Dollar Tree")
        
        
        
        
        
        '''
        Runs the entire mainloop of program
        '''
        self.main.mainloop()

    def rest(self):
        time.sleep(0.1)
        
    def connect_device(self):
        self.connected =  False
        #client2.inter_char_timeout = 0.07
        self.client2 = ModbusSerialClient(method='rtu', port=self.case_port.cget('text'), baudrate=19200, timeout=5, stopbit =2, bytesize=8, parity= 'N')
        self.connected =  self.client2.connect()
        self.sw_major= None
        self.sw_minor= None
        self.sw_prog= None
        self.sw_bios_high= None
        self.sw_bios_low = None
        self.setpoint= None
        self.def_term_time= None
        self.def_term_temp = None
        self.def_interval = None
        self.fan_mngmnt = None
        self.fan_start = None
        self.fan_per_def = None
        return self.connected
    
    def connect_device_dt(self):
        self.connected =  False
        #client2.inter_char_timeout = 0.07
        self.client2 = ModbusSerialClient(method='rtu', port=self.case_port_dt.cget('text'), baudrate=19200, timeout=5, startbit=1, stopbit =2, bytesize=8, parity= 'N')
        self.connected =  self.client2.connect()

        return self.connected
    

    def get_serial_ports(self):
        """ 
            Lists serial port names

            :raises EnvironmentError:
                On unsupported or unknown platforms
            :returns:
                A list of the serial ports available on the system
        """
        if sys.platform.startswith('win'):
            ports = ['COM%s' % (i + 1) for i in range(256)]
        elif sys.platform.startswith('linux') or sys.platform.startswith('cygwin'):
            # this excludes your current terminal "/dev/tty"
            ports = glob.glob('/dev/tty[A-Za-z]*')
        elif sys.platform.startswith('darwin'):
            ports = glob.glob('/dev/tty.*')
        else:
            raise EnvironmentError('Unsupported platform')

        result = []
        for port in ports:
            try:
                s = serial.Serial(port)
                s.close()
                result.append(port)
            except (OSError, serial.SerialException):
                pass
        if len(result)>=1:
            return result
        else:
            return ["----"]

            
        
    
    def connection(self):
        if self.case_text1.get() == "":
            messagebox.showinfo('Epcot 2.2', "Pls put in the device identifier")
        else:
            self.rest()
            self.connect.state(["disabled"])
            self.status.configure(text="Connecting...")
            
            self.status.configure(background = 'yellow')
            self.rest()
            self.curr_sw_version = None
            self.curr_firmware = None
            
        
            self.connected =  False
            self.curr_info = None
            self.rest()
            self.connect.state(["disabled"])
            

                
                
            for vals in range(5):
                if self.connected==False:
                    time.sleep(3)
                    self.connected= self.connect_device()
                
                
                
        
            try:
                    
                    
                time.sleep(1)
                self.sw_major= self.client2.read_input_registers(0x1431,count=1, unit=194)
                
                
                time.sleep(1)
                self.sw_minor = self.client2.read_input_registers(0x1432,count=1, unit=194)
                
                
                time.sleep(1)
                self.sw_prog = self.client2.read_input_registers(0x1433,count=1, unit=194)
                
                
                time.sleep(1)
                self.sw_bios_high= self.client2.read_holding_registers(0x14CF, count=1, unit=194)
                
                
                
                time.sleep(1)
                self.sw_bios_low = self.client2.read_input_registers(0x14D0,count=1, unit=194)
                
                time.sleep(1)
                self.curr_sw_version = str(self.sw_major.registers[0]) +"." + str(self.sw_minor.registers[0]) + "." + str(self.sw_prog.registers[0])
                self.curr_firmware = str(self.sw_bios_high.registers[0]) +"." + str(self.sw_bios_low.registers[0])
                
            except:
                self.status.configure(text="Not Connected Try Again")
                self.status.configure(background = 'red')
                self.status.configure(foreground = "white")
                self.connect.state(["!disabled"])
                self.client2.close()
            else:
                
                self.curr_info = "Connected: Current SW Version: " + self.curr_sw_version +"  " + "Current Firmware Version: " +self.curr_firmware
                self.status.configure(text=self.curr_info )
                self.connect.state(["disabled"])
                self.program.state(["!disabled"])
                self.status.configure(background = 'green')
                self.status.configure(foreground = 'white')
                
        
    def connection_dt(self):
            if self.case_text1_dt.get() == "":
                messagebox.showinfo('Epcot 2.2', "Pls Enter Serial Number")
            else:
                self.rest()
                self.connect_dt.state(["disabled"])
                self.status_dt.configure(text="Connecting...")
                
                self.status_dt.configure(background = 'yellow')
                self.rest()
                self.curr_sw_version = None
                self.curr_firmware = None
                
            
                self.connected =  False
                self.curr_info = None
                self.rest()
                self.connect_dt.state(["disabled"])
                
    
                    
                    
                for vals in range(5):
                    if self.connected==False:
                        time.sleep(3)
                        self.connected= self.connect_device_dt()
                    
                    
                    
            
                try:
                    time.sleep(0.2)
                    self.dt_data["Temp Setpoint"] = self.client2.read_holding_registers(self.dt_addr["Temp Setpoint"][0], count= 1, unit=1)
                            
                    self.dt_data["Temp Setpoint"] = str(self.dt_data["Temp Setpoint"].registers[0])
                    
                    

                            
                            
                       

                        
                        
                        

                    
                except:
                    self.status_dt.configure(text="Not Connected Try Again")
                    self.status_dt.configure(background = 'red')
                    self.status_dt.configure(foreground = "white")
                    self.connect_dt.state(["!disabled"])
                    self.client2.close()
                else:
                    
                    self.curr_info = "Connected.. Ready to be Programmed"
                    self.status_dt.configure(text=self.curr_info )
                    self.connect_dt.state(["disabled"])
                    self.program_dt.state(["!disabled"])
                    self.status_dt.configure(background = 'green')
                    self.status_dt.configure(foreground = 'white')
                    
        
        
        
    def programexec(self):
        
        
        try:
            
            self.xl_data= self.getcaseinfo2()
            self.xl_data.set_index("Case+product", inplace=True)
            

            
            self.setpoint= int(self.xl_data.loc[self.case_menu.cget('text')][2])
            self.def_term_time= int(self.xl_data.loc[self.case_menu.cget('text')][3])
            self.def_term_temp = int(self.xl_data.loc[self.case_menu.cget('text')][4])
            self.def_interval = int(self.xl_data.loc[self.case_menu.cget('text')][5])
            self.fan_mngmnt = int(self.xl_data.loc[self.case_menu.cget('text')][6])
            self.fan_start = int(self.xl_data.loc[self.case_menu.cget('text')][7])
            self.fan_per_def = int(self.xl_data.loc[self.case_menu.cget('text')][8])

            
            
            
            
            self.program.state(["disabled"])
            self.status.configure(text="Programming In Progress")
            self.status.configure(background = "yellow")
            self.status.configure(background = "black")
            
            
            

            time.sleep(1)
            
            self.err[0]= self.client2.write_register(0xc2, self.setpoint, unit=194)
            for vals in range(3):
                if self.err[0].isError()==True:
                    self.err[0]= self.client2.write_register(0xc2, self.setpoint, unit=194)
                
            
            time.sleep(1)
            self.err[1] = self.client2.write_register(0x13e6, self.def_term_time, unit=194)
            for vals in range(3):
                if self.err[1].isError()==True:
                    self.err[1] = self.client2.write_register(0x13e6, self.def_term_time, unit=194)
                
            
            time.sleep(1)
            self.err[2] = self.client2.write_register(0x75, self.def_term_temp, unit=194)
            for vals in range(5):
                if self.err[2].isError()==True:
                    self.err[2] = self.client2.write_register(0x75, self.def_term_temp, unit=194)
                    
            
            time.sleep(1)
            self.err[3] = self.client2.write_register(0x1415, self.def_interval, unit=194)
            for vals in range(5):
                if self.err[3].isError()==True:
                    self.err[3] = self.client2.write_register(0x1415, self.def_interval, unit=194)
            
            time.sleep(1)
            self.err[4] = self.client2.write_register(0x13c5, self.fan_mngmnt, unit=194)
            for vals in range(5):
                if self.err[4].isError()==True:
                    self.err[4] = self.client2.write_register(0x13c5, self.fan_mngmnt, unit=194)
            
            time.sleep(1)
            self.err[5] = self.client2.write_register(0xa1, self.fan_start, unit=194)
            for vals in range(5):
                if self.err[5].isError()==True:
                    self.err[5] = self.client2.write_register(0xa1, self.fan_start, unit=194)            
            
            time.sleep(1)
            self.err[6] = self.client2.write_register(0x13c9, self.fan_per_def, unit=194)
            for vals in range(5):
                if self.err[6].isError()==True:
                    self.err[6] = self.client2.write_register(0x13c9, self.fan_per_def, unit=194)
            

            for vals in self.err:
                if vals.isError()==True:
                    vals[300] = True
            
        except:
            self.status.configure(text="Programming UnSuccessful Try Again")
            self.program.state(["!disabled"])
            self.status.configure(background = 'red')
            self.status.configure(foreground = 'white')
            
        else:
            self.completed = "Programming Completed Succesfully!    "+ "St="+	str(self.err[0].value) +" dp1="+str(self.err[1].value)+" dt1="+str(self.err[2].value)+" dl="+str(self.err[3].value) +	" F0="+str(self.err[4].value)	+" F1="+ str(self.err[5].value)+	" F3="+ str(self.err[6].value)

            self.status.configure(text=self.completed)
            self.status.configure(background = "green")
            self.program.state(["disabled"])
            self.connect.state(["!disabled"])
            self.status.configure(foreground = 'white')
            
            self.client2.close()
            sheet3.insert_row([self.case_text1.get(),self.curr_sw_version, self.curr_firmware,self.setpoint, self.def_term_time, self.def_term_temp,self.def_interval, self.fan_mngmnt, self.fan_start,self.fan_per_def,getpass.getuser(), datetime.datetime.now().strftime("%Y/%m/%d_%H:%M:%S")], index=2 )
        finally:
            "Data struct: Test/Final, case_identifier, sw_major, sw_minor, sw_prog, sw_bios_high, sw_bios_low, setpoint, Setpoint (°F)	Def Term Time (min)	Def Term Temp (°F) 	Def Interval (hour)	Fan Mngmt	Fan Start (°F) 	Fan per Defrost"

            #sheet3.insert_row([self.serial_text1.get(),self.Code_text1.get(), getpass.getuser(),datetime.datetime.now().strftime("%Y/%m/%d_%H:%M:%S"),self.Code_text1.get().lower() in column], index=2 )
        

    def programexec_dt(self):

        
        self.currhh = int(datetime.datetime.now().strftime("%H"))
        self.currmm = int(datetime.datetime.now().strftime("%M"))
        self.dt_vals["Time Hour"] = [self.currhh,self.currhh,self.currhh]
        self.dt_vals["Time Minute"] = [self.currmm,self.currmm,self.currmm]

    
        try:

            #if self.case_text1_dt.get() == self.dt_data["Serial"]:
               #messagebox.showinfo('Epcot 2.2',"Pls put in a new serial number")

            if  self.case_text2_dt.get()=="JNRBHSA-1":
                index =0
                self.case_type_dt = "1 Door MT"
            elif  self.case_text2_dt.get()=="JNRBHSA-2":
                index =1
                self.case_type_dt = "2-5 Door MT"
            elif self.case_text2_dt.get()=="JNRBHSA-3":
                index =1
                self.case_type_dt = "2-5 Door MT" 
            elif  self.case_text2_dt.get()=="JNRBHSA-4":
                index =1
                self.case_type_dt = "2-5 Door MT"
            elif self.case_text2_dt.get()=="JNRBHSA-5":
                index =1
                self.case_type_dt = "2-5 Door MT"  
            elif self.case_text2_dt.get()=="JNRZHSA-2":
                index =2
                self.case_type_dt = "Low Temp"
            elif self.case_text2_dt.get()=="JNRZHSA-3":
                index =2
                self.case_type_dt = "Low Temp"  
            elif self.case_text2_dt.get()=="JNRZHSA-4":
                index =2
                self.case_type_dt = "Low Temp"  
            elif self.case_text2_dt.get()=="JNRZHSA-5":
                index =2
                self.case_type_dt = "Low Temp"   
            else:
               messagebox.showinfo('Epcot 2.2',"Case Type does not exist")
               int('c')
               
               
            self.dt_data["Serial"] = self.case_text1_dt.get()
            self.dt_data["UserName"] = getpass.getuser()
            self.dt_data["Date"] = datetime.datetime.now().strftime("%Y/%m/%d_%H:%M:%S")
            self.dt_data["Case Type"] = self.case_type_dt

         
            for vals in self.dt_addr.keys():
              
                if self.dt_addr[vals][1] =='hr':


                    self.err_dt[0] = self.client2.write_register(self.dt_addr[vals][0], self.dt_vals[vals][index], unit=1)

                    
                    for vals2 in range(5):
                        if self.err_dt[0].isError()==True:
                            time.sleep(0.2)
                            self.err_dt[0]= self.client2.write_register(self.dt_addr[vals][0], self.dt_vals[vals][index], unit=1)
                    self.dt_data[vals] = self.dt_vals[vals][index]
                    #time.sleep(0.2)

                  
                elif self.dt_addr[vals][1] =='coil':
                    
                    time.sleep(0.2)
                    self.err_dt[0] = self.client2.write_coil(self.dt_addr[vals][0], self.dt_vals[vals][index], unit =1)
                    for vals2 in range(5):
                        if self.err_dt[0].isError()==True:
                            time.sleep(0.2)
                            self.err_dt[0] = self.client2.write_coil(self.dt_addr[vals][0], self.dt_vals[vals][index], unit =1)
                    self.dt_data[vals] = self.dt_vals[vals][index]
 
            #self.dt_data[vals] = jsonpickle.encode(self.client2.read_holding_registers(self.dt_addr[vals][0], count= 1, unit=1).registers[0])                
        except:
            self.status_dt.configure(text="Programming UnSuccessful Try Again")
            self.program_dt.state(["!disabled"])
            self.status_dt.configure(background = 'red')
            self.status_dt.configure(foreground = 'white')
        else:
            self.completed = "Programming Completed Succesfully! " + "(Programmed as " + self.case_type_dt +") Serial Number:"+ str(self.dt_data["Serial"]) 
            self.status_dt.configure(text=self.completed)
            self.status_dt.configure(background = "green")
            self.program_dt.state(["disabled"])
            self.connect_dt.state(["!disabled"])
            self.status_dt.configure(foreground = 'white')
            self.case_text1_dt.delete("0",END)
            self.case_text2_dt.delete("0",END)

           
            sheet4.insert_row(
                    [
                            self.dt_data["Serial"]
                            ,self.dt_data["UserName"]
                            ,self.dt_data["Date"]
                            ,self.dt_data["Case Type"]
                            ,self.dt_vals["Temp Unit"][index]
                            ,self.dt_vals["Decimal pt"][index]
                            ,self.dt_vals["Temp Setpoint"][index]
                            ,self.dt_vals["Control Diff"][index]
                            ,self.dt_vals["Minimum SetP"][index]
                            ,self.dt_vals["Maximum SetP"][index]
                            ,self.dt_vals["Comp Min Start Time"][index]
                            ,self.dt_vals["Defrost Interval"][index]
                            ,self.dt_vals["Defrost Term Temp"][index]
                            ,self.dt_vals["Max Defrost Time"][index]
                            ,self.dt_vals["Drip Time"][index]
                            ,self.dt_vals["Alarm Type"][index]
                            ,self.dt_vals["Low Temp ALarm"][index]
                            ,self.dt_vals["High Temp Alarm"][index]
                            ,self.dt_vals["Fan Management"][index]
                            ,self.dt_vals["Fan Start Temp"][index]
                            ,self.dt_vals["Fan Cycle"][index]
                            ,self.dt_vals["Defrost Day"][index]
                            ,self.dt_vals["Defrost Hour"][index]
                            ,self.dt_vals["Defrost Display"][index]
                            ,self.dt_vals["Password"][index]
                                                                ]
                            , index=2 )
            self.postEOLTestData(self.dt_data)
            
            #for vals in self.dt_data.items():
                            #print (vals)
            
     
            
            self.client2.close()
            
            
        finally:
            #self.case_text1_dt.delete("0",END)
            "Data struct: Test/Final, case_identifier, sw_major, sw_minor, sw_prog, sw_bios_high, sw_bios_low, setpoint, Setpoint (°F)	Def Term Time (min)	Def Term Temp (°F) 	Def Interval (hour)	Fan Mngmt	Fan Start (°F) 	Fan per Defrost"

    
    #data upload for Joe
    
    def postEOLTestData(self, payload, verbose=False):
        url = "http://hpx-con-apps:8090/EolTesting/EolTesting/EolEndpoint"
        headers = {'content-type': 'application/json'}
        response = requests.request("POST", url, data=json.dumps(payload), headers=headers)
        #if verbose:
            #print(response.text)

            
    def getcaseinfo(self):
        """
        Gives the names of all information in the excel file available
        """
        info = None
        
        try:
            info = pd.read_excel('Final.xlsx')            
        except:
            try:
                info = pd.read_excel('Test.xlsx')
            except:
                info = pd.DataFrame({'Case+product': ["None Available"]})
                self.main.title("Epcot 2.2.. Err:No Data Loaded")
                #messagebox.showinfo('Epcot 2.2',"Unable to Load Case Types")
                
            else:
                self.main.title("Epcot 2.2 Run Test Version")
                #messagebox.showinfo('Epcot 2.2',"Loaded Run Test Application")
        else:
            self.main.title("Epcot 2.2 Final Customer Version")
            #messagebox.showinfo('Epcot 2.2',"Loaded Final Programming Application")
                

        return info

    def getcaseinfo2(self):
        """
        Gives the names of all information in the excel file available
        """
        info = None
        
        try:
            info = pd.read_excel('Final.xlsx')            
        except:
            try:
                info = pd.read_excel('Test.xlsx')
            except:
                info = pd.DataFrame({'Case+product': ["None Available"]})
                self.main.title("Epcot 2.2.. Err:No Data Loaded")
                #messagebox.showinfo('Epcot 2.2',"Unable to Load Case Types")
                
            else:
                self.main.title("Epcot 2.2 Run Test Version")
                #messagebox.showinfo('Epcot 2.2',"Loaded Run Test Application")
        else:
            self.main.title("Epcot 2.2 Final Customer Version")
            #messagebox.showinfo('Epcot 2.2',"Loaded Final Programming Application")
                

        return info
        
    def lay(self):
        
        for vals in range(max(int(self.bsize), int(self.lsize))):
            self.foundation.rowconfigure(min(vals, int(self.bsize)),weight=1)
            
    '''
    This creates a wibndow and everything in window 
    
    '''
    def Window(self, *args):
        style = ttk.Style()
        style.configure("BW.TLabel", foreground="grey", background="blue")
        '''
        New Richmond Tab 
        '''
        tab0 = ttk.Frame(self.foundation)
        self.foundation.add(tab0, text= args[0])
        
        stepOne = ttk.LabelFrame(tab0, text=" 1. Communication Info: ")
        stepOne.grid(row=3, columnspan=7, sticky='EW', padx=15, pady=5, ipadx=5, ipady=5)

                
        

        
        self.CaseCommPort1 = ttk.Label(stepOne, text="COM Port:", width=13)
        self.CaseCommPort1.grid(column= 1, row= 1, padx= (0,0), sticky ="W")

        
        self.case_port ={}
        self.port = StringVar(self.main)
        self.port.set(self.get_serial_ports()[0])
        self.case_port = OptionMenu(stepOne,self.port,*self.get_serial_ports())
        self.case_port.grid(column= 2, row= 1, sticky ="EW", padx= (0,0))
        
        
        
        stepTwo = ttk.LabelFrame(tab0, text=" 2. Product Info: ")
        stepTwo.grid(row=4, columnspan=7, sticky='EW', padx=15, pady=5, ipadx=5, ipady=5)
        
        self.case = ttk.Label(stepTwo, text="Case Identifier:", width=20)
        self.case.grid(column= 1, row= 1, padx= (0,0), sticky ="W")
        
        self.case_text1 = ttk.Entry(stepTwo, width=45)
        self.case_text1.grid(column= 2, row= 1, sticky ="E", padx= (90,90))
        
        self.case = ttk.Label(stepTwo, text="Case Type:", width=20)
        self.case.grid(column= 1, row= 3, padx= (0,0), sticky ="W")
        
        self.case ={}
        self.case[1] = StringVar(self.main)
        self.case[1].set("Select Case Type")
        self.case_menu = OptionMenu(stepTwo,self.case[1],*self.getcaseinfo()['Case+product'])
        self.case_menu.grid(column= 2, row= 3, sticky ="EW", padx= (0,0))
        
        stepThree = ttk.LabelFrame(tab0, text=" 3. Connection/Programming: ")
        stepThree.grid(row=6, columnspan=7, sticky='EW', padx=15, pady=5, ipadx=5, ipady=5)
        
       
        

        self.connect = ttk.Button(stepThree, text = "Connect", command = self.connection)
        self.connect.grid(column =1, row=1, sticky ="E", padx=90)
        
        self.program = ttk.Button(stepThree, text = "Program", command = self.programexec, state= 'disabled' )
        self.program.grid(column =2, row=1, sticky ="E", padx=30, pady =10)
        
        stepfour = ttk.LabelFrame(tab0, text=" 4. Status" )
        stepfour.grid(row=8, columnspan=7, sticky='EW', padx=15, pady=5, ipadx=5, ipady=5)
        
        self.status = ttk.Label(stepfour, text="Not Connected", width=80)
        self.status.grid(column= 1, row= 10, padx= (0,0), sticky ="W")
        
        '''
        
        Window 1 
        
        '''
        
        
        tab1 = ttk.Frame(self.foundation)
        self.foundation.add(tab1, text= args[1])
        '''
        Setup text box in new Tab
        '''
        self.labelserial = ttk.Label(tab1, text="Device Serial Number", width=37)
        self.labelserial.grid(column= 1, row= 1, padx= (100,50), sticky ="W")
        
        self.serial_text1 = ttk.Entry(tab1, width=25 )
        self.serial_text1.grid(column= 1, row= 1, sticky ="W", padx= (280,90))
        

        self.Code_label = ttk.Label(tab1, text="Pls Type in Alarm Code here -->", width=37)
        self.Code_label.grid(column= 1, row= 2, padx= (100,50), sticky ="W")
        
        
        self.Code_text1 = ttk.Entry(tab1, width=25)
        self.Code_text1.grid(column= 1, row= 2, sticky ="W", padx= (280,90))
        
        
        self.Space1_label = ttk.Label(tab1)
        self.Space1_label.grid(column= 1, row= 3, padx= (40,40), sticky ="W")
        
        
        self.Display_label = ttk.Label(tab1, text="Display description", width=27)
        self.Display_label.grid(column= 1, row= 4, padx= (40,40), sticky ="W")
        
        self.Display_desc = Text(tab1, height=3, width=70, state= DISABLED, relief="sunken", background='gray90', wrap=WORD)
        self.Display_desc.grid(column= 1, row= 5, padx= (40,40), sticky ="W")
        
        self.Space2_label = ttk.Label(tab1)
        self.Space2_label.grid(column= 1, row= 6, padx= (40,40), sticky ="W")
        
        self.Delay_label = ttk.Label(tab1, text="Delay", width=20)
        self.Delay_label.grid(column= 1, row= 7, padx= (40,40), sticky ="W")

        self.Delay_text = Text(tab1, height=1, width=20, state= DISABLED, relief="sunken", background='gray90')
        self.Delay_text.grid(column= 1, row= 8, sticky ="W", padx= (40,40))            
        
        self.Reset_label = ttk.Label(tab1, text="Reset", width=20)
        self.Reset_label.grid(column= 1, row= 7, padx= (300,40), sticky ="W")

        self.Reset_text = Text(tab1, height=1, width=20, state= DISABLED, relief="sunken", background='gray90')
        self.Reset_text.grid(column= 1, row= 8, sticky ="W", padx= (300,40))

        self.Space3_label = ttk.Label(tab1)
        self.Space3_label.grid(column= 1, row= 9, padx= (40,40), sticky ="W")  

        self.Action_label = ttk.Label(tab1, text="Action")
        self.Action_label.grid(column= 1, row= 10, padx= (40,40), sticky ="W")   

        self.Action_text = Text(tab1, height=5, width=70, state= DISABLED, relief="sunken", background='gray90', wrap=WORD)
        self.Action_text.grid(column= 1, row= 11, sticky ="W", padx= (40,40))

        self.Space3_label = ttk.Label(tab1)
        self.Space3_label.grid(column= 1, row= 12, padx= (40,40), sticky ="W") 

        self.trouble_label = ttk.Label(tab1, text="Troubleshooting")
        self.trouble_label.grid(column= 1, row= 14, padx= (40,40), sticky ="W") 
        
        '''
        Setup an information Only Screen
        '''
        self.troubleshooting = Text(tab1,height=8, width=70, state=DISABLED, borderwidth=2, relief="sunken", background='gray90', wrap=WORD)
        self.troubleshooting.grid(column=1,row=15, sticky ="W", padx= (40,40), pady= (40,40))

        '''
        Setup a button in new Tab
        '''
        
        self.button1 = ttk.Button(tab1, text= "Get Alarm Information", command =self.set_trouble_info)
        self.button1.grid(row=2, column =1, sticky ="W", padx= (500,90))
        

        
        '''
        
        Window 2
        '''

        
        tab2 = ttk.Frame(self.foundation)
        self.foundation.add(tab2, text= args[2])
        
        self.button1 = ttk.Button(tab2, text= "Get All Alarm Codes", command =self.get_alarm_codes)
        self.button1.grid(row=1, column =1)

        self.alarms = Text(tab2,height=25, width=70, state=DISABLED, borderwidth=2, relief="sunken", background='gray90')
        self.alarms.grid(column=1,row=2, sticky ="W", padx= (40,40))
        
        
        '''
        Window 3
        '''
        '''
        Nates feedback on resolution
        
        
        '''
        tab3 = ttk.Frame(self.foundation)
        self.foundation.add(tab3, text= args[3])
        
        self.Code_label2 = ttk.Label(tab3, text="Pls Type in Alarm Code here -->", width=37)
        self.Code_label2.grid(column= 1, row= 1, padx= (100,50), sticky ="W")
        
        
        self.Code_text2 = ttk.Entry(tab3, width=25)
        self.Code_text2.grid(column= 1, row= 1, sticky ="W", padx= (280,90))
                

        
        self.Resolution_label = ttk.Label(tab3, text="Actual Resolution")
        self.Resolution_label.grid(column= 1, row= 2, padx= (40,40), sticky ="W")  
        
        self.Resolution = Text(tab3,height=25, width=70, borderwidth=2, relief="sunken", wrap=WORD)
        self.Resolution.grid(column=1,row=2, sticky ="W", padx= (40,40))
        
        self.button2 = ttk.Button(tab3, text= "Send Feedback", command =self.send_feedback)
        self.button2.grid(row=3, column =1, sticky ="W", padx= (500,90))
        
        '''
        Dollar tree Implementation
        '''
        tab4 = ttk.Frame(self.foundation)
        self.foundation.add(tab4, text= args[4])
        
        stepOne_dt = ttk.LabelFrame(tab4, text=" 1. Communication Info: ")
        stepOne_dt.grid(row=3, columnspan=7, sticky='EW', padx=15, pady=5, ipadx=5, ipady=5)
        
        
                
        

        
        self.CaseCommPort1_dt = ttk.Label(stepOne_dt, text="COM Port:", width=13)
        self.CaseCommPort1_dt.grid(column= 1, row= 1, padx= (0,0), sticky ="W")
        
        
        
        self.case_port_dt ={}
        self.port_dt = StringVar(self.main)
        self.port_dt.set(self.get_serial_ports()[0])
        self.case_port_dt = OptionMenu(stepOne_dt,self.port_dt,*self.get_serial_ports())
        self.case_port_dt.grid(column= 2, row= 1, sticky ="EW", padx= (0,0))
        
        
        self.case_type_dt = "none"
        stepTwo_dt = ttk.LabelFrame(tab4, text=" 2. Product Info: ")
        stepTwo_dt.grid(row=4, columnspan=7, sticky='EW', padx=15, pady=5, ipadx=5, ipady=5)
        
        self.dt_serial_lb = ttk.Label(stepTwo_dt, text="Serial Number:", width=20)
        self.dt_serial_lb.grid(column= 1, row= 1, padx= (0,0), sticky ="W")
        
        self.case_text1_dt = ttk.Entry(stepTwo_dt, width=45)
        self.case_text1_dt.grid(column= 2, row= 1, sticky ="W", padx= (0,90))
   
        self.dt_case_lb = ttk.Label(stepTwo_dt, text="Case Type:", width=20)
        self.dt_case_lb.grid(column= 1, row= 4, padx= (0,0), sticky ="W")
        
        self.case_text2_dt = ttk.Entry(stepTwo_dt, width=45)
        self.case_text2_dt.grid(column= 2, row= 4, sticky ="W", padx= (0,90))
        
        stepThree_dt = ttk.LabelFrame(tab4, text=" 3. Connection/Programming: ")
        stepThree_dt.grid(row=6, columnspan=7, sticky='EW', padx=15, pady=5, ipadx=5, ipady=5)
        
       
        

        self.connect_dt = ttk.Button(stepThree_dt, text = "Connect", command = self.connection_dt)
        self.connect_dt.grid(column =1, row=1, sticky ="E", padx=90)
        
        self.program_dt = ttk.Button(stepThree_dt, text = "Program", command = self.programexec_dt, state= 'disabled' )
        self.program_dt.grid(column =2, row=1, sticky ="E", padx=30, pady =10)
        
        stepfour_dt = ttk.LabelFrame(tab4, text=" 4. Status" )
        stepfour_dt.grid(row=8, columnspan=7, sticky='EW', padx=15, pady=5, ipadx=5, ipady=5)
        
        self.status_dt = ttk.Label(stepfour_dt, text="Not Connected", width=80)
        self.status_dt.grid(column= 1, row= 10, padx= (0,0), sticky ="W")
        
        
    def get_file_location(self):
        self.text1.delete(1.0,END)
        self.text1.configure(state =NORMAL)
        self.text1.insert(0.0,filedialog.askopenfilename())
        
        
    def set_trouble_info(self):
        
        if len(self.Code_text1.get())<1 or len(self.serial_text1.get())<1:


                
            self.Delay_text.configure(state=NORMAL)    
            self.Delay_text.delete(1.0,END)
            self.Delay_text.configure(state=DISABLED)
                
            
            self.Reset_text.configure(state=NORMAL)
            self.Reset_text.delete(1.0,END)
            self.Reset_text.configure(state=DISABLED)
            
                
            
            self.Action_text.configure(state=NORMAL)
            self.Action_text.delete(1.0,END)
            self.Action_text.configure(state=DISABLED)
                
            self.troubleshooting.configure(state=NORMAL)    
            self.troubleshooting.delete(1.0,END)
            self.troubleshooting.configure(state=DISABLED)
            
            self.Resolution.delete(1.0,END)
            self.Display_desc.configure(state=NORMAL)
            self.Display_desc.delete(1.0,END)
            
            messagebox.showinfo('Epcot 2.2',"Pls put in Device Serial Number and Alarm Code")
            #self.Display_desc.insert(0.0,"Pls put in Alarm Code")
            self.Display_desc.configure(state=DISABLED)
        
        else:
                
            
            try:
                
                self.colnum= column.index(self.Code_text1.get().lower())
            except:
                self.Display_desc.configure(state=NORMAL)
                self.Display_desc.delete(1.0,END)
                
                messagebox.showinfo('Epcot 2.2', "Alarm Does not Exist, or Is still in development")
                #self.Display_desc.insert(0.0,"Alarm Does not Exist, or Is still in development")
                self.Display_desc.configure(state=DISABLED)
                
                
                self.Delay_text.configure(state=NORMAL)    
                self.Delay_text.delete(1.0,END)
                self.Delay_text.configure(state=DISABLED)
                    
                
                self.Reset_text.configure(state=NORMAL)
                self.Reset_text.delete(1.0,END)
                self.Reset_text.configure(state=DISABLED)
                
                    
                
                self.Action_text.configure(state=NORMAL)
                self.Action_text.delete(1.0,END)
                self.Action_text.configure(state=DISABLED)
                    
                self.troubleshooting.configure(state=NORMAL)    
                self.troubleshooting.delete(1.0,END)
                self.troubleshooting.configure(state=DISABLED)

            
            else:
                            
    
                    
            
            
                if self.Code_text1.get().lower() in column:
                
            
                    self.Display_desc.configure(state=NORMAL)
                    self.Display_desc.delete(1.0,END)
                    
                    self.Display_desc.insert(0.0,sheet.row_values(self.colnum+1)[1])
                    self.Display_desc.configure(state=DISABLED)
                    
                    
                    self.Delay_text.configure(state=NORMAL)
                    self.Delay_text.delete(1.0,END)
                    
                    self.Delay_text.insert(0.0,sheet.row_values(self.colnum+1)[2])
                    self.Delay_text.configure(state=DISABLED)
                    
                    self.Reset_text.configure(state=NORMAL)
                    self.Reset_text.delete(1.0,END)
                    
                    self.Reset_text.insert(0.0,sheet.row_values(self.colnum+1)[3])
                    self.Reset_text.configure(state=DISABLED)
                    
                    self.Action_text.configure(state=NORMAL)
                    self.Action_text.delete(1.0,END)
                    
                    self.Action_text.insert(0.0,sheet.row_values(self.colnum+1)[4])
                    self.Action_text.configure(state=DISABLED)
                    
                    
                    self.troubleshooting.configure(state=NORMAL)
                    self.troubleshooting.delete(1.0,END)
                    self.troubleshooting.insert(0.0,sheet.row_values(self.colnum+1)[5])
                    self.troubleshooting.configure(state=DISABLED)
                    
    
                    
                    
            finally:
                
                    
                sheet2.insert_row([self.serial_text1.get(),self.Code_text1.get(), getpass.getuser(),datetime.datetime.now().strftime("%Y/%m/%d_%H:%M:%S"),self.Code_text1.get().lower() in column], index=2 )


    def send_feedback(self):
        sheet2.insert_row(["", self.Code_text2.get(), getpass.getuser(),datetime.datetime.now().strftime("%Y/%m/%d_%H:%M:%S"),self.Code_text2.get().lower() in column, self.Resolution.get("1.0","end-1c")], index=2 )
        self.Code_text2.delete(first=0, last= END)
        self.Resolution.delete(1.0,END)
        
        
    '''
    function takes in return key and does absolutely nothing with it
    '''



        
    def get_alarm_codes(self):
            
        self.alarms.delete(1.0,END)
        self.alarms.configure(state=NORMAL)
        self.alarms.insert(0.0,column_orig)
            
        
Run_program = Main_Window()


        