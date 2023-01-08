from tkinter import *
from tkinter import ttk


def modbusconnect():
    pass

def iothubconnect():
    pass

def showdatamodbus():
    pass

def senddatacloud():
    pass


root=Tk()
#section 1 - Connection Area


section1=LabelFrame(master=root,text="1: Modbus Connection Information")
section1.grid(row=0,column=0)

section1_label1=Label(master=section1,text="Modbus TCP IP :")
section1_label1.grid(row=0,column=0)

section1_entry1=Entry(master=section1,width=20)
section1_entry1.grid(row=0,column=1)

section1_label2=Label(master=section1,text="Port :",justify="left")
section1_label2.grid(row=0,column=2)

section1_entry2=Entry(master=section1,width=20)
section1_entry2.grid(row=0,column=3)




#section 2 - Selection Area

section2=LabelFrame(master=root,text="2: Select Read/Write Registers/Coils")
section2.grid(row=1,column=0)

section2_label1=Label(master=section2,text="Select Function :")
section2_label1.grid(row=0,column=0)

section2_dropdown1=ttk.Combobox(master=section2,width=50)
section2_dropdown1.grid(row=0,column=1)
section2_dropdown1['values']=("Holding Registers","Digital Inputs","None")

section2_label2=Label(master=section2,text="Select Start Address :")
section2_label2.grid(row=0,column=2)

section2_entry1=Entry(master=section2,width=50)
section2_entry1.grid(row=0,column=3)
section2_entry1.insert(END,"Start from 0")

section2_label3=Label(master=section2,text="Select Number of Registers to Read :")
section2_label3.grid(row=0,column=4)

section2_entry2=Entry(master=section2,width=50)
section2_entry2.grid(row=0,column=5)
section2_entry2.insert(END,"Value greater than 0")

#section 3 - Status Area
section3=LabelFrame(master=root,text="3: Status Information")
section3.grid(row=2,column=0)

section3_textarea1=Text(master=section3,height=2)
section3_textarea1.grid(row=1,column=0)

section3_button1=Button(master=section3,text="Connect",command=modbusconnect)
section3_button1.grid(row=1,column=1)

#section 4 - Result Area
section4=LabelFrame(master=root,text="4: Data ")
section4.grid(row=3,column=0)

root.mainloop()




