"""
    Initial Release : Rohit Maity
    Date : 21/08/2022
"""
from pymodbus.client.sync import ModbusSerialClient as ModbusClient 
import time

# MpxOne Modbus Address List

def tON(secs):
    """
    Initial Release : Rohit Maity
    Date : 01/09/2022

    Timer to Delay Code for specified seconds
    Cannot Have negative or zero values
    Return True after Timer Complete else False 
    """
    value=float(secs)
    if(value<=0.0):
        print("The Timer cannot have zero seconds or negative values...")
        return False
    else:
        try:
            time.sleep(value)
            return True
        except:
            print("Error......")
            return False

class mpxoneserialconnection:
    """

    Initial Release : Rohit Maity
    Date : 21/08/2022

    Creates the serial connection to the mpxone basic controller for read and write registers
    
    """

    def __init__(self,method,port,stopbits,bytesize,parity,baudrate,deviceid):
        """

        Initial Release : Rohit Maity
        Date : 21/08/2022

        client = ModbusClient(method = 'rtu', port='PORT10', stopbits = 2, bytesize = 8, parity = 'N' , baudrate= 38400)

        The methods to connect are::

          - ascii
          - rtu
          - binary

        :param method: The method to use for connection
        :param port: The serial port to attach to
        :param stopbits: The number of stop bits to use
        :param bytesize: The bytesize of the serial messages
        :param parity: Which kind of parity to use
        :param baudrate: The baud rate to use for the serial device

        """

        global gmethod
        global gport 
        global gstopbits 
        global gbytesize
        global gparity
        global gbaudrate
        global gdeviceid
        global connection

        self.method=method
        self.port=port
        self.stopbits=stopbits
        self.bytesize=bytesize
        self.parity=parity
        self.baudrate=baudrate
        self.deviceid=deviceid
        connection=ModbusClient(method = self.method, port=self.port, stopbits = self.stopbits, bytesize = self.bytesize, parity = self.parity , baudrate= self.baudrate)
        print("Object Created.....")

        # client = ModbusClient(method = 'rtu', port='PORT10', stopbits = 2, bytesize = 8, parity = 'N' , baudrate= 38400)
        connection.connect()

        if(connection.connect()):
            print("Connection Succesfull........")
            gmethod=self.method
            gport=self.port
            gstopbits=self.stopbits
            gbytesize=self.bytesize
            gparity=self.parity
            gbaudrate=self.baudrate
            gdeviceid=self.deviceid
            self.__read_holding_registers()
            self.__read_input_registers()
            self.__read_coils()
        else:
            print("Bad Connection.....")

    def __read_holding_registers(self):
        """

        Initial Release : Rohit Maity
        Date : 21/08/2022

        coil  = client.read_holding_registers(0x01,1,unit=1) address, count, slave address
    
        """
        global holdingdata1
        global holdingdata2
        global holdingdata3
        global holdingdata4
        global connection
        try:
            # coil  = client.read_holding_registers(0x01,1,unit=1) address, count, slave address
            var=connection.read_holding_registers(0x00,100,unit=self.deviceid)
            print("Read Holding Registers Success........")
            holdingdata1=var.registers
            var=connection.read_holding_registers(0x101,100,unit=self.deviceid)
            holdingdata2=var.registers
            var=connection.read_holding_registers(0x201,100,unit=self.deviceid)
            holdingdata3=var.registers
            var=connection.read_holding_registers(0x301,100,unit=self.deviceid)
            holdingdata4=var.registers

            #print(f"The value of Set 1 is {holdingdata1} , the value of set 2 is {holdingdata2}, the value of set 3 is {holdingdata3} and the last set is {holdingdata4}")
            
        except:
            print("Could not establish connection to the holding registers..........")
    
    def __read_input_registers(self):
        """

        Initial Release : Rohit Maity
        Date : 21/08/2022

        coil  = client.read_holding_registers(0x01,1,unit=1) address, count, slave address
    
        """
        global inputdata
        global connection
        try:
            # coil  = client.read_holding_registers(0x01,1,unit=1) address, count, slave address
            var=connection.read_input_registers(0x00,100,unit=self.deviceid)
            print("Read Input Registers Success........")
            inputdata=var.registers
            #print(inputdata)
                
        except:
               print("Could not establish connection to the holding registers......")

    def __read_coils(self):
        """

        Initial Release : Rohit Maity
        Date : 21/08/2022

        coil  = client.read_holding_registers(0x01,1,unit=1) address, count, slave address
    
        """
        global inputcoils
        global connection
        try:
            # coil  = client.read_coils(0x01,1,unit=1) address, count, slave address
            var=connection.read_coils(0x00,100,unit=self.deviceid)
            print("Read Input Coils Success...........")
            inputcoils=var.bits
            #print(inputcoils)
                
        except:
               print("Could not establish connection to the input coils.........")

    def refreshData(self):
        self.__read_coils()
        self.__read_holding_registers()
        self.__read_input_registers()

    # def displaydata(self):
    #     print(f"The value of the input registers from address 0-100 are {inputdata} \n")
    #     print(f"The value of the holding registers from address 0-100 registers are {holdingdata1} \n")
    #     print(f"The value of the holding registers from address 100-200 registers are {holdingdata2} \n")
    #     print(f"The value of the holding registers from address 200-300 registers are {holdingdata3} \n")
    #     print(f"The value of the holding registers from address 300-400 registers are {holdingdata4} \n")
    #     print(f"The value of the coil from  address 0-100 are {inputcoils} \n")


    def inputcoils(self):
        self.__read_coils()
        var={}
        for i in range(100):
            key=f"0x{i}"
            var[key]=inputcoils[i]

        # print(var)
        return var

    def inputregisters(self):
        self.__read_input_registers()
        var={}
        for i in range(100):
            key=f"0x{i}"
            var[key]=inputdata[i]

        # print(var)
        return var

    def holdingregisters(self):
        self.__read_holding_registers()
        var={}
        for i in range(100):
            key=f"0x{i}"
            var[key]=holdingdata1[i]
        for i in range(100):
            key=f"0x{i+100}"
            var[key]=holdingdata2[i]
        for i in range(100):
            key=f"0x{i+200}"
            var[key]=holdingdata3[i]
        for i in range(100):
            key=f"0x{i+300}"
            var[key]=holdingdata4[i]
        #print(var)
        return var


class mpxoneTests():
    """

    Initial Release : Rohit Maity
    Date : 21/08/2022

    List the individual tests for the mpxone controller
    
    """
    
    @staticmethod
    def EOLCDUConfig():
        """
         This Test check the status of CDU and turns it off
        """
        global inputcoils
        global connection
        global gdeviceid

        connection.write_coil(61, False, unit = gdeviceid)
        
    @staticmethod
    def EOLEEV():
        """
         This Test check the status of EEV and closes the EEV Valve
        """
        pass

    @staticmethod
    def EOLDrainHeaterTest():
        """
         This Test check the status of Drain Heater and turns it ON
        """
        pass
    
    @staticmethod
    def EOLDefrostHeaterTest():
        """
         This Test check the status of Defrost Heater and turns it ON and effect on Coil Outlet Temp Probe
        """
        pass

    @staticmethod
    def EOLLightTest():
        """
         This Test check the status of CDU and turns it off
        """
        global inputcoils
        global connection
        global gdeviceid

        print("Turning Light On.......")
        connection.write_coil(61, True, unit = gdeviceid)
        time.sleep(3)
        print("Turning Light Off.......")
        connection.write_coil(61, False, unit = gdeviceid)
        time.sleep(3)
        print("Turning Light On.......")
        connection.write_coil(61, True, unit = gdeviceid)
        time.sleep(3)
        print("Turning Light Off.......")
        connection.write_coil(61, False, unit = gdeviceid)


    @staticmethod
    def EOLDefrostTermination():
        """
        Check for Doors and then Check for Change in EOL Defrost Termination Temp Probe
        """
        pass

    @staticmethod
    def EOLFanTest():
        """
        Initialize the EOL Fan Test
        """
        global inputcoils
        global connection
        if(inputcoils[74] == True):
            print("Fan is On")
        else:
            print("Writing the Coil with Address 0x74.......")
            var=connection.write_coil(61, False, unit = gdeviceid)
            print(var)

    @staticmethod
    def EOLBuzzerTest():
        """
        Check for Doors and then Check for Change in EOL Defrost Termination Temp Probe
        """
        for i in range(5):
            var=connection.write_coil(40, True, unit = gdeviceid)
            time.sleep(2)
            var=connection.write_coil(40, False, unit = gdeviceid)



class mpxoneEOL():
    """

    Initial Release : Rohit Maity
    Date : 21/08/2022

    Lists the Type of Tests to be performed  by the Program

    """
    global V_CSM_EOL_Test_Result
    global V_CSM_EOL_Test_Status

    def __init__(self):
        pass
    def runTest(self):
        mpxoneTests.EOLLightTest() 
    def ioTest(self):
        pass
    def fullTest(self):
        pass





if __name__ == "__main__":

    client=mpxoneserialconnection(method = 'rtu', port='PORT10', stopbits = 2, bytesize = 8, parity = 'N' , baudrate= 19200,deviceid=199)

    obj=mpxoneEOL()

    obj.runTest()