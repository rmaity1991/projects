from pymodbus.client.sync import ModbusSerialClient as ModbusClient 

class mpxoneserialconnection:
    """

    Initial Release : Rohit Maity
    Date : 21/08/2022

    Creates the serial connection to the mpxone basic controller for read and write registers
    
    """

    def __init__(self,method,port,stopbits,bytesize,parity,baudrate):
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
        :param timeout: The timeout between serial requests (default 3s)
        :param strict:  Use Inter char timeout for baudrates <= 19200 (adhere
         to modbus standards)
        :param handle_local_echo: Handle local echo of the USB-to-RS485 adaptor
    
        """
        self.method=method
        self.port=port
        self.stopbits=stopbits
        self.bytesize=bytesize
        self.parity=parity
        self.baudrate=baudrate
        self.connection=ModbusClient(method = self.method, port=self.port, stopbits = self.stopbits, bytesize = self.bytesize, parity = self.parity , baudrate= self.baudrate)
        print("Object Created")

        # client = ModbusClient(method = 'rtu', port='PORT10', stopbits = 2, bytesize = 8, parity = 'N' , baudrate= 38400)
        self.connection.connect()

        if(self.connection.connect()):
            print("Connection Succesful")
        else:
            print("Bad Connection")

    def read_holding_registers(self,address,slave,count=200):
        """

        Initial Release : Rohit Maity
        Date : 21/08/2022

        coil  = client.read_holding_registers(0x01,1,unit=1) address, count, slave address
    
        """
        try:
            # coil  = client.read_holding_registers(0x01,1,unit=1) address, count, slave address
            var=self.connection.read_holding_registers(address,count,unit=slave)
            print("Read Holding Registers Success")
            print(var.registers[0])
        except:
            print("Could not establish connection to the holding registers")
    
    def read_input_registers(self,address,slave,count=200):
        """

        Initial Release : Rohit Maity
        Date : 21/08/2022

        coil  = client.read_holding_registers(0x01,1,unit=1) address, count, slave address
    
        """
        try:
               # Starting add, num of reg to read, slave unit.
               # coil  = client.read_holding_registers(0x01,1,unit=1) address, count, slave address
               self.connection.connect()
               var = self.connection.read_input_registers(address,count,unit=slave)
               print("Read Input Register Success")
        except:
               print("Could not establish connection to the holding registers")



    def sendData(self):
        pass
