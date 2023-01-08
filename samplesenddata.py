from azure.iot.device import IoTHubDeviceClient, Message
import time
import random

sensor_data='{{"Current Temperature":{temperature},"Current Pressure":{pressure}}}'

conn_str="HostName=rht-dfr-iothub.azure-devices.net;DeviceId=sensor1;SharedAccessKey=8zfoS7qmbFmOHjlZzSq4PIwUCFbbsnjkSUPrXyoyTfw="

client=IoTHubDeviceClient.create_from_connection_string(conn_str)


count=50

try:
    while(count>0):
        temperature=20+random.randint(1,8)
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
        count=count-1;

except:
    print("Error has Occured")
    

client.shutdown()

        

        