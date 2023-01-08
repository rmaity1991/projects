from azure.iot.device import Message
from azure.iot.device.aio import IoTHubDeviceClient
import numpy as np

connection_string='HostName=rht-dfr-iothub.azure-devices.net;DeviceId=rht-dfr-device1;SharedAccessKey=3J1wrHffOoB60MxXP9EpqQ+37MeaRL5pCH2y1rOqfz4='

PAYLOAD = '{{"temperature": {temperature}}}'

client=IoTHubDeviceClient.create_from_connection_string(connection_string)


simulation_temperature_list=np.random.random(10)


for var in simulation_temperature_list:
    data = PAYLOAD.format(temperature=var)

    message = Message(data)
    print(message)
    print(f"Sending message: {message}")
    client.send_message(message)
    print("Message successfully sent")



client.shutdown()



