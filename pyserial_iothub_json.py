from http import client
from azure.iot.device import IoTHubDeviceClient, Message
import random
import json
import time

conn_str ="HostName=rht-dfr-iothub.azure-devices.net;DeviceId=arduino;SharedAccessKey=07g73jtsrmfXq2xkIjwvtA8vbCCIDnaMcgEM9MIOmAs="
client = IoTHubDeviceClient.create_from_connection_string(conn_str)

message={}
while True:
    message={
        'currentTemp':random.randint(10,20),
        'currentSetPoint':random.randint(10,20),
        'currentPressure1':random.randint(2,5),
        'currentPressure2':random.randint(2,5)
    }

    message_json=json.dumps(message)
    print(message_json)
    iot_message=Message(message_json)
    print(iot_message)

    client.send_message(iot_message)
    time.sleep(2)

    
