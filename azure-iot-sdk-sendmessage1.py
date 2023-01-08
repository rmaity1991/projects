import os
import asyncio
import uuid
from azure.iot.device.aio import IoTHubDeviceClient
from azure.iot.device import Message

messages_to_send = 10


def main():
    # Connection string should never be stored in code, but for simplicity we are using it.
    conn_str = "Insert your connection String"

    # The client object is used to interact with your Azure IoT hub.
    device_client = IoTHubDeviceClient.create_from_connection_string(conn_str)

    # Connect the client.
    device_client.connect()

    def send_test_message(i):
        print("sending message #" + str(i))
        msg = Message("test wind speed " + str(i))
        msg.message_id = uuid.uuid4()
        msg.correlation_id = "correlation-1234"
        msg.custom_properties["tornado-warning"] = "yes"
        msg.content_encoding = "utf-8"
        msg.content_type = "application/json"
        device_client.send_message(msg)
        print("done sending message #" + str(i))

    # send `messages_to_send` messages in parallel
    x=[send_test_message(i) for i in range(1, messages_to_send + 1)]

    # Finally, shut down the client
    device_client.shutdown()


if __name__ == "__main__":
    main()