import paho.mqtt.client as mqtt #import the library for the MQTT
broker_address="broker.emqx.io" # broker server
client = mqtt.Client("rohitmaity") #create new instance of client 
client.connect(broker_address) #connect to broker address
client.subscribe("P1") # subscribe to the topic
client.publish("P1","100") # send data