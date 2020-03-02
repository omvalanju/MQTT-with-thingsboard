import paho.mqtt.client as mqtt
import xlwt
import subprocess
from xlwt import Workbook
from datetime import datetime

count = 2
vibration = 'NA'

wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')

global vibration

def on_connect(client, userdata, flags, rc):
    print("Connected with result code {0}".format(str(rc)))
    client.subscribe("MakerIOTopic")


def on_message(client, userdata, msg):
    global vibration
    print("Message received-> " + msg.topic + " " + str(msg.payload)) 
    vibration = msg.payload.decode("utf-8")
    print(vibration)


client = mqtt.Client("Thinkpad")
client.on_connect = on_connect
client.on_message = on_message
print(client.on_message)
client.connect('127.0.0.1', 1885)


try:
    sheet1.write(1,0,"Time")
    sheet1.write(1,1,"Vibration")
    while True:
        now = datetime.now()
        current_time = now.strftime("%H:%M:%S")
        client.loop()
        print(vibration)
        sheet1.write(count,1,vibration)
        sheet1.write(count,0,current_time)
        count = count + 1
        bashCommand = 'mosquitto_pub -h 127.0.0.1 -t v1/devices/me/telemetry -u 6V5PDTFsqsHc4Wv2X6Qe -m {"Vibration":"' + vibration + '"}'
        output = subprocess.check_output(['bash','-c',bashCommand])

except KeyboardInterrupt:
    print("Exiting")
    wb.save('Vibration.xls')


