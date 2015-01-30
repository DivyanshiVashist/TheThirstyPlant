import paho.mqtt.client as mqtt
import datetime
import json 
import time
from twython import Twython
import xlsxwriter


CONSUMER_KEY = "INSERT YOUR OWN CONSUMER KEY IN HERE!"
CONSUMER_SECRET = "INSERT YOUR OWN CONSUMER SECRET KEY IN HERE!"
ACCESS_KEY= "INSERT YOUR OWN ACCESS KEY IN HERE!"
ACCESS_SECRET = "INSERT YOUR OWN ACCESS SECRET KEY IN HERE!"

api = Twython(CONSUMER_KEY,CONSUMER_SECRET,ACCESS_KEY,ACCESS_SECRET)

start= time.time()
lastval = 0
thedate= 0 

# The callback for when the client receives a CONNACK response from the server.
def on_connect(client, userdata, flags, rc):

    # Subscribing in on_connect() means that if we lose the connection and
    # reconnect then subscriptions will be renewed.
    client.subscribe("/MarksPlant/data")

# The callback for when a PUBLISH message is received from the server.
def on_message(client, userdata, msg):
	
	global start 
	global lastval
	global thedate
	now = datetime.datetime.now()
	value = str(msg.payload)
	
		
	if int(value) >= 750 & int(value)<800 :
		stat = "I'm OK" 
	if 	int(value)>= 800:
		stat= "Help! Plant overboard !"	
	if int(value) < 750:
		stat = 'I need a drink !'

	
	
	
			
	#if time.time() - start >30000 or value!= lastval:
	if time.time() - start >3000 or  abs(int(value)-int(lastval)) >5 :
		api.update_status(status= "Status: "+ stat+ "\n"+"Timestamp: " + str(now.strftime("%Y-%m-%d %H:%M:%S")) + "\n"+"Reading: "+ value  )	
		start = time.time()
		lastval= value 
	
	#Every time post also save reading into txt file 
	if str(now.strftime("%Y-%m-%d")) == thedate:		
		with open(thedate +".txt", 'a') as file:
			file.write(str(now.strftime("%H:%M:%S")) +"|" + value+ "\n" )
	else:
		with open(str(now.strftime("%Y-%m-%d"))+'.txt', 'a') as file:
			file.write(str(now.strftime("%H:%M:%S")) +"|" + value +"\n")
		thedate = str(now.strftime("%Y-%m-%d")) 
			
			
	# At 4pm graph all the data of the day into a graph
	if str(now.strftime("%H:%M:%S")) >= "12:00:00":
		workbook = xlsxwriter.Workbook('Plant Data' + str(now.strftime("%Y-%m-%d"))+'.xlsx')
		worksheet = workbook.add_worksheet()
		file = open(thedate +".txt", 'r')
		xvals=[]
		yvals=[]
		for line in file: 
			splitted= line.split('|') 
			x=splitted[0]
			y=splitted[1]
			xvals.append(x)
			yvals.append(int(y)) 	
		worksheet.write_column('A1', xvals)
		worksheet.write_column('B1', yvals)
		length = str(len(yvals))
		chart = workbook.add_chart({'type': 'line'})
		#chart.add_series({'Time': '=Sheet1!$A1:$A$'+length})
		chart.add_series({'values': '=Sheet1!$B1:$B$'+length})
		worksheet.insert_chart('C1', chart)
		chart.set_title({'name': str(now.strftime("%Y-%m-%d")) + " Readings"})
		workbook.close()
			 
			
		
		
		
	
	
	
	client.publish("MarksPlant/JSON", json.dumps({'Reading':value, "Timestamp":str(now.strftime("%Y-%m-%d %H:%M:%S")) ,"Status":stat	}) )   
    
client = mqtt.Client()
client.on_connect = on_connect
client.on_message = on_message

client.connect("winter.ceit.uq.edu.au", 1883, 60)

# Blocking call that processes network traffic, dispatches callbacks and
# handles reconnecting.
# Other loop*() functions are available that give a threaded interface and a
# manual interface.






client.loop_forever()
