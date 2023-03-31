import slack 
import os 
import win32com.client
import datetime 
from flask import Flask, request, Response 
from pathlib import Path 
from dotenv import load_dotenv
from slackeventsapi import SlackEventAdapter 

flag = False

def createorder (account,Dest,Type,Exch,Life,SecCode,Side,vol,price): 

    rqm = win32com.client.Dispatch("IressServerApi.RequestManager")

    request = rqm.CreateMethod("IOSPLUS", "CLIENTTEST", "OrderCreate3", 0)

    available_fields = request.Output.DataRows.GetAvailableFields()

    request.Input.Parameters.Set('AccountCode', account, 0)
    request.Input.Parameters.Set('Destination', Dest, 0)
    request.Input.Parameters.Set('PricingInstructions', Type, 0)
    request.Input.Parameters.Set('Exchange', Exch, 0)
    request.Input.Parameters.Set('Lifetime', Life, 0)
    request.Input.Parameters.Set('SecurityCode', SecCode, 0)
    request.Input.Parameters.Set('SideCode', Side, 0)
    request.Input.Parameters.Set('OrderVolume', vol, 0)
    request.Input.Parameters.Set('OrderPrice', price, 0)
   

    request.Execute()

    available_fields = request.Output.DataRows.GetAvailableFields()

    row_count = request.Output.DataRows.GetCount()

    column_count = len(available_fields)

    if row_count > 0:

        data = request.Output.DataRows.GetRows(available_fields, 0, -1)


    for row in range(row_count):

        for column in range(column_count):

            if available_fields[column] == "OrderNumber":

                print (data[row][column])
            
            if available_fields[column] == "ErrorMessage":

                print (data[row][column])


env_path = Path ('.')/ '.env'
load_dotenv(dotenv_path=env_path)

app = Flask (__name__)
slack_event_adapter = SlackEventAdapter(os.environ ['SIGNING_SECRET'], '/slack/events', app)

client = slack.WebClient (token=os.environ['SLACK_TOKEN'])
BOT_ID = client.api_call ("auth.test")['user_id']

@slack_event_adapter.on('message')
def message (payload):
    event = payload.get('event', {} )
    channel_id = event.get('channel')
    user_id = event.get ('user')
    text = event.get ('text')
    sub = '=injectorders('
    list = []

    if sub in text:

        list = text.split(',')
        listt = list[0].split ('(')
        listtt = list[4].split (')')
        print (list)
        print (listt)

        symbol = listt[1]
        Exchange  = list[1]
        step = list[2]
        ordercount = list[3]
        spread = listtt[0]
        Lastprice = 22.41
        flag = True
    
    return symbol,Exchange,step,ordercount,spread, Lastprice, flag


        
#@app.route('/injectorders', methods = ['POST'])
#def message_count():
     
#    client.chat_postMessage(channel='test4', text = "Please Enter :  =injectorders(symbol,exchange, step, ordercount, spread)")



while True:

    if flag == True:

        for i in range (int(ordercount)):

            createorder("CLIENTTEST","BESTMKT","LIMIT","TSX" ,"DAY",symbol,1,20,(Lastprice+i*float(step)-float(spread)))

        for i in range (int(ordercount)):

            createorder("CLIENTTEST","BESTMKT","LIMIT","TSX" ,"DAY",symbol,2,20,(Lastprice-i*(step)))

if __name__ == "__main__":
    app.run(debug=True)

