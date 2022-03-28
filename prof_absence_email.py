import win32com.client
from twilio.rest import Client

outlook = win32com.client.Dispatch("Outlook.Application")
outlook_ns = outlook.GetNamespace("MAPI")

MyFolder = outlook_ns.Folders["shreya_sharma1@baylor.edu"].Folders["Inbox"]

Messages = MyFolder.Items
MessageCount = 0

for Message in Messages:
    if Message.UnRead:
        print(Message.sender)
        print(Message.subject)

        if "absence" in Message.subject:
            print("Found message with absence")

            Msg = outlook.CreateItem(0)
            Msg.Importance = 1
            Msg.Subject = "Got your " + Message.sunbect + " email"
            Msg.HTMLBody = "Hi" + str(Message.sender) + "\n" + " sorry you are not well"

            Msg.To = Message.sender.GetExchangeUser().PrimarySmtpAddress 
            Msg.ReadReceiptRequested = True 

            Msg.Send()