#import flask
#import django
import pandas as pd
#import xlwings as xw
from datetime import datetime
import datetime
import os
import win32com.client

#"CM T&T Reports"

import smtplib
from email.message import EmailMessage


def find_folder():

    today = datetime.date.today()

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = outlook.Folders("Change-Tickets")
    inbox = folder.Folders("Inbox")
    sub_folder = inbox.Folders("Information")
    target_folder = sub_folder.Folders("CM T&T Reports")
    #inbox = outlook.GetDefaultFolder(6)
    print(inbox)
    messages = target_folder.Items
    print(messages)
    return messages
    #messages = messages.Sort( "[ReceivedTime]" , True )


def save_attachments(SavePath,folder,subject):
    result = ""
    for message in folder:
        print(message.Senton.date())
        if message.Subject == subject and message.Unread: # or message.Senton.date() == today:
            attachments = message.Attachments
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(SavePath, str(attachment)))
                result = str(SavePath +"/"+ str(attachment))
                #if message.Subject == subject and message.Unread:
                #    message.Unread = False
                break

            break
    return result


def prepare_file(path):
    counter = 0
    df = pd.read_csv(path,header=0)
    df = df.astype(str)
    df['Completion Date'] = pd.to_datetime(df['Completion Date'], format='%Y%m%d', errors='ignore')
    completed = []
    for i in df['Completion Date']:
        i = i[:-4]
        i = datetime.datetime.strptime(i, '%d/%m/%Y %H:%M')
        completed.append(i)
    df['Completion Date'] = completed
    df = df.sort_values(by=['Nordea ID', 'Title', 'Completion Date'], ascending=[True, True, False])
    print(df.head(n=100).to_string())
    return df


def send_mails():
    for i in range(1): #len(mail_list)
        body = "Hell Yea!"

        email = EmailMessage()
        email.set_content(body, subtype='html')

        to =  'M014207'#'krzysztof.sztuk@nordea.com' #MailList[i]
        email['From'] = "marcin.grabowski@nordea.com"
        email['To'] = to
        email['Subject'] = "Siemandero"
        email['bcc'] = "marcin.grabowski@nordea.com"

        smtp_connection = smtplib.SMTP('email.oneadr.net', 25)
        status = smtp_connection.send_message(email)
        print(str(status))
        print(to)
    pass




if __name__ == '__main__':
    # MailList = ["agnieszka.ucinska@nordea.com", 'gabriela.cholewicka@nordea.com', 'krzysztof.sztuk@nordea.com',
    #              'marcin.grabowski@nordea.com']
    # Path = os.path.expanduser("~/Desktop/New")
    # DestinationFolder = find_folder()#func()
    # File_Path = save_attachments(Path,DestinationFolder,"Background Report Job Email Notification")
    # prepare_file(File_Path)
    # prepare_file("C:/Users/M014205/Desktop/New/Helix report.csv")
    # print(File_Path)
    send_mails()
