
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

def merge_emails():
    file1 = pd.read_csv("C:/Users/M014207/Desktop/New/Helix report.csv")
    file2 = pd.read_excel("C:/Users/M014207/Desktop/New/MailList.xlsx")

    file3 = file1[["Title", "Nordea ID", "Last Name", "First Name", "Completion Date"]].merge(file2[["Nordea ID", "Internet_E_mail"]], on= "Nordea ID", how = 'left')
    file3.to_csv("C:/Users/M014207/Desktop/New/Results.csv", index = False)

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
    df = df.drop_duplicates(subset= ["Nordea ID", "Title"])
    df = df.sort_values(by=['Completion Date'], ascending=[True])
    #print(df.head(n=100).to_string())
    return df



def send_mails(data):
    today = datetime.datetime.now()
    months_4 = today - datetime.timedelta(days=120)
    months_5 = today - datetime.timedelta(days=150)
    months_6 = today - datetime.timedelta(days=180)
    c4 = 0
    c5 = 0
    c6 = 0

    for i in data['Completion Date']:
        # print(type(i))
        # l = i.timestamp()
        # print(type(today))
        # print(type(l))
        # print(l)
        # break


        if i < months_6:
            c6 += 1
        elif i < months_5:
            c5 += 1
        elif i < months_4:
            c4 += 1
    print(c6)
    print(c5)
    print(c4)
    print(c4+c5+c6)

    #     if i.timestamp() +  < today:
    #         c6 += 1
    #         print('6 months')
    #     elif i.timestamp() + datetime.timedelta(150) < today:
    #         c5 += 1
    #         print('5 months')
    #     elif i.timestamp() + datetime.timedelta(120) < today:
    #         c4 += 1
    #         print('4 months')
    #
    # print(c4)
    # print(c5)
    # print(c6)



    #     body = "Hell Yea!"
    #
    #     email = EmailMessage()
    #     email.set_content(body, subtype='html')
    #
    #     to =  'M014207'#'krzysztof.sztuk@nordea.com' #MailList[i]
    #     email['From'] = "marcin.grabowski@nordea.com"
    #     email['To'] = to
    #     email['Subject'] = "Siemandero"
    #     email['bcc'] = "marcin.grabowski@nordea.com"
    #
    #     smtp_connection = smtplib.SMTP('email.oneadr.net', 25)
    #     #status = smtp_connection.send_message(email)
    #     print(str(status))
    #     print(to)
    # pass



if __name__ == '__main__':
    # MailList = ["agnieszka.ucinska@nordea.com", 'gabriela.cholewicka@nordea.com', 'krzysztof.sztuk@nordea.com',
    #              'marcin.grabowski@nordea.com']
    # Path = os.path.expanduser("~/Desktop/New")
    # DestinationFolder = find_folder()#func()
    # File_Path = save_attachments(Path,DestinationFolder,"Background Report Job Email Notification")
    #merge_emails()
    # prepare_file(File_Path)
    data = prepare_file("C:/Users/M014207/Desktop/New/Results.csv")
    # print(File_Path)
    send_mails(data)

