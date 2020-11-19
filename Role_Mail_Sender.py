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
    yesterday = datetime.date.today() - datetime.timedelta(days=1)
    result = ""
    for message in folder:
        print(message.Senton.date())
        if message.Subject == subject and message.Unread and message.Senton.date() == yesterday:
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
    file1 = pd.read_csv("C:/Users/M014205/Desktop/New/Helix report.csv")
    file2 = pd.read_excel("C:/Users/M014205/Desktop/New/MailList.xlsx")

    file3 = file1[["Title", "Nordea ID", "Last Name", "First Name", "Completion Date"]].merge(file2[["Nordea ID", "Internet_E_mail"]], on= "Nordea ID", how = 'left')
    file3.to_csv("C:/Users/M014205/Desktop/New/Results.csv", index = False)

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
    df = df.sort_values(by=['Completion Date'], ascending=[False])
    #print(df.head(n=100).to_string())
    df.to_csv("C:/Users/M014205/Desktop/New/Ready.csv", index = False)
    return df


def prepare_mail_body(i, period):
    email = EmailMessage()
    link = ""
    print(data['Title'][i])
    if data['Title'][i] == "Helix Change Manager":
        link = "https://wiki.itgit.oneadr.net/display/ServMan/How+to+become+a+Change+Manager"
    elif data['Title'][i] == "Helix Change Coordinator":
        link = 'https://wiki.itgit.oneadr.net/display/ServMan/How+to+become+a+Change+Coordinator'
    else:
        link = 'https://wiki.itgit.oneadr.net/display/ServMan/How+to+become+a+Change+Implementer'

    if period == '4 months' or period == '5 months':
        #body = f"This is a message for period {period} <br> Role: {data['Title'][i]} <br> Mail: {data['Internet_E_mail'][i]} " \
           #f"<br> First Name: {data['First Name'][i]} <br> Last Name: {data['Last Name'][i]} <br> Date: {data['Completion Date'][i]}<br>"
        body2 = f"Dear {data['First Name'][i]} {data['Last Name'][i]},<br> <br> Please be informed that you have accomplished training " \
                f"for the role of {data['Title'][i]} at least {period} ago. Recertification is required to keep the role" \
                f" no later than 6 months from the last certification. " \
                f"<br> <br>Training can be found on People Portal. All steps are described under below link: " \
                f"<br> <br>{link}<br> <br>" \
                f"If the training is not repeated, the role will be revoked.<br><br><br> Kind Regards,<br>" \
                f"IT Change Management Team"
        email['Subject'] = f"{data['Title'][i]} role recertification required."
    else:
        body2 = f"Dear {data['First Name'][i]} {data['Last Name'][i]},<br> <br>Please be informed that your {data['Title'][i]} role in Helix" \
                f" has been revoked As you did not accomplish the training for recertification. <br> <br>" \
                f"Training can be found on People Portal. When you will complete the training you can access the" \
                f" role via ITSSP after 24h. All steps are described under below link. <br> <br> " \
                f"https://wiki.itgit.oneadr.net/display/ServMan/How+to+become+a+Change+Manager <br> <br>" \
                f"If the training is not repeated, the role will be revoked.<br><br><br> Kind Regards,<br>" \
                f"IT Change Management Team"
        email['Subject'] = f"{data['Title'][i]} role revoked. Recertification required."

    #rola, period
    email.set_content(body2, subtype='html')

    to = 'krzysztof.sztuk@nordea.com'
    email['From'] = "marcin.grabowski@nordea.com"
    email['To'] = to
    email['bcc'] = "marcin.grabowski@nordea.com"

    smtp_connection = smtplib.SMTP('email.oneadr.net', 25)
    status = smtp_connection.send_message(email)
    print(str(status))


def send_mails(data):
    today = datetime.datetime.now()
    months_4 = today - datetime.timedelta(days=120)
    months_5 = today - datetime.timedelta(days=150)
    months_6 = today - datetime.timedelta(days=180)
    c4 = 0
    c5 = 0
    c6 = 0
    c = 0
    for i in data.index:
        c +=1
        # if data['Title'][i] == "Helix Change Coordinator":
        #    data['Title'][i] = "Helix Release Manager/Coordinator"

        if data['Title'][i] == "Helix Release Manager/Coordinator":
            continue

        if data['Completion Date'][i] < months_6:
            prepare_mail_body(i, "6 months")
            break

        elif data['Completion Date'][i] < months_5:
            prepare_mail_body(i, "5 months")
            break
        elif data['Completion Date'][i] < months_4:
            prepare_mail_body(i, "4 months")
            break

    print(c)
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


if __name__ == '__main__':
    # MailList = ["agnieszka.ucinska@nordea.com", 'gabriela.cholewicka@nordea.com', 'krzysztof.sztuk@nordea.com',
    #              'marcin.grabowski@nordea.com']
    #Path = os.path.expanduser("~/Desktop/New")
    #DestinationFolder = find_folder()#func()
    #File_Path = save_attachments(Path,DestinationFolder,"Background Report Job Email Notification")
    #merge_emails()
    #prepare_file(File_Path)
    data = prepare_file("C:/Users/M014205/Desktop/New/Results.csv")
    #print(File_Path)
    send_mails(data)
