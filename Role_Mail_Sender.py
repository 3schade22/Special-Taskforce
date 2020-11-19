import pandas as pd
# import xlwings as xw
from datetime import datetime
import datetime
import os
import win32com.client

# "CM T&T Reports"

import smtplib
from email.message import EmailMessage



def find_folder():
    today = datetime.date.today()

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = outlook.Folders("Change-Tickets")
    inbox = folder.Folders("Inbox")
    sub_folder = inbox.Folders("Information")
    target_folder = sub_folder.Folders("CM T&T Reports")
    # inbox = outlook.GetDefaultFolder(6)
    print(inbox)
    messages = target_folder.Items
    print(messages)
    return messages
    # messages = messages.Sort( "[ReceivedTime]" , True )


def save_attachments(SavePath, folder, subject):
    yesterday = datetime.date.today() - datetime.timedelta(days=1)
    result = ""
    for message in folder:
        print(message.Senton.date())
        if message.Subject == subject and message.Unread and message.Senton.date() == yesterday:
            attachments = message.Attachments
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(SavePath, str(attachment)))
                result = str(SavePath + "/" + str(attachment))
                # if message.Subject == subject and message.Unread:
                #    message.Unread = False
                break

            break
    return result


def merge_emails():
    file1 = pd.read_csv("C:/Users/M014207/Desktop/New/Helix report.csv")
    file2 = pd.read_excel("C:/Users/M014207/Desktop/New/MailList.xlsx")

    file3 = file1[["Title", "Nordea ID", "Last Name", "First Name", "Completion Date"]].merge(
        file2[["Nordea ID", "Internet_E_mail"]], on="Nordea ID", how='left')
    file3.to_csv("C:/Users/M014207/Desktop/New/Results.csv", index=False)


def prepare_file(path):
    counter = 0
    df = pd.read_csv(path, header=0)
    df = df.astype(str)
    df['Completion Date'] = pd.to_datetime(df['Completion Date'], format='%Y%m%d', errors='ignore')
    completed = []
    for i in df['Completion Date']:
        i = i[:-4]
        i = datetime.datetime.strptime(i, '%d/%m/%Y %H:%M')
        completed.append(i)
    df['Completion Date'] = completed
    df = df.sort_values(by=['Nordea ID', 'Title', 'Completion Date'], ascending=[True, True, False])
    df = df.drop_duplicates(subset=["Nordea ID", "Title"])
    df = df.sort_values(by=['Completion Date'], ascending=[False])
    # print(df.head(n=100).to_string())
    df.to_csv("C:/Users/M014207/Desktop/New/Ready.csv", index=False)
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
        # body = f"This is a message for period {period} <br> Role: {data['Title'][i]} <br> Mail: {data['Internet_E_mail'][i]} " \
        # f"<br> First Name: {data['First Name'][i]} <br> Last Name: {data['Last Name'][i]} <br> Date: {data['Completion Date'][i]}<br>"
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

    # rola, period
    email.set_content(body2, subtype='html')


    if len(data['Internet_E_mail'][i] ) < 4:
        missing_mails.append((data["Nordea ID"][i], data['Title'][i]))
        pass



    # to = 'marcin.grabowski@nordea.com' #data['Internet_E_mail'][i]
    #
    # email['From'] = "krzysztof.sztuk@nordea.com"
    # email['To'] = to
    # email['bcc'] = "marcin.grabowski@nordea.com"
    #
    # smtp_connection = smtplib.SMTP('email.oneadr.net', 25)
    # status = smtp_connection.send_message(email)
    # print(str(status))

def send_missing_emails():

    email = EmailMessage()
    if missing_mails != []:
        body = str(missing_mails)

    else:
        body = "No missing emails found"

    email.set_content(body, subtype='html')
    #email.add_attachment(attachment)
    to = "krzysztof.sztuk@nordea.com" #"change.tickets@nordea.com"

    email['From'] = "krzysztof.sztuk@nordea.com"
    email['To'] = to
    email['bcc'] = "marcin.grabowski@nordea.com"

    smtp_connection = smtplib.SMTP('email.oneadr.net', 25)
    status = smtp_connection.send_message(email)
    print(str(status))

def testing_dataFrame():
    people = {
        "Title": [],
        "Mail": [],
        "Date": [],
        "Months": []
    }

    people["Title"] = testing_list1
    people["Mail"] = testing_list2
    people["Date"] = testing_list3
    people["Months"] = testing_list4

    test_frame = pd.DataFrame(data=people)
    #print(test_frame.head(n=100).to_string())
    test_frame.to_csv("C:/Users/M014207/Desktop/New/Sendout_Info.csv", index=False)
    # return test_attachment


def send_mails(data):
    today = datetime.datetime.now()
    months_4 = today - datetime.timedelta(days=120)
    months_5 = today - datetime.timedelta(days=150)
    months_6 = today - datetime.timedelta(days=180)

    c = 0
    for i in data.index:
        c += 1

        if data['Title'][i] == "Helix Release Manager/Coordinator":
            continue

        if data['Completion Date'][i] < months_6:
            prepare_mail_body(i, "6 months")
            testing_list1.append(data['Title'][i])
            testing_list2.append(data['Internet_E_mail'][i])
            testing_list3.append(data['Completion Date'][i])
            testing_list4.append(6)
            #break

        elif data['Completion Date'][i] < months_5:
            prepare_mail_body(i, "5 months")
            testing_list1.append(data['Title'][i])
            testing_list2.append(data['Internet_E_mail'][i])
            testing_list3.append(data['Completion Date'][i])
            testing_list4.append(5)
            #break

        elif data['Completion Date'][i] < months_4:
            prepare_mail_body(i, "4 months")
            testing_list1.append(data['Title'][i])
            testing_list2.append(data['Internet_E_mail'][i])
            testing_list3.append(data['Completion Date'][i])
            testing_list4.append(4)
           # break

    print(c)
    # return testing_dataFrame()






if __name__ == '__main__':
    missing_mails = []
    testing_list1 = []
    testing_list2 = []
    testing_list3 = []
    testing_list4 = []

    # Path = os.path.expanduser("~/Desktop/New")
    # DestinationFolder = find_folder()#func()
    # File_Path = save_attachments(Path,DestinationFolder,"Background Report Job Email Notification")
    # merge_emails()
    # prepare_file(File_Path)
    data = prepare_file("C:/Users/M014207/Desktop/New/Results.csv")
    # print(File_Path)
    send_mails(data)
    #attachment = "C:/Users/M014207/Desktop/New/Sendout_Info.csv"
    send_missing_emails()
