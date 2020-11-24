import pandas as pd
# import xlwings as xw
from datetime import datetime
import datetime
import os
import win32com.client

from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

# "CM T&T Reports"

import smtplib
from email.message import EmailMessage


# Find Outlook folder, where the report is
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


# Save the most current report
def save_attachments(SavePath, folder, subject):
    yesterday = datetime.date.today() - datetime.timedelta(days=1)
    file_path = ""
    for message in folder:
        print(message.Senton.date())
        if message.Subject == subject and message.Unread and message.Senton.date() == yesterday:
            attachments = message.Attachments
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(SavePath, str(attachment)))
                file_path = str(SavePath + "/" + str(attachment))
                # if message.Subject == subject and message.Unread:
                #    message.Unread = False
                break

            break
    return file_path


# add mail addresses to the report
def merge_emails(file1):
    file1 = pd.read_csv("C:/Users/Public/Documents/HelixRoles/Helix report.csv")
    file2 = pd.read_excel("C:/Users/Public/Documents/HelixRoles/MailList.xlsx")

    file3 = file1[["Title", "Nordea ID", "Last Name", "First Name", "Completion Date"]].merge(
        file2[["Nordea ID", "Internet_E_mail"]], on="Nordea ID", how='left')
    file3.to_csv("C:/Users/Public/Documents/HelixRoles/Results.csv", index=False)
    file_path = "C:/Users/Public/Documents/HelixRoles/Results.csv"

    return file_path


# prepare data for the send out
def prepare_file(file_path):
    counter = 0
    df = pd.read_csv(file_path, header=0)
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
    df.to_csv("C:/Users/Public/Documents/HelixRoles/Ready.csv", index=False)
    return df


# choose the correct mail template and trigger email send out
def choose_mail_type(data):
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
    testing_dataFrame()


def prepare_mail_body(i, period):
    email = EmailMessage()
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

    email = MIMEMultipart()
    if missing_mails != []:
        body = str(missing_mails)

    else:
        body = "No missing emails found"

    body = MIMEText(body, 'plain')
    email.attach(body)     #set_content(body, subtype='html')
    with open('C:/Users/Public/Documents/HelixRoles/Sendout_Info.csv','rb') as file:
    # Attach the file with filename to the email
        email.attach(MIMEApplication(file.read(), Name="Sendout Info.csv"))
    # attach_file_name = 'C:/Users/Public/Documents/HelixRoles/Sendout_Info.csv'
    # attach_file = open(attach_file_name, 'rb')  # Open the file as binary mode
    # payload = MIMEBase('application', 'octate-stream')
    # payload.set_payload((attach_file).read())
    # encoders.encode_base64(payload)  # encode the attachment
    # add payload header with filename
    # payload.add_header('Content-Decomposition', 'attachment', filename=attach_file_name)
    # email.attach(payload)
    #email.add_attachment(attachment)
    to = "marcin.grabowski@nordea.com" #"change.tickets@nordea.com"

    email['From'] = "marcin.grabowski@nordea.com"
    email['To'] = to
    email['bcc'] = "krzysztof.sztuk@nordea.com"

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
    test_frame.to_csv("C:/Users/Public/Documents/HelixRoles/Sendout_Info.csv", index=False)
    # return test_attachment


if __name__ == '__main__':
    missing_mails = []
    testing_list1 = []
    testing_list2 = []
    testing_list3 = []
    testing_list4 = []

    # Path = os.path.expanduser("C:/Users/Public/Documents/HelixRoles")
    # DestinationFolder = find_folder()#func()
    # File_Path = save_attachments(Path, DestinationFolder, "Background Report Job Email Notification")
    # File_Path = merge_emails(File_Path)
    # data = prepare_file(File_Path)
    data = prepare_file("C:/Users/Public/Documents/HelixRoles/Results.csv")
    # print(File_Path)
    choose_mail_type(data)
    # attachment = "C:/Users/M014207/Desktop/New/Sendout_Info.csv"
    send_missing_emails()
