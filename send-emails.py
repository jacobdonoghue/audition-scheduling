import smtplib, ssl, email
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import shutil
import pandas as pd
from decouple import config

def readStudentTimes():
    print('readingStudentTimes2')
    studentTimes = {}
    df = pd.read_excel('auditions-times.xlsx')
    list_of_columns = df.columns.values
    groups = []
    for j in range(0, len(list_of_columns)):
        col = df[list_of_columns[j]]
        if (j == 0):
            for i in range(0, len(col)):
                groups.append(col[i])
        else:
            for i in range(0, len(col)):
                email = col[i]
                if (not pd.isna(email)):
                    times = []
                    # do nothing
                    if (email in studentTimes):
                        times = studentTimes[email]
                    times.append([groups[i], col.name])
                    studentTimes[email] = times

    # print(studentTimes)
    return studentTimes

def readLocations():
    f = open('locations.txt', 'r')
    lines = f.readlines()
    f.close()
    
    locations = {}
    for line in lines:
        arr = line.split(":")
        locations[arr[0]] = arr[1]
    
    return locations

def validate(studentTimes):
    prev = set() # "group-time" shouldn't occur more than once

    for email in studentTimes:
        times = set()
        for subArr in studentTimes[email]:

            s = str(subArr[0])+str(subArr[1])
            if s in prev:
                return False
            elif subArr[1] in times:
                print('time', subArr[1], 'already in times', times, studentTimes[email], email)
                return False
            prev.add(s)
            times.add(subArr[1])
    return True


def buildMessage(studentTimes, locations):
    message = "Hello!\n We're so excited to hear you sing! Here are your audition times:\n\n"
    for subArr in studentTimes:
        message += subArr[0]+": "+subArr[1]+" at "+locations[subArr[0]]+"\n"
    
    message += "\n Please warm up beforehand, and show up 5 minutes early to your first slot\n\n If, for any reason, you can't make these times, please respond to this email and we'll find a different time to fit you in. \n\n Much love, \n\n   the Dartmouth Acapella Groups :)"
    # print('message', message)
    return message

def main(send):
    emailTimes = readStudentTimes()
    valid = validate(emailTimes)
    if (valid):
        
        locations = readLocations()
        # log in
        smtp_server = "smtp.gmail.com"
        port = 587  # For starttls
        sender_email = config('EMAIL')
        password = config('PASSWORD')
        # Create a secure SSL context
        context = ssl.create_default_context()

        error = None
        # Try to log in to server and send email
        try:
            server = smtplib.SMTP(smtp_server,port)
            server.ehlo() # Can be omitted
            server.starttls(context=context) # Secure the connection
            server.ehlo() # Can be omitted
            server.login(sender_email, password)

            # send email to each student
            i = 0
            for studEmail in emailTimes.keys():
                message = MIMEMultipart()
                message["From"] = sender_email
                message["To"] = studEmail # replace with your email for testing purposes
                message["Subject"] = "Acapella Audition Times!!!"
                body = buildMessage(emailTimes[studEmail], locations)
                message.attach(MIMEText(body, "plain"))
                text = message.as_string()
                if (send):
                    server.sendmail(sender_email, studEmail, text)
                else:
                    print('sending email '+str(i)+": \n", body)
                    k = 0
                i += 1
        except Exception as e:
            # Print any error messages to stdout
            error = e
        finally:
            server.quit() 

        if (error == None):
            print('script exited with no errors')
        else:
            print('script exited with error', error)
    else:
        print('invalid data, cannot send out emails')

main(False)
