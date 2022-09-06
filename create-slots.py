import smtplib, ssl, email
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import shutil
import importlib.util
import xlsxwriter
import pandas as pd
def generateTimeSlots(baseline, n, slotSize):
    numSlots = n
    while (not numSlots % slotSize == 0):
        numSlots += 1

    time = baseline
    slots = []
    for i in range(0, numSlots):
        slots.append(time)
        [hours, minutes] = time.replace("am", "").replace("pm", "").split(':')
        intMinutes = int(minutes) + slotSize
        intHours = int(hours)
        if (intMinutes + slotSize >= 60):
            intMinutes = intMinutes + slotSize - 60
            intHours = intHours + 1
            if (intHours == 13):
                intHours = 1

        newTime = str(intHours)+":"+str(intMinutes)
        if (len(str(intMinutes))== 1):
            newTime = str(intHours)+":0"+str(intMinutes)
        
        if (intHours < 11):
            newTime += "pm"
        else:
            newTime += "am"

        time = newTime
    return slots

# converts studentTimes (# they are in order) -- to actual times
def getRealTimes(studentTimes, slots):
    # print(studentTimes)
    realTimes = []
    for subArr in studentTimes:
        # print(subArr)
        arr = [subArr[0], slots[subArr[1]]]
        realTimes.append(arr)
    return realTimes

def getCol(i):
    cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    if (i >= len(cols)):
        return cols[i // len(cols) - 1] + cols[i % len(cols)]
    
    return cols[i]

def writeToTxt(totalTimes, emailTimes):
    # delete old directory if exists
    try:
        shutil.rmtree('auditions-times')
    except OSError as e:
        print('error: directory does not exist '+str(e))

    os.mkdir('auditions-times')

    # export to text files in auditions directory
    for key in totalTimes.keys():
        f = open("auditions-times/auditions-"+key+".txt", "x")
        f.write(key+",")
        for email in totalTimes[key]:
            if not email == 0:
                f.write(email+",")
        f.close()

    # print(emailTimes)
    f2 = open("auditions-times/student-times.txt", "x")
    for key in emailTimes.keys():
        string = key
        studTimes = emailTimes[key]
        for subArr in studTimes:
            string += "/"+str(subArr[0])+"-"+str(subArr[1])
        string += '\n'
        f2.write(string)
    f2.close()
    return True

def writeToExcel(totalTimes, slots):
    # write to excel
    try:
        os.remove('auditions-times-girls.xlsx')
    except OSError as e:
        # do nothing
        print('OSError:', e)

    workbook = xlsxwriter.Workbook('auditions-times-girls.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})
    rowNum = 2
    colWidthsSet = False
    for group in totalTimes.keys():

        if (not colWidthsSet):
            for i in range(0, len(totalTimes[group]) + 1):
                worksheet.set_column(i, i, 30)
            colWidthsSet = True
        worksheet.write(getCol(0) + str(rowNum), group, bold)

        v = totalTimes[group]
        for j in range(0, len(v)):
            if (not v[j] == 0):
                worksheet.write(getCol(j + 1) + str(rowNum), v[j])

        rowNum += 1

    for i in range(0, len(slots)):
        worksheet.write(getCol(i + 1) + '1', slots[i], bold)

    workbook.close()

def generateTestEmails(STUDENT_COUNT):
    input = []
    for i in range(0, STUDENT_COUNT):
        input.append('email'+str(i)+'@dartmouth.edu')
    return input

def initializeTotalTimes(groups):
    totalTimes = {}
    for group in groups:
        totalTimes[group] = []
    return totalTimes

def validateEmails(emails):
    errorFound = False
    for email in emails:
        # check for @dartmouth.edu
        ending = "dartmouth.edu"
        arr =  email.split("@")
        if (len(arr) < 2 or arr[1] != "dartmouth.edu"):
            print('FOUND INVALID EMAIL:', email)
            errorFound = True 
    return errorFound

# genIdentity: 1 = man, 2 = woman, 3 = other
def getEmails(genIdentity):
    emails = []
    df = pd.read_excel('auditions-emails.xlsx')
    list_of_columns = df.columns.values

    i = 1
    for index, row in df.iterrows():
        # print('row', row)
        name = row[1]
        identity = row[3]
        email = row[5]
        if ((identity == 'Man' and genIdentity == 1) or (identity == 'Woman' and genIdentity == 2) or (genIdentity == 3 and (identity == 'Woman' or identity == 'Man'))):
            emails.append(email)
        i += 1

    # print('res', emails)
    print('Read a total of', len(emails), 'emails')

    errorFound = validateEmails(emails)
    return [emails, errorFound]
   
def main():
    SLOT_SIZE = 5
    [input, errors] = getEmails(2)
    if errors:
        print('VALIDATION ERRORS FOUND: FIX IN EXCEL SHEET BEFORE PROCEEDING')
        # return
    else:
        print('All emails passed validation!')
    # STUDENT_COUNT = len(input)
    STUDENT_COUNT = 30
    input = generateTestEmails(STUDENT_COUNT)
    slots = generateTimeSlots('3:15pm', STUDENT_COUNT, SLOT_SIZE)
    groups = ['Aires', 'Dodecaphonics', 'Cords', 'Brovertones', 'Sings']
    # groups = ['Rockapellas', 'Sings', 'Decibelles', 'Dodecaphonics', 'Subtleties']
    totalTimes = initializeTotalTimes(groups)
    emailTimes = {} 

    y = 0
    lastY = 0
    while (y < len(input)):
        if ((y + 1) % 5 == 0):

            times = [[0] * 5, [0] * 5, [0] * 5, [0] * 5, [0] * 5]

            for i in range(0, 5):
                email = input[i + y - 4]
                studentTimes = []

                for j in range(0, 5):

                    index = i + j
                    if (index > 4):
                        index = index - 5
                    times[j][index] = email
                    studentTimes.append([groups[index], j + y - 4])
                
                # convert studentTimes to actual times
                realTimes = getRealTimes(studentTimes, slots)
                emailTimes[email] = realTimes

            for i in range(0, len(groups)):
                totalTimes[groups[i]].extend(times[i])
            lastY = y + 1
        y = y + 1

    # get stragglers (if not divisible by 5)
    if (not y % 5 == 0):
        times = [[0] * 5, [0] * 5, [0] * 5, [0] * 5, [0] * 5]
        for i in range(0, y - lastY):
            email = input[i + lastY]
            studentTimes = []

            for j in range(0, 5):
                index = i + j
                if (index > 4):
                    index = index - 5
                times[j][index] = email
                studentTimes.append([groups[index], j + lastY])

            # convert studentTimes to actual times
            realTimes = getRealTimes(studentTimes, slots)
            emailTimes[email] = realTimes

        for i in range(0, len(groups)):
            totalTimes[groups[i]].extend(times[i])
    print('studentCount', STUDENT_COUNT)
    print(totalTimes, slots)
    writeToExcel(totalTimes, slots)

main()