import smtplib, ssl, email
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import shutil
import importlib.util
import xlsxwriter

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
        os.remove('auditions-times.xlsx')
    except OSError as e:
        # do nothing
        print('')

    workbook = xlsxwriter.Workbook('auditions-times.xlsx')
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

def main():
    STUDENT_COUNT = 102
    SLOT_SIZE = 5

    input = generateTestEmails(STUDENT_COUNT)
    slots = generateTimeSlots('11:00am', STUDENT_COUNT, SLOT_SIZE)
    groups = ['Aires', 'Brovertones', 'Cords', 'Dodecaphonics', 'Sings']
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

    writeToExcel(totalTimes, slots)

main()