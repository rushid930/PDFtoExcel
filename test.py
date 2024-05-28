# importing required modules 
from pypdf import PdfReader 
  
# creating a pdf reader object 
reader = PdfReader('OFW_Messages_Report_2024-05-20_18-53-36.pdf') 

import xlsxwriter

workbook = xlsxwriter.Workbook('arrays.xlsx')
worksheet = workbook.add_worksheet()

fullPDF = []

count = 0
for x in reader.pages:
    page = reader.pages[count]
    count += 1
  
    # extracting text from page 
    text = page.extract_text()

    myArray = text.split("\n")

    wordsMessage = []
    messageMatches = ["Message", "of 126"]
    sentMatches = ["Sent: "]
    sentText = ""
    fromMatches = ["From: "]
    toMatches = ["To: "]
    subjectMatches = ["Subject: "]

    fullPage = []

    for line in myArray:
        #Start of Message
        if all(x in line for x in messageMatches):
            wordsMessage.append(line)
        if all(x in line for x in sentMatches):
            myArray2 = line.split('Sent: ')
            sentText = ("Sent: " + myArray2[1])  
            messageText = (myArray2[0]) 
            wordsMessage.append(sentText)
        if all(x in line for x in fromMatches):
            wordsMessage.append(line)
        if all(x in line for x in toMatches):
            wordsMessage.append(line)
        if all(x in line for x in subjectMatches):
            wordsMessage.append(line)
            wordsMessage.append(messageText)
            fullPage.append(wordsMessage.copy())
            wordsMessage.clear()

    fullPDF.extend(fullPage.copy())
    fullPage.clear()

col = 0

for row, data in enumerate(fullPDF):
    worksheet.write_row(row, col, data)

workbook.close()
