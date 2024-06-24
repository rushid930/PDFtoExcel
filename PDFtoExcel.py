# importing required modules 
from pypdf import PdfReader
import tkinter as tk
from tkinter import filedialog
import xlsxwriter
from datetime import datetime

def import_file():
    file_path = filedialog.askopenfilename(title="Select a file", filetypes=[("Python files", "*.pdf"), ("All files", "*.*")])

    if file_path:
        reader = PdfReader(file_path)

        newFilePath = file_path.replace('.pdf', '.xlsx')
        workbook = xlsxwriter.Workbook(newFilePath)
        worksheet = workbook.add_worksheet()

        fullPDF = [["Count", "Sent", "From", "To", "First Viewed", "Subject", "Message", "Attachment"]]

        text = ''
        count = 0
        for x in reader.pages:
            page = reader.pages[count]
            count += 1
        
            # extracting text from page 
            text += page.extract_text()
            text += '\n'

        myArray = text.split("\n")

        wordsMessage = []
        messageMatches = ["Message", "of"]
        sentMatches = ["Sent: 0", "Sent: 1"]
        sentText = ''
        messageText = ''
        fromMatches = ["From: "]
        toMatches = ["To: "]
        subjectMatches = ["Subject: "]
        attachMatches = ["Attachments: "]
        pageMatches = ["Page", "of"]

        messageLine = ''
        attachLine = ''
        count2 = 0

        messageCount = 0
        sentCount = 0
        textCount = 0
        fromCount = 0
        toCount = 0
        subjectCount = 0
        attachCount = 0
        pageCount = 0

        for line in myArray:
            #Start of Message
            if all(x in line for x in messageMatches):
                if attachCount == 1:
                    # Append to Message Array
                    wordsMessage.append(messageLine)
                    wordsMessage.append(sentText)
                    wordsMessage.append(fromLine)
                    wordsMessage.append(toLine)
                    wordsMessage.append(firstViewLine)
                    wordsMessage.append(subjectLine)
                    messageText = messageText[:100]
                    wordsMessage.append(messageText)
                    wordsMessage.append(attachLine)
                    fullPDF.append(wordsMessage.copy())

                    wordsMessage.clear()
                    attachLine = ''
                    messageText = ''
                    messageCount = 0
                    sentCount = 0
                    textCount = 0
                    fromCount = 0
                    toCount = 0
                    subjectCount = 0
                    attachCount = 0
                    pageCount = 0

                messageArray = line.split('of ')
                messageLine = int(messageArray[0].replace('Message ', ''))
                messageCount = 1
            if messageCount == 1:
                if textCount == 1 or (attachCount == 1 and pageCount == 1 and messageText == ''):
                    messageText += line
                if textCount != 2:
                    textCount = 1
            if any(x in line for x in sentMatches) & textCount == 1:
                sentArray = line.split('Sent: ')
                myArray2 = messageText.split('Sent: ')
                sentText = sentArray[1].replace('at ', '')
                messageText = myArray2[0]
                
                sentCount = 1
                textCount = 2                    
            if all(x in line for x in fromMatches) & sentCount == 1:            
                fromLine = line.replace('From: ', '')
                fromCount = 1
            if all(x in line for x in toMatches) & fromCount == 1:
                myArray3 = line.split('(First Viewed: ')
                toLine = myArray3[0].replace('To: ', '')
                firstViewLine = myArray3[1].replace('at ', '').replace(')', '')
                toCount = 1
            if all(x in line for x in subjectMatches) & toCount == 1:
                subjectLine = line.replace('Subject: ', '').replace('RE: ', '').replace('Re: ', '')
                subjectCount = 1            
            if subjectCount == 1:                 
                if all (x in line for x in attachMatches):
                    attachLine = line.replace('Attachments: ', '')
                attachCount = 1
                pageCount = 0
            if all(x in line for x in pageMatches):
                pageCount = 1

            count2 += 1
        wordsMessage.append(messageLine)
        wordsMessage.append(sentText)
        wordsMessage.append(fromLine)
        wordsMessage.append(toLine)
        wordsMessage.append(firstViewLine)
        wordsMessage.append(subjectLine)
        messageText = messageText[:100]
        wordsMessage.append(messageText)
        wordsMessage.append(attachLine)
        fullPDF.append(wordsMessage.copy())
        print("done")

        col = 0

        for row, data in enumerate(fullPDF):
            worksheet.write_row(row, col, data)

        workbook.close()

# Create the main Tkinter window
root = tk.Tk()
root.title("Import File Example")
 
# Create an "Import File" button
import_button = tk.Button(root, text="Import File", command=import_file)
import_button.pack(pady=100)

root.mainloop()