# importing required modules 
from pypdf import PdfReader
import tkinter as tk
from tkinter import filedialog
import xlsxwriter

def import_file():
    file_path = filedialog.askopenfilename(title="Select a file", filetypes=[("Python files", "*.pdf"), ("All files", "*.*")])

    if file_path:
        reader = PdfReader(file_path)

        newFilePath = file_path.replace('.pdf', '.xlsx')
        workbook = xlsxwriter.Workbook(newFilePath)
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
            messageMatches = ["Message", "of"]
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
                    messageText = messageText[:100]
                    wordsMessage.append(sentText)
                if all(x in line for x in fromMatches):
                    wordsMessage.append(line)
                if all(x in line for x in toMatches):
                    wordsMessage.append(line)
                if all(x in line for x in subjectMatches):
                    subjectLine = line.replace('RE: ', '').replace('Re: ', '')
                    wordsMessage.append(subjectLine)
                    wordsMessage.append(messageText)
                    fullPage.append(wordsMessage.copy())
                    wordsMessage.clear()

            fullPDF.extend(fullPage.copy())
            fullPage.clear()

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