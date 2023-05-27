import PyPDF2
import openpyxl


pdfFileObj = open('1.pdf', 'rb')
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
pdfReader.numPages

pageObj = pdfReader.getPage(0)
mytext = pageObj.extractText()


emp=[]
wb = openpyxl.load_workbook('1.xlsx')
sheet = wb.active
sheet.title = 'MyPDF'

lists = list(mytext.split(" "))
print(lists)
while True:
    word=input("Enter the word:")
    if word in lists:
        emp.append(word)
        print(emp)
    else:
        print("Not in Pdf")
        exit()
    for i, j in enumerate(emp):
        sheet.cell(row=1+i,column=2).value = j
        

    
    sheet['A1'] = i
    
    wb.save('1.xlsx')
    print('DONE!!'  )
         

        