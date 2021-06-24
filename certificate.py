from docx import Document
import openpyxl
from docx2pdf import convert

loc = "./Details.xlsx"

wb = openpyxl.load_workbook(loc)
sheet = wb.active
# print(sheet.max_row)
# print(sheet.cell(row = 3, column =2).value)


for i in range(4,sheet.max_row+1):
    document = Document('./template.docx')
    for paragraph in document.paragraphs:
        if paragraph.text == 'Dearest $$$$$,':
            paragraph.text = paragraph.text.replace('$$$$$',str(sheet.cell(row=i,column=2).value))
            document.save('./Certificates/' + sheet.cell(row=i,column=2).value + '.docx')
            break
            

