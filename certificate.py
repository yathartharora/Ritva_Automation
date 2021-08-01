from docxtpl import DocxTemplate
import openpyxl

wb = openpyxl.load_workbook('./Gargi.xlsx')
sheet = wb.active



def edit(x):
    tmpl = DocxTemplate('./temp.docx')
    context = {
        "Name": x
    }
    tmpl.render(context)
    tmpl.save('./EarthlyArtsy/' + x + '.docx')


for i in range(45,56):
    cell = sheet.cell(row=i,column=1)
    edit(cell.value)
