import os
from openpyxl import load_workbook
from datetime import datetime

current_datetime = datetime.now().strftime("%d-%m-%Y %H-%M-%S")
str_current_datetime = str(current_datetime)

filepath = os.getenv('USERPROFILE')

workbook_hours = load_workbook(filepath + "\\Desktop\\" + "Gerador\\Adicional Noturno.xlsx", data_only=True)

workbook_base = load_workbook(filepath + "\\Desktop\\" + "Gerador\\Planilha_Base_AN.xlsx")

worksheet_base = workbook_base ["Adicional Noturno"]

sheets = workbook_hours.sheetnames

if not os.path.exists(filepath + "\\Desktop\\" + "Gerador\\Gerados"):
    os.mkdir(filepath + "\\Desktop\\" + "Gerador\\Gerados")
else:
    print ("A pasta já existe")

count = 0

for worksheet_hours in workbook_hours.worksheets[0:1]:
    os.system('cls')
    print ("\nProcessando Informações da planilha", worksheet_hours.title,"\n")    
    for registration_cell, add_cell in zip (worksheet_hours['A:A'],worksheet_hours['C:C']):
        if registration_cell.value is not None:
            if add_cell.value == "Sim":
                registration = worksheet_hours.cell(row=registration_cell.row, column=1).value
                name = worksheet_hours.cell(row=registration_cell.row, column=2).value
                session = worksheet_hours.cell(row=registration_cell.row, column=4).value
                time_convert = worksheet_hours.cell(row=registration_cell.row, column=7).value
                decimal_minutes = ((time_convert * 60) / 100)
                worksheet_base.cell(row=count+5,column=1).value = registration
                worksheet_base.cell(row=count+5,column=2).value = name
                worksheet_base.cell(row=count+5,column=3).value = session
                worksheet_base.cell(row=count+5,column=7).value = decimal_minutes
                count += 1
    
    workbook_base.save(filepath + "\\Desktop\\" + "Gerador\\Gerados\\Tabela Adicional Noturno Professores" + "(" + str_current_datetime + ")" + ".xlsx")

input("\nPressione qualquer tecla para encerrar a aplicação...")