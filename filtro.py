from openpyxl import Workbook, load_workbook
from dictionario import dictionary
import time


print(dictionary['QRO'])
time.sleep(5)
print(dictionary['QROO'])

def initial_excel():
    new_excel = Workbook()
    
    # Add Headers
    new_excel.worksheets[0]['A1'].value = 'Ciudad'
    new_excel.worksheets[0]['B1'].value = 'Estado'
    new_excel.worksheets[0]['C1'].value = 'clave_Estado'
    new_excel.worksheets[0]['D1'].value = 'Lada'
    new_excel.worksheets[0]['E1'].value = 'Longitud'
    new_excel.worksheets[0]['F1'].value = 'Pais'
    return new_excel


phone_data = load_workbook('telefonia.xlsx')
sheet = phone_data['Hoja1']
columnas = sheet.max_row
contador = 2
new_excel = initial_excel()


def handleCity(full_city):
    newArray = full_city.split(',')
    return newArray
    

while contador <= columnas:
    fullCity = sheet[f'A{contador}'].value
    fullCity = handleCity(fullCity)
    print(fullCity[0])
    new_excel.worksheets[0][f'A{contador}'].value = fullCity[0]
    new_excel.worksheets[0][f'B{contador}'].value = fullCity[1]
    new_excel.worksheets[0][f'C{contador}'].value = fullCity[1]
    new_excel.worksheets[0][f'D{contador}'].value = sheet[f'B{contador}'].value
    new_excel.worksheets[0][f'E{contador}'].value = sheet[f'C{contador}'].value
    new_excel.worksheets[0][f'F{contador}'].value = 'Mexico'
    print(contador)
    contador += 1
    
# new_excel.save(f'phone_clean_data.xlsx')


