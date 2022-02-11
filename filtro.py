from openpyxl import Workbook, load_workbook
from dictionario import dictionary
import time


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

# Separate the city from the state key 
def handleCity(fullCity):
    newArray = fullCity.split(',')
    return newArray

# Look for the corresponding state according to its key
def handleSatet(keyState):
    keydictionary = dictionary.keys()
    print(keydictionary)
    if keyState in keydictionary:
        print(f'si esta {keyState}')
    else:
        print(f'No lo encontramos {keyState}')
        
    # print(dictionary)

while contador <= columnas:
    fullCity = sheet[f'A{contador}'].value
    fullCity = handleCity(fullCity)
    nameState = handleSatet(fullCity[1])
    # print(fullCity[0])
    new_excel.worksheets[0][f'A{contador}'].value = fullCity[0]
    new_excel.worksheets[0][f'B{contador}'].value = fullCity[1]
    new_excel.worksheets[0][f'C{contador}'].value = fullCity[1]
    new_excel.worksheets[0][f'D{contador}'].value = sheet[f'B{contador}'].value
    new_excel.worksheets[0][f'E{contador}'].value = sheet[f'C{contador}'].value
    new_excel.worksheets[0][f'F{contador}'].value = 'Mexico'
    # print(contador)
    contador += 1
    time.sleep(2)
    
# new_excel.save(f'phone_clean_data.xlsx')


