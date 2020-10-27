import pandas as pd
import math
import openpyxl

cidade = {}

file = "Mapa de Oportunidades - Brasíla - 11.09.2020.xlsb"
print("lendo planilha...")
sheet = pd.read_excel(file, engine='pyxlsb',sheet_name = "BASE INFORMAÇÕES", skiprows=16, usecols=["BAIRRO / CIDADE", 'HPs CADASTRADOS', 'INSTALADOS', ' HP LIVRE'])

print("estruturando dados...")
for index, row in sheet.iterrows():
    if row["BAIRRO / CIDADE"] not in cidade:
        cidade[row["BAIRRO / CIDADE"]] = {}
    if "HPs CADASTRADOS" not in cidade[row["BAIRRO / CIDADE"]]:
        cidade[row["BAIRRO / CIDADE"]]["HPs CADASTRADOS"] = 0
    if "INSTALADOS" not in cidade[row["BAIRRO / CIDADE"]]:
        cidade[row["BAIRRO / CIDADE"]]["INSTALADOS"] = 0
    if "HP LIVRE" not in cidade[row["BAIRRO / CIDADE"]]:
        cidade[row["BAIRRO / CIDADE"]]["HP LIVRE"] = 0
        

    cidade[row["BAIRRO / CIDADE"]]["HPs CADASTRADOS"] += row["HPs CADASTRADOS"]
    if math.isnan(row["INSTALADOS"]) == False:
        cidade[row["BAIRRO / CIDADE"]]["INSTALADOS"] += row["INSTALADOS"]
    cidade[row["BAIRRO / CIDADE"]]["HP LIVRE"] += row[" HP LIVRE"]


for key, value in cidade.items():
    percentage = value["INSTALADOS"] * 100 / value["HPs CADASTRADOS"]
    cidade[key]["penetracao"] = str(percentage) + "%"

cidade_sorted = dict(sorted(cidade.items(), key=lambda k: k[1]['HPs CADASTRADOS'], reverse=True))

print("criando arquivo excel...")
# Create the workbook and sheet for Excel
workbook = openpyxl.Workbook()
new_sheet = workbook.active

new_sheet.cell(row = 1, column = 1, value="BAIRRO / CIDADE")
new_sheet.cell(row = 1, column = 2, value="HPs CADASTRADOS")
new_sheet.cell(row = 1, column = 3, value="INSTALADOS")
new_sheet.cell(row = 1, column = 4, value="HP LIVRE")

# openpyxl does things based on 1 instead of 0
row = 2
for key,values in cidade_sorted.items():
    # Put the key in the first column for each key in the dictionary
    new_sheet.cell(row=row, column=1, value=key)
    column = 2
    for element in values.values():
        # Put the element in each adjacent column for each element in the tuple
        new_sheet.cell(row=row, column=column, value=element)
        column += 1
    row += 1

workbook.save(filename="feedback.xlsx")

print("Concluido !!")