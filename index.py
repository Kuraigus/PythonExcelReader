import pandas as pd


# usecols=["BAIRRO / CIDADE", 'HPs\nCADASTRADOS', '% \nPENETRAÇÃO ']


file = "Mapa de Oportunidades - Brasíla - 11.09.2020.xlsb"
print("lendo planilha...")
sheet = pd.read_excel(file, engine='pyxlsb',sheet_name = "BASE INFORMAÇÕES", skiprows=15)
print("resultado:")
columns = sheet.columns.ravel()

print(columns[8])
print(columns[9])
print(columns[10])
print(columns[11] * 100)
