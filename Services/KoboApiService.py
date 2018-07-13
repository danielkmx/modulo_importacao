import xlrd


#Abrir o arquivo XLS

workbook = xlrd.open_workbook("Resultado da Enquete por LABELS.xlsx")

#Abrir a aba
worksheet = workbook.sheet_by_name("ENQUETE_DESIGNAÇÃO_RJ")

total_linhas = worksheet.nrows
total_colunas = worksheet.ncols



table = list()
record = list()

for x in range(total_colunas):
    for y in range(total_linhas):
        if worksheet.cell(y, x).value:
            record.append(worksheet.cell(y, x).value)
    table.append(record)
    record = []


for x in table:
    print(x)
