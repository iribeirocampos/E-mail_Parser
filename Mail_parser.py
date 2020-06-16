import csv
import re
import xlsxwriter

workbook = xlsxwriter.Workbook("teste.xlsx")
worksheet = workbook.add_worksheet()
string = ""

with open("teste.csv", "r") as csv_file:
    csv_reader = csv.reader(csv_file)

    for lines in csv_reader:
        for line in lines:
            string += line

inicios = []

[inicios.append(m.start()) for m in re.finditer("Enviada", string)]

row = 0
col = 0
for data in inicios:
    data = string[data + 9 : data + 37]
    data = data.replace("de", ",")
    dia, mes, ano = data.split(",")

    try:
        dia = int(dia)
        worksheet.write(row, col + 2, dia)
        worksheet.write(row, col + 1, mes)
        hora = ano[5:11]
        ano = ano[1:5]
        ano = int(ano)
        worksheet.write(row, col, ano)
        worksheet.write(row, col + 3, hora)
    except Exception:
        pass
    row += 1
workbook.close()
