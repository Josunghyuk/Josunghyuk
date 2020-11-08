#프로그래밍언어 프로그램 개발
import openpyxl

filename = "abc.xlsx"
filedata = openpyxl.load_workbook(filename)
detaildata = filedata.worksheets[0]

data = []
for row in detaildata.rows:
    data.append([
        row[0].value,
        row[1].value,
        row[2].value,
        row[3].value,
        row[4].value,
        row[5].value,
        row[6].value,
        row[7].value,
        row[8].value,
        row[9].value
    ])

print(data)
