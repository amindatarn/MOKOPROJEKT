import xlrd

sheet = xlrd.open_workbook(r"C:\Users\AmindaTärn\Desktop\Python\Min data\BA\AUTOGYRONbalconyKOPIA.xlsx").sheet_by_index(0)

print(sheet.cell_value(20, 6))

i = 1
while i < (sheet.nrows):
    x = (sheet.cell_value(i, 5)).lower()
    print(x)

    i += 1

