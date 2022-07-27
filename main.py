from openpyxl import load_workbook

reader = load_workbook('JULY.xlsx')
sheet = reader.active
with open('data.txt','r') as f:
    # print(sheet['A1'].value,sheet['B1'].value,sheet['C1'].value,sheet[f'D1'].value,sheet['E1'].value)
    for cell in range(1,98):
            datatoenter = f.readline().split(',')
            # print(sheet[f'C{cell}'].value)
            for i in range(2,99):
                if sheet[f'A{i}'].value == int(datatoenter[0]):
                    sheet[f'B{i}'].value = datatoenter[1]
                    sheet[f'D{i}'].value = int(datatoenter[2])
                    sheet[f'E{i}'].value = int(datatoenter[3])
                    sheet[f'F{i}'].value = int(datatoenter[4])
                    sheet[f'H{i}'].value = int(datatoenter[5])
                    sheet[f'I{i}'].value = sheet[f'G{i}'].value-sheet[f'H{i}'].value
                    sheet[f'J{i}'].value = sheet[f'D{i}'].value+sheet[f'E{i}'].value+sheet[f'F{i}'].value+sheet[f'H{i}'].value


for j in range(1,99):
    print(sheet[f'A{j}'].value,sheet[f'B{j}'].value,sheet[f'C{j}'].value,sheet[f'D{j}'].value,sheet[f'E{j}'].value,sheet[f'F{j}'].value,sheet[f'G{j}'].value,sheet[f'H{j}'].value,sheet[f'I{j}'].value,sheet[f'J{j}'].value)
            #     sheet[f'D{cell}'].value = 100
            #     print(sheet[f'C{cell}'].value,sheet[f'D{cell}'].value )
            # print(sheet[f'A{cell}'].value,sheet[f'B{cell}'].value,sheet[f'C{cell}'].value,sheet[f'D{cell}'].value)
# sheet['A2'].value = 100
sheet['D99'].value="=sum(D2:D98)"
sheet['E99'].value="=sum(E2:E98)"
sheet['F99'].value="=sum(F2:F98)"
sheet['G99'].value="=sum(G2:G98)"
sheet['H99'].value="=sum(H2:H98)"
sheet['I99'].value="=sum(I2:I98)"
sheet['J99'].value="=sum(J2:J98)"

print(sheet['D99'].value,sheet['E99'].value,sheet['F99'].value,sheet['G99'].value,sheet['H99'].value,sheet['I99'].value,sheet['J99'].value)

reader.save('test.xlsx')