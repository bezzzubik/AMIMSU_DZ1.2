import openpyxl as oxl

excel_file = 'NU.xlsx'
try:
    df = oxl.load_workbook(excel_file)
except FileNotFoundError:
    print(f"File not found")
    exit()

sheet = df.active
mass=[]

for row in sheet.iter_rows(min_row=1, max_row = 15, values_only=True):
    s = list(row)
    for i in range(15):
        s[i] = float(s[i])
    mass.append(s)

mass.insert(0, [0, 1,2,3,4,5,6,7,8,9,10,11,12,13,14,15])

for i in range(1,16):
    mass[i].insert(0, i)

for row in mass:
    for el in row:
        print(f"{el:7.2f}", end ="")
    print()
print()
print()

Ig = 1


def progg():
    for i in range(1, 16):

        for j in range(i+1, 16):
            if mass[j][i] != 0:
                #Перемещаем до ближайшего ненулевого значения или до диагонального
                lj = j
                while i != lj and mass[lj-1][i] == 0:
                    for li in range(16):
                        a = mass[li][lj]
                        mass[li][lj] = mass[li][lj-1]
                        mass[li][lj-1] = a
                    for li in range(16):
                        a = mass[lj][li]
                        mass[lj][li] = mass[lj-1][li]
                        mass[lj-1][li] = a
                    lj -= 1


        for j in range(i+1, 16):
            if mass[i][j] != 0:
                #Перемещаем до ближайшего ненулевого значения или до диагонального
                lj = j
                while i != lj and mass[i][lj-1] == 0:
                    for li in range(16):
                        a = mass[li][lj]
                        mass[li][lj] = mass[li][lj-1]
                        mass[li][lj-1] = a
                    for li in range(16):
                        a = mass[lj][li]
                        mass[lj][li] = mass[lj-1][li]
                        mass[lj-1][li] = a
                    lj -= 1
    return


def print_mass():
    for row in mass:
        for el in row:
            print(f"{el:7.2f}", end ="")
        print()
    print()
    print()
    return


progg()
#print(f"MASSIV {k:f} :")
print("-----------------------------------------------------------------------------")
print_mass()
print("-----------------------------------------------------------------------------")


sheet = df.create_sheet("Final")
for row, itm in enumerate(mass, start=1):
    for column, item in enumerate(itm, start=1):
        sheet.cell(row = row, column = column).value = item

df.save("Itog.xlsx")
