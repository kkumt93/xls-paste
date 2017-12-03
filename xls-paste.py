import xlwings as xw
import sys

#Input Ini File
args = sys.argv

#Read Ini File
for line in open(args[1], "r"):
    if line[0] == "#":
        continue
    ini = line.split()
    
    if(ini[0]=="xls_file"):
        xls_file = ini[1]
    if(ini[0]=="sheets"):
        sheets = ini[1]
    if(ini[0]=="out_file"):
        out_file = ini[1]

#Write  Excel
wb = xw.Book(xls_file)
for file_path in open("file_list.txt", "r"):
    if file_path[0] == "#":
        continue
    path = file_path.split()

    #irow = int(path[2])
    icol     = int(path[3])
    offset_y = int(path[4])
    offset_x = int(path[5])

    for y in range(icol):
        x = 0
        for data_path in open(path[0], "r"):
            data = data_path.split()
            wb.sheets[path[1]].range(x+offset_x+1, y+offset_y+1).value = data[y]
            x=x+1

#Save Excel
wb.save(out_file)