import xlwings as xw
import sys

args = sys.argv

xls_file = args[1]
out_file = args[2]
ini_file = args[3]

#Write  Excel
wb = xw.Book(xls_file)
for file_path in open(ini_file, "r"):
    if file_path[0] == "#":
        continue
    path = file_path.split()

    col = sum(1 for line in open(path[0])) - 1
    offset_y = int(path[2])
    offset_x = int(path[3])

    for y in range(col):
        x = 0
        for data_path in open(path[0], "r"):
            data = data_path.split()
            wb.sheets[path[1]].range(x+offset_x+1, y+offset_y+1).value = data[y]
            x=x+1
        
#Save Excel
wb.save(out_file)