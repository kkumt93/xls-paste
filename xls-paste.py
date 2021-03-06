import xlwings as xw
import sys

args = sys.argv

xls_file = args[1]
out_file = args[2]
file_list = args[3]

#Write  Excel
wb = xw.Book(xls_file)
for line_data in open(file_list, "r"):
    if line_data[0] == "#":
        continue
    split_line_data = line_data.split()

    sheets = split_line_data[1]
    offset_x = int(split_line_data[2]) + 1
    offset_y = int(split_line_data[3]) + 1

    y = 0
    for text_data in open(split_line_data[0], "r"):
        data = text_data.split()
        wb.sheets[sheets].range(offset_y + y, offset_x).value = data
        y += 1

#Save Excel
wb.save(out_file)