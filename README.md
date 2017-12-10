# xls-paste

Qiita

https://qiita.com/kkumt93/items/47650cb3c7db58624c04

# Requirements

- Python 3.5.1
- xlwings 0.11.4
- pywin32 221

## Installation

**xlwings**

http://docs.xlwings.org/en/stable/installation.html#installation

# Usage

```
python xls-paste.py template.xls out.xls file_list.txt
```

## template.xls

You set the empty graph area on Excel sheets.

![img1](https://camo.qiitausercontent.com/3f34575e2d3730ad2f40d317e5e4101843670e98/68747470733a2f2f71696974612d696d6167652d73746f72652e73332e616d617a6f6e6177732e636f6d2f302f39373833332f30656164303364322d313334612d313934362d626165362d6439333064366534666262652e706e67)

## out.xls

This Argument is unique output file name.

## file_list.txt

### Format
```
#data path          sheet-name  offsetX offsetY
./data/data1.txt    Sheet1      0       0
./data/data2.txt    Sheet2      0       0
./data/data1.txt    Sheet3      0       0
./data/data2.txt    Sheet3      7       0
```

### data1.txt

```
*   10  20  30  40  50
A   1   2   3   4   5
B   6   7   8   9   10
C   11  12  13  14  15
D   16  17  18  19  20
E   21  22  23  24  25
```

## Execution result
![img2](https://camo.qiitausercontent.com/ac60ebac84ca2458c022ed4b4fe989f1f711df14/68747470733a2f2f71696974612d696d6167652d73746f72652e73332e616d617a6f6e6177732e636f6d2f302f39373833332f39333534336562642d326266382d626433382d633633312d3735616331323462616632322e706e67)
