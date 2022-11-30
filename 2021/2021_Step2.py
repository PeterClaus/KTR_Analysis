import collections
import os
import openpyxl

path1 = "/Users/w0z0341/Desktop/2021_Temp1"
path2 = "/Users/w0z0341/Desktop/2021_Temp2"
path3 = "/Users/w0z0341/Desktop/2021_Destination"

files1 = os.listdir(path1)
files2 = os.listdir(path2)

Map = collections.defaultdict(list)

for f in files1:
    if f.endswith(".xlsx") and f[0].isalpha():
        wb1 = openpyxl.load_workbook(path1 + "/" + f)
        ws1 = wb1["Sheet1"]

        for row in range(3, 10000):
            sec = ws1.cell(row, 9).value
            if sec is None:
                break
            Map[f[:3]].append(sec)
        wb1.close()

for K in Map:
    for f in files2:
        if f.startswith(K):
            wb2 = openpyxl.load_workbook(path2 + "/" + f)
            ws2 = wb2.active
            curList = Map[K]
            for i in range(len(curList)):
                ws2.cell(i+2, 2).value = curList[i]
            ws2.cell(1, 5).value = "C/N Ratio"
            for row in range(2, 10000):
                if ws2.cell(row, 6) is None:
                    break
                s = 0
                count = 0
                col = 6
                while col >= 6:
                    if ws2.cell(row, col).value is None:
                        break
                    C = ws2.cell(row, col + 1).value
                    N = ws2.cell(row, col).value
                    s += C/N
                    count += 1
                    col += 2
                if count == 0:
                    continue
                ws2.cell(row, 5).value = format(s/count, ".2f")
            wb2.save(path3 + "/" + f[:3] + "_result.xlsx")







