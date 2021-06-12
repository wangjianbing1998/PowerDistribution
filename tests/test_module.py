import xlrd

import solver

test = xlrd.open_workbook("../data/算例.xlsx")
sheet = test.sheet_by_name("线路表")
print(sheet.row_values(1)[15], sheet.row_values(1)[10], sheet.row_values(1)[11])

s = solver.Solver()

# 打印出所有的场景信息
"""for i in range(0, 12):
    print(s.status_99[i])"""

# 打印出所有的线路结构信息
for i in s.arcs:
    print(i)

# 打印出所有的线路节点信息
for i in s.nodeDict:
    print(i)
print(s.sceneNums)
