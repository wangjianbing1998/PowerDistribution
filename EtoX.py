import clr
clr.FindAssembly("HUST_ProS_InData.dll")
from HUST_ProS_InData import *

instance = HUST_ProS_InData()
xlsFile = 'C:/Users/zhang/Desktop/调试用/Hust数据/南疆电网.xlsx'
xlsFile = xlsFile.replace("/","\\")
exePath = 'D:/HustPros_Setup/'
exePath = exePath.replace("/","\\")
msg = ''

print(xlsFile)
print(exePath)
result = instance.HUST_ProS_InData_Excel2XML(xlsFile,exePath,msg)
print(result)