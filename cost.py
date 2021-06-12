#coding=utf-8
#加载模块
#from EtoX import writeInfoToXml
import sys
import re
import xlwt
import xlrd
import xlutils
from xlutils.copy import copy
import os
import win32com.client as win32
from win32com.client import Dispatch
import xlwings

import Lin

#from pandas.core.frame import DataFrame
#import pandas as pd

#该程序用来计算投资费用、运行费用指标，有风光水火投资成本

class cost():
	def __init__(self):
		super().__init__()
		def search(path=".", name=""):  #search函数找到想要的几个方案的XML文档添加到result中
			for item in os.listdir(path):
				item_path = os.path.join(path, item)
				if os.path.isdir(item_path):
					search(item_path, name)
				elif os.path.isfile(item_path):
					if name in item:
						global result
						result.append(item_path)
		global result
		result = []

		if Lin.ResFangAn == "0":
			FangAnname = "全部方案"
			for i in range(len(Lin.huangbo)):  # 读取方案编号
				search(path=Lin.filepath, name=Lin.filename[0:-4] + str(Lin.huangbo[i]) + '_RST.xml')  #
		else:
			FangAnname = '方案' + str(Lin.ResFangAn)
			search(path=Lin.filepath, name=Lin.filename[0:-4] + str(Lin.ResFangAn) + '_RST.xml')

		self.book = xlwt.Workbook()
		for ii in range(len(result)):
			if Lin.ResFangAn == "0":
				x = self.book.add_sheet('方案' + str(Lin.huangbo[ii]), cell_overwrite_ok=True)  # 改sheeet名字为方案几，并赋给x
				biaotou = '方案' + str(Lin.huangbo[ii])
			else:
				biaotou = FangAnname
				x = self.book.add_sheet(FangAnname, cell_overwrite_ok=True)

			title1 = ['总投资成本（百万元）','年运行成本（百万元）', '总成本（百万元）', '发电煤耗（万吨标煤）', '煤耗成本（百万元）', '年排污费（百万元）']
			title2 = ['风电投资成本（百万元）', '光伏投资成本（百万元）', '水电投资成本（百万元）', '火电投资成本（百万元）','储能投资成本（百万元）']
			for j in range(len(title1)):  # 把上面的表头名写入表格
				x.write(0, j+1, title1[j])
			for j in range(len(title2)):  # 把上面的表头名写入表格
				x.write(2, j+1, title2[j])
			x.write(1, 0, biaotou)


			filewz1 = result[ii]
			f1 = open(filewz1, "r", encoding='UTF-8')  #以读的方式打开XML文档
			content1 = f1.read()
			f1.close()

			#新建一些空集来存放读取的XML文档的列内容
			list1 = []
			list2 = []
			list3 = []
			list4 = []
			list5 = []
			list6 = []
			list7 = []
			list8 = []
			list9 = []
			list10 = []
			list11 = []
			list12 = []
			list13 = []
			list14 = []
			list15 = []
			list16 = []
			list17 = []
			list18 = []
			list19 = []
			list20 = []
			list21 = []
			list22 = []
			list23 = []
			list24 = []
			list25 = []
			list26 = []
			list27 = []
			list28 = []
			list29 = []

			channelList = re.findall("<TEC.*?</TEC>", content1, re.S)
			for j1 in channelList:
				list1.append(re.findall("<Prj>(.*?)</Prj>", j1, re.S)[0])  # 把TEC页中对应名称的列赋值给对应的list
				list2.append(re.findall("<Yrs>(.*?)</Yrs>", j1, re.S)[0])
				list29.append(re.findall("<Hyd>(.*?)</Hyd>", j1, re.S)[0])
				list3.append(re.findall("<sID>(.*?)</sID>", j1, re.S)[0])
				list4.append(re.findall("<dID>(.*?)</dID>", j1, re.S)[0])
				list5.append(re.findall("<Flg>(.*?)</Flg>", j1, re.S)[0])
				list6.append(re.findall("<T1>(.*?)</T1>", j1, re.S)[0])
				list7.append(re.findall("<T2>(.*?)</T2>", j1, re.S)[0])
				list8.append(re.findall("<T3>(.*?)</T3>", j1, re.S)[0])
				list9.append(re.findall("<T4>(.*?)</T4>", j1, re.S)[0])
				list10.append(re.findall("<T5>(.*?)</T5>", j1, re.S)[0])
				list11.append(re.findall("<T6>(.*?)</T6>", j1, re.S)[0])
				list12.append(re.findall("<T7>(.*?)</T7>", j1, re.S)[0])
				list13.append(re.findall("<T8>(.*?)</T8>", j1, re.S)[0])
				list14.append(re.findall("<T9>(.*?)</T9>", j1, re.S)[0])
				list15.append(re.findall("<T10>(.*?)</T10>", j1, re.S)[0])
				list16.append(re.findall("<T11>(.*?)</T11>", j1, re.S)[0])
				list17.append(re.findall("<T12>(.*?)</T12>", j1, re.S)[0])
				list18.append(re.findall("<T13>(.*?)</T13>", j1, re.S)[0])
				list19.append(re.findall("<T14>(.*?)</T14>", j1, re.S)[0])
				list20.append(re.findall("<T15>(.*?)</T15>", j1, re.S)[0])
				list21.append(re.findall("<T16>(.*?)</T16>", j1, re.S)[0])
				list22.append(re.findall("<T17>(.*?)</T17>", j1, re.S)[0])
				list23.append(re.findall("<T18>(.*?)</T18>", j1, re.S)[0])
				list24.append(re.findall("<T19>(.*?)</T19>", j1, re.S)[0])
				list25.append(re.findall("<T20>(.*?)</T20>", j1, re.S)[0])
				list26.append(re.findall("<T21>(.*?)</T21>", j1, re.S)[0])
				list27.append(re.findall("<T22>(.*?)</T22>", j1, re.S)[0])
				list28.append(re.findall("<T23>(.*?)</T23>", j1, re.S)[0])

			flg='C01'    #行标识符为C01，代表电源总计
			invcost=[]    #总投资等年值
			#gudingcost=[] #年固定费用
			kebiancost=[] #年可变费用
			paiwucost=[]  #年排污费用
			totalcost=[]  #总费用=投资+运行
			meihao=[]#
			meihaocost=[]
			oprcost=[]
			for i1 in range(len(list1)):
				if float(list3[i1]) == 0 and float(list4[i1]) == 1 and list5[i1] ==flg:  #索引到系统（0）年总计（1）电源总计（C01）这一行的行数i1
					invcost.append(float(list10[i1]))
					#gudingcost.append(float(list11[i1]))
					kebiancost.append(float(list12[i1]))
					paiwucost.append(float(list13[i1]))
					# totalcost.append(float(list14[i1]))
					meihao.append(float(list18[i1]))
					meihaocost.append(float(list18[i1])*7)#煤单价700元/吨
			# oprcost =[a+b+c for a, b, c in zip(gudingcost,kebiancost,paiwucost)]  #年运行费用
			oprcost = [a + b for a, b in zip( kebiancost, paiwucost)]  # 年运行费用

			#读装机容量
			self.book1 = xlrd.open_workbook(Lin.filepath + '方案总表.xls', 'w+')  # 以读的方式打开
			xx = self.book1.sheet_by_index(0)
			nrow = xx.nrows#总行数
			ncol = xx.ncols#总列数
			Finvcost=[]
			Ginvcost = []
			Sinvcost = []
			Hinvcost = []
			Cinvcost = []#储能成本

			for i in range(1, nrow):#不要开始标题行
				if float(xx.cell(i, 0).value) == float(Lin.huangbo[ii]):
					Finvcost.append(float(xx.cell(i, 5).value)*7*0.1175) #百万元/MW
					Ginvcost.append(float(xx.cell(i, 6).value) * 6*0.1102)  # 百万元/MW
					Sinvcost.append(float(xx.cell(i, 7).value) * 7*0.1061)  # 百万元/MW
					Hinvcost.append(float(xx.cell(i, 8).value) * 4*0.1175)  # 百万元/MW
					Cinvcost.append(float(xx.cell(i, 9).value) * 2.5*0.2296)  # 百万元/MW

			invcost=[a + b +c+d+e for a, b ,c,d,e in zip(Finvcost, Ginvcost,Sinvcost,Hinvcost,Cinvcost)]
			totalcost = [a + b  for a, b in zip(invcost,oprcost)]

			#开始写表
			x.write(3, 1, Finvcost[0])
			x.write(3, 2, Ginvcost[0])
			x.write(3, 3, Sinvcost[0])
			x.write(3, 4, Hinvcost[0])
			x.write(3, 5, Cinvcost[0])

			x.write(1, 1, invcost[0])
			x.write(1, 2, oprcost[0])
			x.write(1, 3, totalcost[0])
			x.write(1, 4, meihao[0])
			x.write(1, 5, meihaocost[0])
			x.write(1, 6, paiwucost[0])

		self.book.save(Lin.filepath1 +  '云南系统' + str(Lin.shuipingnian)+'年' +FangAnname+'指标统计表' + '.xls')  #保存表格

