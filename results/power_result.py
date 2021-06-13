# encoding=utf-8
'''
@File    :   power_result.py    
@Contact :   jianbingxiaman@gmail.com
@License :   (C)Copyright 2020-2021, John Hopcraft Lab-CV
@Desciption : 
@Modify Time      @Author    @Version
------------      -------    --------
2021/5/9 19:36   jianbingxia     1.0    
'''
import logging
from collections import defaultdict

import pandas as pd

from results.base_result import BaseResult, HourData, OneData
from util import save_result


class PowerResultElement(object):
	xianluchaoliu_arc = {
		0: ['L1-0', 'L2-0'],
		1: ['L1-0', 'L3-1'],
		2: ['L2-0', 'L4-2'],
		3: ['L3-1', 'L5-3'],
		4: ['L4-2', 'L5-4'],
		5: ['L5-3', 'L5-4'],
	}

	def __init__(self, solver):
		self.solver = solver
		self.original_result = solver.matrixchange()

		self.node_ids = self.solver.TOTAL_NODE_ID

		self.has_result = set()
		self.fuhe = defaultdict(HourData)  # 负荷
		self.xianluchaoliu = defaultdict(list)  # 线路潮流：0.且末-若羌，1.罗布泊-若羌 [HourData]
		self.zhuangji = defaultdict(list)  # 装机，分别是火电装机、新能源装机、风电装机、光伏装机 [OneData]
		self.huodianchuli = defaultdict(HourData)  # 火电出力
		self.xinnengyuanchuli = defaultdict(HourData)  # 火电出力
		self.shudianxiangonglv = defaultdict(HourData)  # 输电线功率
		self.dianliyingyu = defaultdict(HourData)  # 电力盈余
		self.xinnengyuanqidian = defaultdict(HourData)  # 新能源弃电

		self.excel_df = {}
		self.scene_nums = self.solver.sceneNums
		for scene_num in self.scene_nums:
			excel = pd.read_excel(f'data/xlsx/算例-2020-丰水年-南疆电网-场景{scene_num}_GEN.xlsx', sheet_name=2)
			excel = excel.loc[(excel['Hyd'] == 4)
							  & ((excel['sID'] == 1) | (excel['sID'] == 3) | (excel['sID'] == 4))
							  & (5 <= excel['dID']) & (excel['dID'] <= 11)
							  & (501 <= excel['Flg']) & (excel['Flg'] <= 512)]

			self.excel_df[scene_num] = excel
		for node_id in self.node_ids:
			self.get_zhuangji(node_id)

		self.zhuangji_0124 = OneData()
		for node_id in [0, 1, 2, 4]:
			huodianzhuangji, xinnengyuanzhuanji, _, _ = self.zhuangji[node_id]
			self.zhuangji_0124 += huodianzhuangji + xinnengyuanzhuanji

		logging.debug(f'PowerResultElement Constructed')

	def __call__(self, *args, **kwargs):
		assert len(args) == 3 and args[0] in self.scene_nums and 1 <= args[
			1] <= 12, f'Expected PowerResultElement(scene_num,month,date_type) but got ({args})'
		if args not in self.has_result:
			self.get_data(*args)

		self.has_result.add(args)
		return self

	def get_data(self, scene_num, month, date_type=False):
		logging.debug(f'Get PowerResultElement Data ... ')
		for node_id in self.node_ids:
			self.get_fuhe(node_id, scene_num, month, date_type=date_type)
			self.get_xianluchaoliu(node_id, scene_num, month, date_type=date_type)
			self.get_huodianchuli(node_id, scene_num, month, date_type=date_type)
			self.get_xinnengyuanchuli(node_id, scene_num, month, date_type=date_type)
			self.get_shudianxiangonglv(node_id, scene_num, month, date_type=date_type)
			self.get_dianliyingyu(node_id, scene_num, month, date_type=date_type)
			self.get_xinengyuanqidian(node_id, scene_num, month, date_type=date_type)

	def get_fuhe(self, node_id, scene_num, month, date_type=False):
		if node_id == self.solver.OUTPUT_NODE_ID:
			power_demand = self.solver.statusDict[scene_num, month].powerDemand  # 外送曲线
			self.fuhe[node_id] = HourData(power_demand)
		else:
			self.fuhe[node_id] = HourData([0] * 24)

	def get_xianluchaoliu(self, node_id, scene_num, month, date_type=False):

		def __xianluchaoliu_arc(node_from_to, scene_num, month, date_type=False):
			if isinstance(node_from_to, list):
				for nft in node_from_to:
					__xianluchaoliu_arc(nft, scene_num, month, date_type)
				return

			power_on_arcs = self.original_result[scene_num, month].powerOnArcs
			for power_on_arc in power_on_arcs:
				if power_on_arc[1] == node_from_to:
					self.xianluchaoliu[node_id].append(HourData(power_on_arc[5:]))
					break

		if node_id not in self.node_ids:
			raise ValueError(f'Expected node_id in {self.node_ids}, but got {node_id}')

		__xianluchaoliu_arc(self.xianluchaoliu_arc[node_id], scene_num, month, date_type)

	def get_zhuangji(self, node_id):

		def _zhuangji(power_generators, types: [int]) -> int:
			return sum(
				[power_generator.zhuangji for power_generator in power_generators if
				 power_generator.type in types and node_id == power_generator.systemID])

		power_generators = self.solver.powerGenerators

		huodian_zhuangji = _zhuangji(power_generators, [101])
		fengdian_zhuangji = _zhuangji(power_generators, [200])
		guangfu_zhuangji = _zhuangji(power_generators, [300])

		xinnengyuan_zhuangji = fengdian_zhuangji + guangfu_zhuangji

		self.zhuangji[node_id] = [OneData(huodian_zhuangji), OneData(xinnengyuan_zhuangji), OneData(fengdian_zhuangji),
								  OneData(guangfu_zhuangji)]

	def get_huodianchuli(self, node_id, scene_num, month, date_type=False):

		thermalPowerOnNode = self.original_result[scene_num, month].thermalPowerOnNode
		for thermalPowerThisNode in thermalPowerOnNode:
			if node_id == thermalPowerThisNode[0]:
				self.huodianchuli[node_id] = HourData(thermalPowerThisNode[5:])

	def get_xinnengyuanchuli(self, node_id, scene_num, month, date_type=False):
		cleanPowerOnNode = self.original_result[scene_num, month].cleanPowerOnNode
		for cleanPowerThisNode in cleanPowerOnNode:
			if node_id == cleanPowerThisNode[0]:
				self.xinnengyuanchuli[node_id] = HourData(cleanPowerThisNode[5:])

	def get_shudianxiangonglv(self, node_id, scene_num, month, date_type=False):
		excel__ = self.excel_df[scene_num]
		excel__ = excel__.loc[excel__['Flg'] == month + 500]

		if not date_type:  # 优化前
			if node_id == 3:  # 和田 sID=1
				excel = excel__.loc[excel__['sID'] == 1]
				shudianxiangonglv = self.__shudianxiangonglv__(excel)

			elif node_id == 5:  # 阿克苏 sID=3
				excel = excel__.loc[excel__['sID'] == 3]
				shudianxiangonglv = self.__shudianxiangonglv__(excel)

			elif node_id in [0, 1, 2, 4]:  # 巴州 sID=4
				excel = excel__.loc[excel__['sID'] == 4]
				huodianzhuangji, xinnengyuanzhuanji, _, _ = self.zhuangji[node_id]
				zhuangji = huodianzhuangji + xinnengyuanzhuanji
				shudianxiangonglv = self.__shudianxiangonglv__(excel)
				shudianxiangonglv *= (zhuangji / self.zhuangji_0124).data[0]
			else:
				raise ValueError(f'Expected node_id in range(5), but got {node_id}')

			self.shudianxiangonglv[node_id] = shudianxiangonglv

	def __shudianxiangonglv__(self, excel):
		shudianxiangonglv = HourData()
		for line in excel.iloc[:, 6:30].values:
			hour_data = HourData(line)
			shudianxiangonglv += hour_data
		shudianxiangonglv /= 7
		return shudianxiangonglv

	def get_dianliyingyu(self, node_id, scene_num, month, date_type=False):
		abandonOnNode = self.original_result[scene_num, month].abandonOnNode
		for abandonThisNode in abandonOnNode:
			if node_id == abandonThisNode[0]:
				self.dianliyingyu[node_id] = HourData(abandonThisNode[5:])

	def get_xinengyuanqidian(self, node_id, scene_num, month, date_type=False):
		abandonOnNode = self.original_result[scene_num, month].abandonOnNode
		for abandonThisNode in abandonOnNode:
			if node_id == abandonThisNode[0]:
				self.xinnengyuanqidian[node_id] = HourData(abandonThisNode[5:])

	def to_excel(self, file_path):

		def append_line(df, line_datas: 'data or datas', node_id, project_names: 'name or names'):
			global project_id
			if isinstance(line_datas, list):
				for index, line_data in enumerate(line_datas):
					df = df.append(
						pd.DataFrame([[project_id, node_id, project_names[index]] + line_data.data], columns=columns))
					project_id += 1
			else:
				df = df.append(pd.DataFrame([[project_id, node_id, project_names] + line_datas.data], columns=columns))
				project_id += 1
			return df

		columns = ['项目编号', '节点', '项目', 'max', 'min', 'ave', '24小时总计'] + [f'{h}h' for h in range(1, 25)]
		df = pd.DataFrame(columns=columns)
		for node_id in self.node_ids:
			df = append_line(df, self.fuhe[node_id], node_id, '负荷')
			df = append_line(df, self.xianluchaoliu[node_id], node_id,
							 [f'线潮{xianlu}' for xianlu in self.xianluchaoliu_arc[node_id]])
			df = append_line(df, self.zhuangji[node_id], node_id, ['火电装机', '新能源装机', '风电装机', '光伏装机'])
			df = append_line(df, self.huodianchuli[node_id], node_id, '火电出力')
			df = append_line(df, self.xinnengyuanchuli[node_id], node_id, '新能源出力')
			df = append_line(df, self.shudianxiangonglv[node_id], node_id, '输电线功率')
			df = append_line(df, self.dianliyingyu[node_id], node_id, '电力盈余')
			df = append_line(df, self.xinnengyuanqidian[node_id], node_id, '新能源弃电')

		self.data = df

		save_result(df, file_path, sort_bys=['节点', '项目编号'])


class PowerResult(BaseResult):

	def __init__(self, solver, debug=False):
		BaseResult.__init__(self, solver, debug=debug)

	def __call__(self, *args, **kwargs):
		assert len(args) == 3 and isinstance(args[0], int) and isinstance(args[1], int) \
			   and 1 <= args[1] <= 12, f'Expected item is (sceneNum,month,dateType), but got {args}'
		if args not in self._data:
			power_result_element = PowerResultElement(self.solver)(*args)
			self._data[args] = power_result_element
		return self._data[args]

	def build_data(self, date_type=False):
		"""Do None"""
		logging.debug(f'Build Data PowerResult')
		for scene_num in self.solver.sceneNums:
			for month in range(1, 13):
				logging.info(f'Build {scene_num},{month},{date_type} PowerResult')
				self(scene_num, month, date_type)
				if self.debug:
					break
			if self.debug:
				break

	def __repr__(self):
		return f'{self._data}'

	def to_excel(self, scene_num, month, date_type=False, file_path="result0125.xlsx"):
		data = self(scene_num, month, date_type)
		data.to_excel(file_path)


project_id = 0
