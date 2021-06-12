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

from results.base_result import BaseResult, MonthData


class ElectricityResultElement(object):

	def __init__(self, solver, power_result, debug=False):
		self.solver = solver
		self.debug = debug
		self.month_status = solver.month_status
		self.node_ids = solver.TOTAL_NODE_ID
		self.scene_nums = solver.sceneNums
		self.power_result = power_result

		self._data = {
			'系统电量需求': defaultdict(),
			'火力发电': defaultdict(),
			'新能源发电': defaultdict(),
			'输电线送电': defaultdict(),
			'新能源弃电': defaultdict(),
		}  # 还有联外{线路}:defaultdict()

		self.columns = ['节点', '项目', '合计'] + [f'{month}月' for month in range(1, 13)]

		self.data = pd.DataFrame(columns=self.columns)
		logging.debug(f'ElectricityResultElement Construced')

	def __call__(self, *args, **kwargs):
		assert len(args) <= 1, f'Expected ElectricityResultElement(date_type) but got ({args})'
		date_type = False
		if len(args) == 1:
			date_type = args[0]
		self.get_data(date_type)

		return self

	def get_data(self, date_type=False):
		logging.debug('Get ElectricityResult Data ...')
		for node_id in self.node_ids:
			self.get_base(node_id, 'fuhe', '系统电量需求', date_type)
			self.get_base(node_id, 'xianluchaoliu', '联络线外送', date_type)
			self.get_base(node_id, 'huodianchuli', '火力发电', date_type)
			self.get_base(node_id, 'xinnengyuanchuli', '新能源发电', date_type)
			self.get_base(node_id, 'shudianxiangonglv', '输电线送电', date_type)
			self.get_base(node_id, 'xinnengyuanqidian', '新能源弃电', date_type)

	def get_base(self, node_id, from_name, to_name, date_type=False):
		logging.debug(f'Get {to_name} for {node_id} in {date_type}')

		def get_one_base(index=0, project_name=None):
			if project_name is None:
				project_name = to_name

			month_data = self.get_month_data(from_name, node_id, index, date_type)
			self._data[project_name][node_id] = month_data
			self.data = self.data.append(
				pd.DataFrame([[node_id, project_name] + month_data.data], columns=self.columns))

		if from_name == 'xianluchaoliu':
			sample_power_result = self.power_result(0, 1, date_type)
			len_data = len(sample_power_result.xianluchaoliu_arc[node_id])

			for index in range(len_data):
				xianlu = sample_power_result.xianluchaoliu_arc[node_id][index]
				project_name = f'联外{xianlu}'
				if project_name not in self._data:
					self._data[project_name] = defaultdict()

				get_one_base(index, project_name=project_name)
		else:
			get_one_base()

	def get_month_data(self, attr_name, node_id, index=0, date_type=False):
		month_data = []
		for month in range(1, 13):
			d_sum = 0
			for scene_num in range(len(self.scene_nums)):
				power_result_element = self.power_result(scene_num, month, date_type)
				if hasattr(power_result_element, attr_name):
					data = getattr(power_result_element, attr_name)
					hour_data = data[node_id]
					if isinstance(hour_data, list):
						if index > len(hour_data):
							raise ValueError(
								f'Expected index < len(attr_name), but got index={index},len(attr_name) = {len(hour_data)}')
						hour_data = hour_data[index]

					d_sum += hour_data.get_sum() * 30 * self.month_status[month][scene_num]
				else:
					raise ValueError(f'Expected attr in power_result_element, but got {attr_name}')

			month_data.append(d_sum)

		return MonthData(month_data)

	def to_excel(self, file_path):
		self.data.sort_values(by=['节点', '项目'], ascending=True, inplace=True)
		self.data.to_csv(file_path, index=False, encoding='utf_8_sig')
		self.data.to_excel(file_path.replace('.csv', '.xlsx'), index=False)
		logging.info(f'Saved {file_path}')


class ElectricityResult(BaseResult):

	def __init__(self, solver, power_result, debug=False):
		BaseResult.__init__(self, solver, debug=debug)
		self.power_result = power_result
		self.build_data()
		logging.debug(f'ElectricityResult Construced')

	def __call__(self, *args, **kwargs):
		assert len(args) == 1, f'Expected item is (date_type), but got {args}'
		if args[0] not in self._data:
			self._data[args[0]] = ElectricityResultElement(self.solver, self.power_result)(args[0])

		return self._data[args[0]]

	def build_data(self, date_type=False):
		logging.debug(f'Build Data ElectricityResult')
		self(date_type)

	def __repr__(self):
		return f'{self._data}'

	def to_excel(self, date_type=False, file_path="electricity_result.csv"):
		data = self(date_type)
		data.to_excel(file_path)
