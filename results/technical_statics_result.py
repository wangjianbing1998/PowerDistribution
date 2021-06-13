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

import pandas as pd

from results.base_result import BaseResult, OneData
from util import save_result


class TechnicalStaticsResultElement(object):

	def __init__(self, solver, power_result, electricity_result, date_type=False):
		self.data = {
			'项目编号': [],  # [OneData()]
			'项目': ['火电', '新能源', '风电', '光伏', '合计'],  # [OneData()]
			'装机': [0 for _ in range(5)],  # [OneData()]
			'投资': [0 for _ in range(5)],
			'煤耗成本': [0 for _ in range(5)],
			'年运行成本': [0 for _ in range(5)],
			'总发电量': [0 for _ in range(5)],
			'弃电量': [0 for _ in range(5)],
		}
		self.data['项目编号'] = list(range(len(self.data['项目'])))

		self.solver = solver
		self.power_result_element = power_result(0, 1, date_type)
		self.electricity_result_element = electricity_result(date_type)

		self.date_type = date_type
		self.get_data()

		logging.debug(f'TechnicalStaticsResultElement Construced')

	def get_data(self):
		logging.debug(f'Get TechnicalStaticsResultElement Data ...')
		self.get_zhuangji()
		self.get_touzi()
		self.get_ranhaochengben()
		self.get_nianyunxingchengben()
		self.get_zongfadianliang()
		self.get_qidianliang()

	def get_zhuangji(self):
		zhuangji = [OneData(), OneData(), OneData(), OneData()]

		for node_id in self.solver.TOTAL_NODE_ID:
			huodian_zhuangji, xinnengyuan_zhuangji, fengdian_zhuangji, guangfu_zhuangji = \
				self.power_result_element.zhuangji[
					node_id]
			zhuangji[0] += huodian_zhuangji
			zhuangji[1] += xinnengyuan_zhuangji
			zhuangji[2] += fengdian_zhuangji
			zhuangji[3] += guangfu_zhuangji

		zhuangji.append(zhuangji[0] + zhuangji[1])
		self.data['装机'] = [one_data.data[0] for one_data in zhuangji]

	def get_touzi(self):
		huodian_zhuangji, _, fengdian_zhuangji, guangfu_zhuangji, _ = self.data['装机']

		def _get_HxL_(types: [int]) -> int:
			return sum(
				[power_generator.dongtaitouzi * power_generator.totalNums for power_generator in power_generators if
				 power_generator.type in types])

		def _get_H_(types: [int]) -> int:
			return sum(
				[power_generator.totalNums for power_generator in power_generators if power_generator.type in types])

		power_generators = self.solver.powerGenerators
		huodian_touzi = huodian_zhuangji * _get_HxL_([101]) / _get_H_([101])
		fengdian_touzi = huodian_zhuangji * _get_HxL_([200]) / _get_H_([200])
		guangfu_touzi = huodian_zhuangji * _get_HxL_([300]) / _get_H_([300])

		self.data['投资'] = [huodian_touzi,
						   fengdian_touzi + guangfu_touzi,
						   fengdian_touzi,
						   guangfu_touzi,
						   huodian_touzi + fengdian_touzi + guangfu_touzi,
						   ]

	def get_ranhaochengben(self):
		huodianfadianliang = self.sum_electricity_rst('火力发电')

		def _get_FCxFPxH_(types: [int]) -> int:
			return sum(
				[power_generator.fuelConsumption * power_generator.fuelPrice * power_generator.totalNums for
				 power_generator in power_generators if power_generator.type in types])

		def _get_H_(types: [int]) -> int:
			return sum(
				[power_generator.totalNums for power_generator in power_generators if power_generator.type in types])

		power_generators = self.solver.powerGenerators
		huodian_touzi = huodianfadianliang * _get_FCxFPxH_([101]) / _get_H_([101]) / 10 ** 3
		self.data['煤耗成本'] = [huodian_touzi]
		self.data['煤耗成本'].append(0)
		self.data['煤耗成本'].append(0)
		self.data['煤耗成本'].append(0)
		self.data['煤耗成本'].append(huodian_touzi)

	def sum_electricity_rst(self, item):
		return sum(
			[self.electricity_result_element._data[item][node_id].data[0] for node_id in self.solver.TOTAL_NODE_ID])

	def get_nianyunxingchengben(self, ):
		huodianfadianliang = self.sum_electricity_rst('火力发电')
		xinnengyuanfadianliang = self.sum_electricity_rst('新能源发电')

		def _get_yunxingfeiPlusYunwei(types: [int]) -> int:
			return sum(
				[power_generator.yunweifeilv + power_generator.yunxingfei for
				 power_generator in power_generators if power_generator.type in types])

		power_generators = self.solver.powerGenerators

		huodianyunxingchengben = huodianfadianliang * _get_yunxingfeiPlusYunwei([101])
		xinnengyuanchengben = xinnengyuanfadianliang * _get_yunxingfeiPlusYunwei([200, 300])

		self.data['年运行成本'] = [huodianyunxingchengben, xinnengyuanchengben, 0, 0,
							  huodianyunxingchengben + xinnengyuanchengben]

	def get_zongfadianliang(self):
		huodianfadianliang = self.sum_electricity_rst('火力发电')
		xinnengyuanfadianliang = self.sum_electricity_rst('新能源发电')

		self.data['总发电量'] = [huodianfadianliang, xinnengyuanfadianliang, 0, 0,
							 huodianfadianliang + xinnengyuanfadianliang]

	def get_qidianliang(self):
		xinnengyuanfadianliang = self.sum_electricity_rst('新能源弃电')
		self.data['弃电量'] = [0, xinnengyuanfadianliang, 0, 0, xinnengyuanfadianliang]

	def to_excel(self, filepath):
		self.data = pd.DataFrame(self.data)
		save_result(self.data, filepath)

	def __repr__(self):
		return f'{self.data}'


class TechnicalStaticsResult(BaseResult):

	def __init__(self, solver, power_result, electricity_result, debug=False):
		BaseResult.__init__(self, solver, debug=debug)
		self._data = {}
		self.power_result = power_result
		self.electricity_result = electricity_result

		self.build_data(False)
		logging.debug(f'TechnicalStaticsResult Construced')

	def __call__(self, *args, **kwargs):
		assert len(args) == 1, f'Expected len(args) == 1, but got {args}'
		date_type = args[0]
		if date_type not in self._data:
			self._data[date_type] = TechnicalStaticsResultElement(self.solver, self.power_result,
																  self.electricity_result,
																  date_type)

		return self._data[date_type]

	def build_data(self, optimized=False):
		self(optimized)
		logging.debug(f'TechnicalStaticsResult Builded Data')

	def __repr__(self):
		return f'{self._data}'

	def to_excel(self, date_type=False, file_path="result0125.xlsx"):
		data = self(date_type)
		data.to_excel(file_path)
