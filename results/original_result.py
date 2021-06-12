# encoding=utf-8
'''
@File    :   original_result.py    
@Contact :   jianbingxiaman@gmail.com
@License :   (C)Copyright 2020-2021, John Hopcraft Lab-CV
@Desciption : 
@Modify Time      @Author    @Version
------------      -------    --------
2021/5/15 20:30   jianbingxia     1.0    
'''
import pandas as pd

from results.base_result import BaseResult
from util import mergeresult


class OriginalResult(BaseResult):
	def __init__(self, solver, debug):
		BaseResult.__init__(self, solver, debug)

	def build_data(self):
		self._data = self.solver.matrixchange()

	def __call__(self, *args, **kwargs):
		assert 2 <= len(args) <= 3 and isinstance(args[0], int) and isinstance(args[1], int) \
			   and 1 <= args[1] <= 12, f'Expected item is (sceneNum,month,dateType), but got {args}'
		return self._data[args[0], args[1]]

	def __repr__(self):
		return f'{self._data}'

	def save_to_excel(self, solution, file_path="result0125.xlsx"):
		def preprocess_(df, columns):
			df.rename(columns=columns, inplace=True)
			return df

		def save_nodes(dfs, sheet_names, file_path):
			write = pd.ExcelWriter(file_path)
			for df, sheet_name in zip(dfs, sheet_names):
				df.to_excel(write, sheet_name=sheet_name, index=False)
			write.save()

			save_nodes(dfs=[preprocess_(pd.DataFrame(df), columns={
				0: "节点ID", 1: "项目", 2: "max", 3: "min", 4: "ave", 5: "1h", 6: "2h", 7: "3h", 8: "4h", 9: "5h", 10: "6h",
				11: "7h", 12: "8h",
				13: "9h", 14: "10h", 15: "11h", 16: "12h", 17: "13h", 18: "14h", 19: "15h", 20: "16h", 21: "17h",
				22: "18h", 23: "19h", 24: "20h", 25: "21h", 26: "22h", 27: "23h", 28: "24h"
			}) for df in [mergeresult(solution)]]
						   + [preprocess_(pd.DataFrame(df), columns={
				0: "首节点ID", 1: "线路描述", 2: "max", 3: "min", 4: "线路容量", 5: "1h", 6: "2h", 7: "3h", 8: "4h", 9: "5h",
				10: "6h",
				11: "7h",
				12: "8h",
				13: "9h", 14: "10h", 15: "11h", 16: "12h", 17: "13h", 18: "14h", 19: "15h", 20: "16h", 21: "17h",
				22: "18h", 23: "19h", 24: "20h", 25: "21h", 26: "22h", 27: "23h", 28: "24h"
			}) for df in [solution.powerOnArcs]],
					   sheet_names=['结果表'] + ['线路潮流'],
					   file_path=file_path)

	def to_excel(self, scene_num, month, file_path="result0125.xlsx"):
		data = self(scene_num, month)
		self.save_to_excel(data, file_path)
