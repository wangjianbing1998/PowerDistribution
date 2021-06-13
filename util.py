# encoding=utf-8
"""
@Time: 2020/12/1 10:31
@File: util.py
@Description:
"""
import logging
import pickle

import pandas as pd

pd.set_option('expand_frame_repr', False)
pd.set_option('display.max_rows', 20)
pd.set_option('precision', 2)
import numpy as np


def defaultDictList(d: dict, key, value):
	if key not in d:
		d[key] = [value]
	else:
		d[key].append(value)


def checkListSame(A, B):
	return set(A) == set(B)


def checkIsInteger(A):
	return type(A) in [np.int, np.int64, np.int32]


def mergeresult(solution):
	A = solution.abandonOnNode
	B = solution.powerOnNode
	C = solution.thermalPowerOnNode
	D = solution.cleanPowerOnNode
	X1 = A + B + C + D
	return X1


def saveToExcel(solution):
	def preprocess_(df, columns):
		df.rename(columns=columns, inplace=True)
		return df

	def save_nodes(dfs, sheet_names, file_path):
		write = pd.ExcelWriter(file_path)
		for df, sheet_name in zip(dfs, sheet_names):
			df.to_excel(write, sheet_name=sheet_name, index=False)
		write.save()

	save_nodes(dfs=[preprocess_(pd.DataFrame(df), columns={24: "节点ID"}) for df in
					[solution.powerOnNode, solution.cleanPowerOnNode, solution.thermalPowerOnNode,
					 solution.abandonOnNode]]
				   + [preprocess_(pd.DataFrame(df), columns={24: "电站ID"}) for df in
					  [solution.cleanPowerOnStation, solution.thermalPowerOnStation]]
				   + [preprocess_(pd.DataFrame(df), columns={1: "电站ID"}) for df in
					  [solution.openSchemaOnStation, solution.forcePowerOnStation]],
			   sheet_names=['上层出力-节点', "风光联合出力-节点", "火电实际出力-节点", "弃电"]
						   + ["风光联合出力-电站", "火电实际出力-电站"]
						   + ["开机方案", "强迫出力"],
			   file_path="result0.xlsx")
	save_nodes(
		dfs=[preprocess_(pd.DataFrame(df), columns={24: "线路ID", 25: "首端ID", 26: "末端ID"}) for df in
			 [solution.powerOnArcs]],
		sheet_names=['边上功率'], file_path="result_arc.xlsx")


def save_object(obj, filename):
	with open(filename, 'wb') as f:
		s = pickle.dumps(obj)
		f.write(s)

	logging.info(f'Obejct {filename} Saved')


def load_object(filename):
	with open(filename, 'rb') as f:
		obj = pickle.loads(f.read())

	logging.info(f'Obejct {filename} Loaded')
	return obj


def save_result(df, file_path, sort_bys=None):
	if sort_bys is not None:
		df.sort_values(by=sort_bys, ascending=True, inplace=True)
	if '项目编号' in df:
		df.drop('项目编号', axis=1, inplace=True)
	df.to_csv(file_path, index=False, encoding='utf_8_sig')
	df.to_excel(file_path.replace('.csv', '.xlsx'), index=False)


if __name__ == "__main__":
	class A():
		AA = 1

		def __init__(self, b):
			self.b = b

		def __repr__(self):
			return f'{self.AA},{self.b}'


	save_object(A(123), '123.pkl')
	print(load_object('123.pkl'))
