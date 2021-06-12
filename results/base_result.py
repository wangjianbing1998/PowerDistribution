# encoding=utf-8
'''
@File    :   base_result.py    
@Contact :   jianbingxiaman@gmail.com
@License :   (C)Copyright 2020-2021, John Hopcraft Lab-CV
@Desciption : 
@Modify Time      @Author    @Version
------------      -------    --------
2021/5/9 19:36   jianbingxia     1.0    
'''
from abc import ABC, abstractmethod


def ave(A):
	return sum(A) / len(A) if len(A) else 0


class BaseData(ABC):
	def __init__(self, size=28):
		self.size = size
		self.data = [None for _ in range(size)]

	def __repr__(self):
		return f'{self.data}'

	@abstractmethod
	def __add__(self, other):
		"""add other BaseData """
		pass

	@abstractmethod
	def __len__(self):
		"""the len of data"""
		pass

	def __iadd__(self, other):
		data = self + other
		self.data = data.data
		return self

	def __isub__(self, other):
		data = self - other
		self.data = data.data
		return self

	def __idiv__(self, other_data):
		data = self / other_data
		self.data = data.data
		return self

	def __imul__(self, other_data):
		data = self * other_data
		self.data = data.data
		return self


class HourData(BaseData):
	"""data is a list of 4 hour"""

	def __init__(self, hour_data=[0 for _ in range(24)]):

		BaseData.__init__(self, 28)

		if hasattr(hour_data, 'tolist'):
			hour_data = hour_data.tolist()
		elif hasattr(hour_data, 'to_list'):
			hour_data = hour_data.to_list()

		self.data = [max(hour_data), min(hour_data), ave(hour_data), sum(hour_data)] + hour_data

	def __add__(self, other):
		if len(self.data) != len(other.data):
			raise ValueError(
				f'Expected len(self.data) == len(other.data), but got {len(self.data)} and {len(other.data)}')
		data = [a + b for a, b in zip(self.data[4:], other.data[4:])]
		return HourData(data)

	def __radd__(self, other):
		return self + other

	def __len__(self):
		return len(self.data) - 4

	def __truediv__(self, other_data):
		data = [a / other_data for a in self.data[4:]]
		return HourData(data)

	def __mul__(self, other_data):
		data = [a * other_data for a in self.data[4:]]
		return HourData(data)

	def get_sum(self):
		return self.data[3]


class MonthData(BaseData):
	"""data is a list of 4 hour"""

	def __init__(self, month_data=[0 for _ in range(12)]):

		BaseData.__init__(self, 13)

		if hasattr(month_data, 'tolist'):
			month_data = month_data.tolist()
		elif hasattr(month_data, 'to_list'):
			month_data = month_data.to_list()

		self.data = [sum(month_data)] + month_data

	def __add__(self, other):
		if len(self.data) != len(other.data):
			raise ValueError(
				f'Expected len(self.data) == len(other.data), but got {len(self.data)} and {len(other.data)}')
		data = [a + b for a, b in zip(self.data[1:], other.data[1:])]
		return MonthData(data)

	def __radd__(self, other):
		return self + other

	def __len__(self):
		return len(self.data) - 1

	def __truediv__(self, other_data):
		data = [a / other_data for a in self.data[1:]]
		return MonthData(data)

	def __mul__(self, other_data):
		data = [a * other_data for a in self.data[1:]]
		return MonthData(data)


class OneData(BaseData):

	def __init__(self, data=0):
		BaseData.__init__(self, 28)
		self.data[0] = data

	def __add__(self, other):
		if isinstance(other, OneData):
			other = other.data[0]

		data = self.data[0] + other
		return OneData(data)

	def __len__(self):
		return 1

	def __truediv__(self, other_data):
		if isinstance(other_data, OneData):
			other_data = other_data.data[0]

		data = self.data[0] / other_data
		return OneData(data)

	def __mul__(self, other_data):
		if isinstance(other_data, OneData):
			other_data = other_data.data[0]

		data = self.data[0] * other_data
		return OneData(data)


class BaseResult(ABC):
	sceneNum, month = 0, 1

	def __init__(self, solver=None, debug=False):
		self._data = {}
		self.solver = solver
		self.debug = debug

	@abstractmethod
	def build_data(self):
		"""build all data"""
		pass

	def __getitem__(self, item):
		assert 2 <= len(item) <= 3 and isinstance(item[0], int) and isinstance(item[1], int) \
			   and 1 <= item[1] <= 12, f'Expected item is (sceneNum,month,dateType), but got {item}'

		return self._data[item]
