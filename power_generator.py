from util import checkIsInteger


class PowerGenerator:  # 一个节点所包含的所有种类电站的列表
	def __init__(self, stationID: int, systemID, fuelConsumption: int, fuelPrice: int, characterID: int, type: int,
				 totalNums: int, techPower: float, dongtaitouzi: float, singleCapacity: int, zhuangji: int,
				 yunweifeilv: int, yunxingfei: int):  # 初始化数据格式
		# 实例化
		self.fuelConsumption = fuelConsumption
		self.fuelPrice = fuelPrice
		self.stationID = stationID
		self.systemID = systemID
		self.characterID = characterID
		self.type = type
		self.totalNums = totalNums
		self.techPower = techPower
		self.dongtaitouzi = dongtaitouzi
		self.singleCapacity = singleCapacity
		self.zhuangji = zhuangji
		self.yunweifeilv = yunweifeilv
		self.yunxingfei = yunxingfei

		self.forcePower = techPower * singleCapacity

	# __repr__()方法是类的实例化对象用来做“自我介绍”的方法
	def __str__(self):  # 前后都有双下划线的为系统内置方法或属性
		return (
			f" 电站ID={self.stationID},"
			f" 系统ID={self.systemID},"
			f" 燃料单耗(g/kWh)={self.fuelConsumption},"
			f" 燃料单价(元/t)={self.fuelPrice},"
			f" 火电特性ID={self.characterID},"
			f" 电站类型={self.type},"
			f" 台数={self.totalNums},"
			f" 技术出力={self.techPower},"
			f" 动态投资={self.dongtaitouzi},"
			f" 单机容量={self.singleCapacity}"
			f" 装机={self.zhuangji}"
			f"\n"
		)

	def getTotalCapacity(self):  # 各电站总容量=单机*台数

		return self.singleCapacity * self.totalNums

	def isCleanPowerStation(self):  # 判断是否为新能源，返回True或者False
		return self.type == 200 or self.type == 300

	# return self.type >= 200

	def cost(self):  # 不同型号机组的单位发电成本，用于比较机组的经济性
		return self.fuelPrice * self.fuelConsumption / 1000

	def __lt__(self, other):  # 对实例的排序，按照单位发电成本排序
		return self.cost() < other.cost()


class PowerGeneratorDict(object):  # 把电站信息存储在字典中
	def __init__(self, powerGenerators: [PowerGenerator]):
		self._powerGenerators = {}  # 前边有一个下划线的为内变量
		self.thermalGeneratorIds = []  # 存火电ID并排序
		self.cleanGeneratorIds = []  # 存新能源ID并排序
		for powerGenerator in powerGenerators:
			# 这样做是为了将电站的唯一标识stationID作为字典的键值，方便取
			self._powerGenerators[powerGenerator.stationID] = powerGenerator  # 值是电站ID对应的电站对象
			if powerGenerator.isCleanPowerStation():
				self.cleanGeneratorIds.append(powerGenerator.stationID)
			else:
				self.thermalGeneratorIds.append(powerGenerator.stationID)
		# 对电站ID默认从小到大进行排序
		self.thermalGeneratorIds.sort()
		self.cleanGeneratorIds.sort()

	def __getitem__(self, item):
		# 如果在类中定义了__getitem__()方法，那么他的实例对象（假设为P）就可以通过P[key]取值。
		# 这个用在65行：self._powerGenerators[powerGenerator.stationID] = powerGenerator，即根据键值ID取值
		assert len(item) == 1 and checkIsInteger(item), f'关键字必须为电站ID, 错误内容：{item}'
		# assert用来判断键值的格式是否有错
		return self._powerGenerators[item]
