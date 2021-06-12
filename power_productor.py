# encoding=utf-8
from collections import OrderedDict
from typing import List

from power_generator import PowerGenerator
from status import Status
from util import defaultDictList, checkIsInteger


class PowerProductor(object):
    '''
    第三部的计算结果，代表每一个电站在每一个场景下的新能源/火电出力情况
    分为两种productor，分别是新能源和火电出力
    如果是新能源productor，那么thermalProduct为0
    如果是火电productor,那么cleanProduct为0

    status:Status
    powerGenerator:PowerGenerator
    cleanProduct:float，新能源出力
    thermalProduct:float，火电出力,只读属性，由开机数open_schema * 单位技术出力techPower求得
    open_schema:int，火电开机数
    '''

    def __init__(self,
                 status: Status,
                 powerGenerator: PowerGenerator,
                 cleanProduct: List[int] = None
                 ):
        if cleanProduct is None:
            cleanProduct = [0 for _ in range(24)]
        self.status = status
        self.powerGenerator = powerGenerator
        self.cleanProduct = cleanProduct

        self._open_schema = 0  # 火电开机数量
        self.distributeSchema = [0] * 24  # 表示该电站除强迫出力外的分配出力
        self.abandonDistribute = [0] * 24  # 正值，表示该电站的弃电

    @property
    def techPowerMulCapacity(self):
        return self.powerGenerator.techPower * self.powerGenerator.singleCapacity

    def __repr__(self):
        return f"出力情况={self.status}," \
               f"powerGenerator={self.powerGenerator}" \
               f"火电强迫出力={self.forceThermalPowerUnderSchema}" \
               f"火电开机台数={self._open_schema}" \
               f"新能源出力={self.cleanProduct}" \
               f"\n"

    def isCleanProductor(self):
        return self.powerGenerator.isCleanPowerStation()

    def isThermalProductor(self):
        return not self.isCleanProductor()

    def adjustOneOpenSchema(self, plus=True):
        '''
        在当前的电站中新开一个机器或者关闭一个机器
        :param plus:
        :return: 如果开机或者关机成功，则返回True，否则返回False
        '''
        if plus:
            if not self.isFull():
                self._open_schema += 1
            else:
                return False
        else:
            if not self.isEmpty():
                self._open_schema -= 1
            else:
                return False
        return True

    def isEmpty(self):
        return self.isThermalProductor() and self._open_schema == 0

    def isFull(self):
        return self.isThermalProductor() and self._open_schema == self.powerGenerator.totalNums

    def thermalPowerCapacity(self):
        """
        当前电站的极限火电出力情况，开完该电站所有的电机的总火电出力
        """
        return self.powerGenerator.totalNums * self.powerGenerator.singleCapacity

    def thermalPowerCapacityUnderSchema(self):
        """
        当前开机方案下的最大火电出力
        :return:
        """
        return self._open_schema * self.powerGenerator.singleCapacity

    def capacityUnderSchemaAtHour(self, hour):
        """
        返回给定小时的最大出力, 风光电和火电通用，对于火电，是当前开机方案下的容量，对于风光电是当前小时的被动出力
        :param hour:
        :return:
        """
        return self.thermalPowerCapacityUnderSchema() + self.cleanProduct[hour]

    def currentOutputAtHour(self, hour, highLevelOnly=False):
        """
        返回给定小时该类电站的实际出力，对于火电，包括强制出力和一分配的出力，对于风光电，是当前小时的被动出力
        :param hour:
        :param highLevelOnly: 是否只计算上层网络的输出，在计算电力提升空间时，不考虑弃电的减少，计算对网络的输出时，需要扣除弃电
        :return:
        """
        output = self.forceThermalPowerUnderSchema + self.cleanProduct[hour] + self.distributeSchema[hour]
        return output - self.abandonDistribute[hour] if highLevelOnly else output

    def currentThermalAtHour(self, hour):
        """
        返回给定小时火电的实际出力，（弃电之前的单独火电出力，不包括风光联合新能源出力）

        """
        return self.forceThermalPowerUnderSchema + self.distributeSchema[hour]

    def NewenergythisNode(self, hour):
        """
        返回该节点的风光联合出力
        :param hour:
        :return:
        """
        return self.cleanProduct[hour]

    @property
    def open_schema(self):
        return self._open_schema

    @open_schema.setter
    def open_schema(self, value):
        if value < 0 or value > self.powerGenerator.totalNums:
            raise ValueError(f'Expected 0<=open_schema<={self.powerGenerator.totalNums}, but got {value}')
        self._open_schema = value

    @property
    def maxthermalPower(self):
        """
        当前开机方案下的最大火电出力
        :return:
        """
        return self._open_schema * self.powerGenerator.singleCapacity

    @property
    def forceThermalPowerUnderSchema(self):
        """
        该类型电站的火电强迫出力
        :return:
        """
        return self._open_schema * self.powerGenerator.techPower * self.powerGenerator.singleCapacity

    def cost(self):
        return self.powerGenerator.fuelPrice * self.powerGenerator.fuelConsumption / 1000. * self.forceThermalPowerUnderSchema

    def debugResult(self):
        return f"stationID: {self.powerGenerator.stationID}, open_count: {self._open_schema},\n \t\t\t\t{self.distributeSchema}"

    def __lt__(self, other):
        """
        PowerGenerator的默认排序按照经济性从小到大排列
        :param other:
        :return:
        """
        return self.cost() < other.cost()


class PowerProductDict(object):
    """
    同一个场景下，一个PowerProductor汇总的map
    {(sceneNum,month):{powerGenerator:powerProductDict,...},...}
    使用方法：
    powerProductDict[sceneNum,month] -> {powerGenerator:PowerProductor,...}
    powerProductDict[sceneNum,month,stationID] -> PowerProductor
    powerProductDict[sceneNum,month,powerGenerator] -> PowerProductor
    """

    def __init__(self, powerProductors: [PowerProductor]):
        self._powerProductors = {}
        for powerProductor in powerProductors:
            here = self._powerProductors
            elem = (powerProductor.status.sceneNum, powerProductor.status.month)
            if elem not in here:
                here[elem] = {}
            here = here[elem]
            here[powerProductor.powerGenerator] = powerProductor

    def __getitem__(self, item):
        __ERROR_STRING__ = f'item must be (sceneNum,month) or (sceneNum,month,stationID) or (sceneNum,month,powerGenerator), but got {item}'
        assert 2 <= len(item) <= 3 and checkIsInteger(item[0]) and checkIsInteger(item[1]), __ERROR_STRING__
        if len(item) == 2:
            return self._powerProductors[item]
        elif len(item) == 3:
            if checkIsInteger(item[2]):
                return self.__getItemByStationID(*item)
            elif isinstance(item[2], PowerGenerator):
                return self.__getItemByPowerGenerator(*item)
            else:
                raise ValueError(__ERROR_STRING__)
        else:
            raise ValueError(__ERROR_STRING__)

    def __getItemByStationID(self, sceneNum, month, stationID):
        for powerGenerator in self._powerProductors[sceneNum, month]:
            if powerGenerator.stationID == stationID:
                return self._powerProductors[sceneNum, month][powerGenerator]

    def __getItemByPowerGenerator(self, sceneNum, month, powerGenerator):
        return self._powerProductors[sceneNum, month][powerGenerator]


class PowerProductList(object):
    """
   同一个Status和nodeID下，多个PowerProductor汇总的map
   {(sceneNum,month):{nodeID:[PowerProduct]},...}

   使用方法：
    PowerProductList[sceneNum,month] -> {nodeID:[PowerProduct],...}
    PowerProductList[sceneNum,month,nodeID] -> [PowerProductor]

    """

    def __init__(self, powerProductors: [PowerProductor], _sortedNodeIDs=None):
        self._powerProductorList = {}
        self._sortedNodeIDs = _sortedNodeIDs

        for powerProductor in powerProductors:
            here = self._powerProductorList
            elem = (powerProductor.status.sceneNum, powerProductor.status.month)
            if elem not in here:
                here[elem] = {}
            here = here[elem]
            defaultDictList(here,
                            powerProductor.powerGenerator.systemID,
                            powerProductor)

        self.sortNodesByValidNodeIDs(self._sortedNodeIDs)

    def __getitem__(self, item):
        __ERROR_STRING__ = f'item must be (sceneNum,month) or (sceneNum,month,nodeID), but got {item}'
        assert checkIsInteger(item[0]) and checkIsInteger(item[1]), __ERROR_STRING__
        if len(item) == 2:
            return self._powerProductorList[item]
        elif len(item) == 3:
            if checkIsInteger(item[2]):
                return self.__getItemsByNodeID(*item)
            else:
                raise ValueError(__ERROR_STRING__)
        else:
            raise ValueError(__ERROR_STRING__)

    def __getItemsByNodeID(self, sceneNum, month, nodeID):
        if nodeID not in self._powerProductorList[sceneNum, month]:
            raise ValueError(
                f'Expected nodeID in _powerProductorList={self._powerProductorList[sceneNum, month].keys()}, but got {nodeID},')
        return self._powerProductorList[sceneNum, month][nodeID]

    def getThermalCapacityUnderSchema(self, sceneNum, month, nodeID):
        """
        返回sceneNum, month情况下的nodeID里全部已经开机的火电容量总和
        :param sceneNum:
        :param month:
        :param nodeID:
        :return:
        """
        powerProducers = self[sceneNum, month, nodeID]
        capacity = 0
        for powerProducer in powerProducers:
            if powerProducer.isThermalProductor():
                capacity += powerProducer.thermalPowerCapacityUnderSchema()
        return capacity

    def sortNodesByValidNodeIDs(self, VALID_NODE_ID):
        """
        根据`VALID_NODE_ID`中新的nodeID的顺序，对该结构中的nodeID重新排序
        """
        if VALID_NODE_ID is None:
            return
        self._sortedNodeIDs = VALID_NODE_ID
        for (sceneNum, month), nodes in self._powerProductorList.items():
            ordered_nodes = OrderedDict(
                (nodeID, self[sceneNum, month, nodeID]) for nodeID in VALID_NODE_ID)

            self._powerProductorList[sceneNum, month] = ordered_nodes

    def sortStationsByCost(self):
        """
        根据电站在所有场景下的火电出力经济性，对所有PowerProductors进行排序
        """
        for (sceneNum, month), nodes in self._powerProductorList.items():
            for nodeID in nodes:
                nodes[nodeID] = sorted(nodes[nodeID])

    def getCleanProductors(self, sceneNum, month, nodeID=None):
        """
        返回所有的新能源电站PowerProductors
        """
        if nodeID is not None:
            return [powerProductor for powerProductor in self[sceneNum, month, nodeID] if
                    powerProductor.isCleanProductor()]
        else:
            res = []
            for nodeID in self[sceneNum, month]:
                res.extend(self.getCleanProductors(sceneNum, month, nodeID))
            return res

    def getThermalProductors(self, sceneNum, month, nodeID=None):
        """
        返回所有的火电电站PowerProductors
        """
        if nodeID is not None:
            return [powerProductor for powerProductor in self[sceneNum, month, nodeID] if
                    powerProductor.isThermalProductor()]
        else:
            res = []
            for nodeID in self[sceneNum, month]:
                res.extend(self.getThermalProductors(sceneNum, month, nodeID))
            return res

    def isFull(self, sceneNum, month, nodeID=None):
        """
        判断当前状态下的节点nodeID中的所有火电电站是否都开满
        """
        if nodeID is not None:

            for productor in self[sceneNum, month, nodeID]:
                if productor.isThermalProductor() and not productor.isFull():
                    return False
            return True
        else:
            return all([self.isFull(sceneNum, month, nodeID) for nodeID in self[sceneNum, month]])

    def isEmpty(self, sceneNum, month, nodeID=None):
        """
        判断当前状态下的节点nodeID中的所有火电电站是否都是关机状态
        """
        if nodeID is not None:
            for productor in self[sceneNum, month, nodeID]:
                if productor.isThermalProductor() and not productor.isEmpty():
                    return False
            return True
        else:
            return all([self.isEmpty(sceneNum, month, nodeID) for nodeID in self[sceneNum, month]])

    def maxThermalPower(self, sceneNum, month, nodeID=None):
        """
        返回当前状态下的节点nodeID中的所有极限火电出力之和
        如果nodeID不指定，那么就是取得所有node的所有极限火电出力之和
        """
        if nodeID is not None:
            return sum(
                productor.thermalPowerCapacity() for productor in self.getThermalProductors(sceneNum, month, nodeID))
        else:
            return sum(self.maxThermalPower(sceneNum, month, nodeID) for nodeID in self[sceneNum, month])

    def maxCleanPower(self, sceneNum, month, nodeID=None):
        """
        返回当前状态下的节点nodeID中的所有新能源出力之和
        """
        if nodeID is not None:
            return sum(productor.cleanProduct for productor in self.getCleanProductors(sceneNum, month, nodeID))
        return sum([self.maxCleanPower(sceneNum, month, nodeID) for nodeID in self[sceneNum, month]])

    def maxPower(self, sceneNum, month, nodeID=None):
        """
        返回当前状态下的节点nodeID中的所有能源出力之和
        """
        if nodeID is None:
            return self.maxThermalPower(sceneNum, month, nodeID) + self.maxCleanPower(sceneNum, month, nodeID)
        return sum([self.maxPower(sceneNum, month, nodeID) for nodeID in self[sceneNum, month]])

    def getCurrentPowerByNode(self, nodeId, sceneNum, month, hour):
        powerProducers = self[sceneNum, month, nodeId]
        return sum(p.currentOutputAtHour(hour, highLevelOnly=True) for p in powerProducers)

    def printOpenSchema(self, sceneNum, month, nodeID):
        s = f'节点{nodeID} 火电开机情况\n'
        for powerProduct in self.getThermalProductors(sceneNum, month, nodeID):
            s += f"{powerProduct.powerGenerator.stationID}: {powerProduct.open_schema}/{powerProduct.powerGenerator.totalNums}, Full={powerProduct.isFull()}\n"
        return s
