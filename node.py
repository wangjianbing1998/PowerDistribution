from collections import OrderedDict
from typing import List

import power_productor
from power_generator import PowerGenerator
from power_productor import PowerProductor
from util import checkIsInteger


class Node:
    def __init__(
            self,
            nodeID: int,  # nodeID为节点/系统的ID
            powerGenerators: List[PowerGenerator],  # powerGenerators为该节点的发电机所组成的列表，已经按照单位成本排好序了
            commutationMax: int,  # commutationMax,换流Max，来自算例-系统表
            commutationMin: int,  # commutationMin,换流Min，来自算例-系统表
    ):

        self.nodeID = nodeID
        self.commutationMin = commutationMin
        self.commutationMax = commutationMax
        self.powerGenerators = powerGenerators
        # 初始化了三个用于快速查询node中含有的stationId的电站的结构
        self.thermalPowerStationIds = []  # 该节点的火电站Id列表
        self.cleanPowerStationIds = []  # 该节点的新能源电站Id列表

        for p in self.powerGenerators:
            if p.isCleanPowerStation():
                self.cleanPowerStationIds.append(p.stationID)
            else:
                self.thermalPowerStationIds.append(p.stationID)

        self.__reInitStationIdIterator()

        self.stationIds = self.thermalPowerStationIds.copy()  # 该节点的所有电站StationId列表
        self.stationIds.extend(self.cleanPowerStationIds)

        # 记录stationID到powerGenerator的映射，方便查找
        self.id2powerGenerator = {}
        for powerGenerator in self.powerGenerators:
            self.id2powerGenerator[powerGenerator.stationID] = powerGenerator

    def __reInitStationIdIterator(self):
        self.thermalPowerStationIndex = 0

    def getNextThermalPowerStationID(self):
        """
        如果该节点一个电站都没有，则返回None
        如果该节点的所有电站都遍历完了，则重新遍历，返回None
        """
        if len(self.thermalPowerStationIds) == 0:
            return None  # 代表一个火电都没有
        if self.thermalPowerStationIndex < len(self.thermalPowerStationIds):
            index_ = self.thermalPowerStationIds[self.thermalPowerStationIndex]
            self.thermalPowerStationIndex += 1
            return index_
        else:
            self.thermalPowerStationIndex = 0
            return None

    def getCurrentThermalPowerStationID(self):
        if 0 <= self.thermalPowerStationIndex < len(self.thermalPowerStationIds):
            return self.thermalPowerStationIds[self.thermalPowerStationIndex]
        return None

    def __repr__(self):
        return (
            f"节点ID={self.nodeID}, "
            f"powergenerator={self.powerGenerators}, "
            f"换流器最大容量={self.commutationMax}, "
            f"换流器最小容量={self.commutationMin}\n"
        )

    def getStationIDs(self) -> list:
        """
        获得一个节点的所有电站StationId列表
        :return:
        """
        return self.stationIds

    def __getitem__(self, stationID):
        """
        获得该节点拥有的与stationID的对应的类型的电站的容量
        :param stationID: 电站ID
        :return:
        """
        return self.id2powerGenerator[stationID]

    def getPowerGenerators(self, stationIDs=None):
        """
        :param nodeID:
        :return:
        """
        if stationIDs is not None:
            return [self.id2powerGenerator[stationID] for stationID in stationIDs]
        else:
            return self.powerGenerators

    def canIncrease(self, productsThisNode: List[PowerProductor], hour):
        """
        判断该节点是否可以继续分配待分配电力，返回能够增加的数额

        :param hour:
        :param productsThisNode:
        :param powerDistributed:
        :param openSchema: 该节点的openSchema，{PowerGenerator:openNum}
        :return: 
        """
        openCapacity = sum(p.capacityUnderSchemaAtHour(hour) for p in productsThisNode)
        currentPower = sum(p.currentOutputAtHour(hour, highLevelOnly=False) for p in productsThisNode)
        # 该节点的总输出不能超过火电开机容量+风电量 或 换电站容量
        return min(openCapacity, self.commutationMax) - currentPower

    @staticmethod
    def canDrop(productsThisNode: List[PowerProductor], hour):
        """
        当前该节点还可以弃多少电
        :param productsThisNode:
        :param hour:
        :return:
        """
        # 该节点的电站的上层网络出力之和
        return sum(p.currentOutputAtHour(hour, highLevelOnly=True) for p in productsThisNode)

    def dropBy(self, productDict: power_productor.PowerProductDict, sceneNum: int, month: int, hour: int,
               powerToDrop: int):
        """
        将该节点的当前电力丢弃给定值
        :param productDict:
        :param sceneNum:
        :param month:
        :param hour:
        :param powerToDrop:
        :return:
        """
        for stationId in self.stationIds:
            if powerToDrop == 0:
                break
            product = productDict[sceneNum, month, stationId]
            powerCanDrop = min(product.currentOutputAtHour(hour, highLevelOnly=True), powerToDrop)
            product.abandonDistribute[hour] += powerCanDrop
            powerToDrop -= powerCanDrop

        return powerToDrop == 0

    @staticmethod
    def canDecrease(productsThisNode: List[PowerProductor], hour):
        """
        判断该节点是否可以减少已分配的电力，返回能够减少的数额
        :param productsThisNode:
        :param hour:
        :return:
        """
        return sum(p.distributeSchema[hour] for p in productsThisNode)

    def increaseBy(self, productDict: power_productor.PowerProductDict, sceneNum: int, month: int, hour: int,
                   powerToIncrease: int):
        """
        增加给定节点的给定小时的电力输出，返回提高是否成功

        :param month:
        :param sceneNum:
        :param productDict:
        :param hour:
        :param powerToIncrease:
        :return:
        """
        # thermalPowerStationIds的顺序在初始化时，已经按照经济性排序，因此这里从最经济的电站开始增加
        for stationId in self.thermalPowerStationIds:
            if powerToIncrease == 0:
                break
            product = productDict[sceneNum, month, stationId]
            powerCanIncrease = product.capacityUnderSchemaAtHour(hour) - product.currentOutputAtHour(hour)
            powerCanIncrease = min(powerCanIncrease, powerToIncrease)
            product.distributeSchema[hour] += powerCanIncrease
            powerToIncrease -= powerCanIncrease

        return powerToIncrease == 0

    def decreaseBy(self, productDict: power_productor.PowerProductDict, sceneNum: int, month: int, hour: int,
                   powerToDecrease: int):
        """
        降低给定节点的给定小时电力输出，返回是否减少成功, 从经济性最差的机组开始减

        :param month:
        :param sceneNum:
        :param productDict:
        :param hour:
        :param powerToDecrease:
        :return:
        """
        for stationId in reversed(self.thermalPowerStationIds):
            if powerToDecrease == 0:
                break
            product = productDict[sceneNum, month, stationId]
            powerCanDecrease = product.distributeSchema[hour]
            powerCanDecrease = min(powerCanDecrease, powerToDecrease)
            product.distributeSchema[hour] -= powerCanDecrease
            powerToDecrease -= powerCanDecrease

        return powerToDecrease == 0


class NodeDict(object):
    def __init__(self, nodes: [Node]):
        self._nodes = {}
        self.node2stationIDs = {}

        for node in nodes:
            self._nodes[node.nodeID] = node
            self.node2stationIDs[node.nodeID] = [powerGenerator.stationID for powerGenerator in node.powerGenerators]

    def __getitem__(self, item):
        assert checkIsInteger(item), f'标识符必须为[节点ID], 错误内容：{item}'
        return self._nodes[item]

    @property
    def totalNodeIDs(self):
        return self._nodes.keys()

    def sortedNodesByValidNodeID(self, VALID_NODE_ID):
        self._nodes = OrderedDict((nodeID, self._nodes[nodeID]) for nodeID in VALID_NODE_ID)
