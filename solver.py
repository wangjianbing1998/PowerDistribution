import os

from results.electricity_result import ElectricityResult
from results.technical_statics_result import TechnicalStaticsResult

os.environ['NUMEXPR_MAX_THREADS'] = '16'
import logging
import os
import sys
from collections import defaultdict
from statistics import mean
from typing import List, NoReturn, Dict

import numpy as np
import pandas as pd
from sklearn.utils import Bunch

import Path
import graph
import power_generator
import power_productor
from graph import Graph
from node import Node, NodeDict
from single_variables import STEP_IN_POWER_DISTRIBUTION, V0_IN_FLOW_CHECK
from status import Status, StatusDict
from util import checkListSame, defaultDictList, save_object, load_object

pd.set_option("expand_frame_repr", False)
pd.set_option("display.max_rows", 20)
pd.set_option("precision", 2)

logging.basicConfig(
    level=logging.DEBUG,
    # format="%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s",
    filemode="a",
)
Path.filepath = 'data/'
Path.filename = '算例.xlsx'


def readExcelData():
    # 读取表格，后期根据场景曲线是否由程序自动生成来调整读取方式
    sheet_scene = pd.read_excel(
        os.path.join(os.path.dirname(__file__), Path.filepath + '中间数据/' + "场景曲线.xlsx"), sheet_name='场景概率'
    )
    sheet_out = pd.read_excel(
        os.path.join(os.path.dirname(__file__), Path.filepath + '中间数据/' + "场景曲线.xlsx"), sheet_name='外送曲线'
    )
    sheet_new_energy = pd.read_excel(
        os.path.join(os.path.dirname(__file__), Path.filepath + '中间数据/' + "场景曲线.xlsx"), sheet_name='新能源'
    )
    sheet_routine = pd.read_excel(
        os.path.join(os.path.dirname(__file__), Path.filepath + Path.filename), sheet_name='线路表'
    )
    sheet_plant = pd.read_excel(
        os.path.join(os.path.dirname(__file__), Path.filepath + Path.filename), sheet_name='电站表'
    )
    sheet_system = pd.read_excel(
        os.path.join(os.path.dirname(__file__), Path.filepath + Path.filename), sheet_name='系统表'
    )
    sheet_output_attributes = pd.read_excel(
        os.path.join(os.path.dirname(__file__), Path.filepath + Path.filename), sheet_name='外送特性'
    )
    sheet_status_probability = pd.read_excel(
        os.path.join(os.path.dirname(__file__), Path.filepath + '中间数据/' + "场景曲线.xlsx"), sheet_name='场景概率', index_col=[0]
    )

    return (
        sheet_scene,  # 场景曲线-场景概率
        sheet_new_energy,  # 场景曲线-外送曲线
        sheet_out,  # 场景曲线-新能源
        sheet_plant,  # 算例-电站表
        sheet_routine,  # 算例-线路表
        sheet_system,  # 算例-系统表
        sheet_output_attributes,  # 算例-外送特性
        sheet_status_probability,  # 场景曲线-场景概率
    )


class Solver:
    def __init__(self):
        # 包含Dict的就是字典
        self.arcs = []  # 所有点对点线路的列表
        self.statusDict = None  # StatusDict
        self.powerGeneratorDict = None

        self.powerProductors = []
        self.powerProductDict = None
        self.powerProductList = None

        self.thermalPowerStations = []  # 记录所有的火电站
        self.cleanPowerStations = []  # 记录所有的新能源电站

        self.powerGenerators = []

        self.graph = None  # 拓扑图，邻接表形式
        self.nodeDict = None  # 图上的节点{nodeID:Node}
        self.sceneNums = set()  # 记录场景ID，不可重复
        self.VALID_NODE_ID = []  # 记录所有非出节点的nodeID，不可重复，在sortPath中会被排序
        self.OUTPUT_NODE_ID = 0  # 记录出力节点ID
        self.TOTAL_NODE_ID = []  # [self.OUTPUT_NODE_ID]+self.VALID_NODE_ID
        # 下面的变量都已经被初始化了，可以直接用
        self.arcCost = defaultdict(
            float
        )  # 记录每两个节点之间的线路电阻Cost，比如self.arcCost[x,y]代表节点x和节点y之间的路径最小cost

        self.month_status = defaultdict(list)
        self.kr = float  # 备用率，来自算例.xlsx-外送特性-O2
        # 读取所有数据
        (
            sheet_scene,
            sheet_new_energy,
            sheet_out,
            sheet_plant,
            sheet_routine,
            sheet_system,
            sheet_output_attributes,
            sheet_status_probability,
        ) = readExcelData()

        # 初始化一些数据表格中的变量
        self._buildSheetVariables(sheet_output_attributes, sheet_system)

        # 初始化Arcs数据
        self._buildArcs(sheet_routine)

        # 初始化SceneNums
        self._buildScene(sheet_scene)
        # 初始化powerGenerators
        self.powerGenerators = self._buildPowerGenerator(sheet_plant)

        # 初始化场景数据
        statusList = self._buildStatus(sheet_out, sheet_new_energy)

        self._buildNodes(sheet_system, self.powerGenerators)

        self._buildPowerProduct(
            sheet_new_energy=sheet_new_energy,
            statusList=statusList,
        )  # 提取到了非汇总类数据

        self._build_status_probability(sheet_status_probability=sheet_status_probability)

    """
	*******************************数据读取部分*************************************
	"""

    def _build_status_probability(self, sheet_status_probability):
        """self.month_status[month]=[probability of status index]"""
        for month in range(1, 13):
            month_name = f'{month}月'
            self.month_status[month] = sheet_status_probability[month_name].values.tolist()

    def _buildStatus(self, sheet_out, sheet_new_energy):
        # logging.info("初始状态列表")#Init Status list
        statusList = []
        for _, newenergy_row_values in sheet_new_energy.iterrows():
            if newenergy_row_values['电站ID'] == -99:
                for _, out_row_values in sheet_out.iterrows():
                    if out_row_values['场景ID'] == newenergy_row_values['场景ID'] and out_row_values['月份'] == \
                            newenergy_row_values['月份']:
                        statusList.append(Status(
                            sceneNum=out_row_values['场景ID'],
                            month=out_row_values['月份'],
                            powerDemand=out_row_values[2:],
                            totalCleanProduct=newenergy_row_values[2:26]
                        ))
        self.statusDict = StatusDict(statusList)
        return statusList

    def _buildScene(self, sheet_scene):
        """
		从表场景概率中读取场景ID作为一个集合，数据类型未改变
		:param sheet_scene:
		:return:
		"""
        # logging.info("初始化场景设置")#Init sceneNums set
        for _, scene_row_values in sheet_scene.iterrows():
            self.sceneNums.add(scene_row_values["场景ID"])

        for month in range(1, 13):
            self.month_status[month] = [0 for _ in range(len(self.sceneNums))]

    def _buildPowerGenerator(self, sheet_plant):
        # logging.info("初始化电源设置")#Init PowerGenerator set
        powerGenerators = []
        for _, plant_row_values in sheet_plant.iterrows():
            generator = power_generator.PowerGenerator(stationID=plant_row_values['电站ID'],
                                                       systemID=plant_row_values['所属系统ID'],
                                                       fuelConsumption=plant_row_values['燃料单耗'],
                                                       fuelPrice=plant_row_values['燃料单价'],
                                                       characterID=plant_row_values['特性ID'],
                                                       type=plant_row_values['电站类型'],
                                                       totalNums=plant_row_values['台数'],
                                                       techPower=plant_row_values['技术出力'],
                                                       dongtaitouzi=plant_row_values['动态投资'],
                                                       singleCapacity=plant_row_values['单机容量'],
                                                       zhuangji=plant_row_values['装机'],
                                                       yunweifeilv=plant_row_values['运维费率'],
                                                       yunxingfei=plant_row_values['运行费'],
                                                       )
            if generator.isCleanPowerStation():
                self.cleanPowerStations.append(generator)
            else:
                self.thermalPowerStations.append(generator)

            powerGenerators.append(generator)

        self.powerGeneratorDict = power_generator.PowerGeneratorDict(powerGenerators)
        return powerGenerators

    def _buildSheetVariables(self, sheet_output_attributes, sheet_system):
        """
		读取kr，VALID_NODE_ID,OUTPUT_NODE_ID
		kr为备用率，在输出特性表中
		目前认为的有效节点ID为表中"HUST节点ID"值为0、"类型"为101的节点
		输出节点ID为表中"HUST节点ID"值为0，"类型"为100的节点
		:param sheet_system:
		:return:
		"""
        # logging.info("Init Sheet Variables")
        # 初始化备用率
        self.kr = sheet_output_attributes.iloc[0, 14]

        # 初始化中间节点和输出节点
        for _, scene_row_values in sheet_system.iterrows():
            if scene_row_values["HUST节点ID"] == 0 and scene_row_values["类型"] == 101:
                self.VALID_NODE_ID.append(scene_row_values["系统ID"])
            elif scene_row_values["HUST节点ID"] == 0 and scene_row_values["类型"] == 100:
                self.OUTPUT_NODE_ID = scene_row_values["系统ID"]

        self.VALID_NODE_ID = list(set(self.VALID_NODE_ID))

    def _buildArcs(self, sheet_routine):
        logging.info("初始化线路列表")
        for col, routine_row_values in sheet_routine.iterrows():
            if (
                    routine_row_values["长度"] == 0
                    or routine_row_values["电阻"] == 0
                    or routine_row_values["回路数"] == 0
            ):
                continue
            # 计算电阻cost,以及容量,omega为单回路时的电阻，N为回路数
            omega = routine_row_values["长度"] * routine_row_values["电阻"]

            N = routine_row_values["回路数"]
            G = (1 / (omega * 2)) * int(N / 2) + (1 / omega) * (N % 2)

            cost = 1 / G

            nodeFrom = routine_row_values["首端ID"]
            nodeTo = routine_row_values["末端ID"]
            # capacity=routine_row_values["额定容量"]*math.ceil(N/2)
            capacity = routine_row_values["额定容量"] * N
            # 仅仅只考虑上层节点，即0-5之间的节点及线路
            if nodeFrom in self.VALID_NODE_ID + [
                self.OUTPUT_NODE_ID
            ] and nodeTo in self.VALID_NODE_ID + [self.OUTPUT_NODE_ID]:
                self.arcCost[nodeFrom, nodeTo] = cost
                self.arcCost[nodeTo, nodeFrom] = cost
                self.arcs.append(
                    graph.Arc(routine_row_values["线路ID"], nodeFrom, nodeTo, cost, capacity)
                )

    def _buildNodes(self, sheet_system, powerGenerators: [power_generator.PowerGenerator]):
        logging.info("初始化节点列表")  # Init nodeDict list
        powerGenerators = sorted(powerGenerators)
        node2commutation = {}
        for nodeID in self.VALID_NODE_ID + [self.OUTPUT_NODE_ID]:
            commutationMax, commutationMin = None, None
            for col, system_row_values in sheet_system.iterrows():
                if system_row_values["系统ID"] == nodeID:
                    commutationMax = system_row_values["换流Max"]
                    commutationMin = system_row_values["换流Min"]
            node2commutation[nodeID] = Bunch(
                commutationMin=commutationMin, commutationMax=commutationMax
            )

        _powerGeneratorList = {}
        # powerGenerators中已经进行了排序, 因此向列表中添加时也保持有序
        for powerGenerator in powerGenerators:
            if powerGenerator.systemID in self.VALID_NODE_ID + [self.OUTPUT_NODE_ID]:
                defaultDictList(_powerGeneratorList, powerGenerator.systemID, powerGenerator)
        nodes = []
        for nodeID in _powerGeneratorList:
            nodes.append(Node(
                nodeID=nodeID,
                powerGenerators=_powerGeneratorList[nodeID],
                commutationMax=node2commutation[nodeID].commutationMax,
                commutationMin=node2commutation[nodeID].commutationMin,
            ))
        self.nodeDict = NodeDict(nodes)
        return nodes

    def _buildPowerProduct(self, sheet_new_energy, statusList):
        # logging.info("初始化状态列表")#Init Power Product list
        for powerGenerator in self.cleanPowerStations:
            # 新能源电站
            for col, newenergy_row_values in sheet_new_energy.iterrows():
                if newenergy_row_values['电站ID'] not in [-99, -999]:
                    for status in statusList:
                        if status.sceneNum == newenergy_row_values['场景ID'] \
                                and status.month == newenergy_row_values['月份']:
                            if powerGenerator.stationID == newenergy_row_values['电站ID']:
                                self.powerProductors.append(
                                    power_productor.PowerProductor(status=status, powerGenerator=powerGenerator,
                                                                   cleanProduct=newenergy_row_values[2:-1].tolist()))

        # 火电站
        for status in statusList:
            for powerGenerator in self.thermalPowerStations:
                self.powerProductors.append(
                    power_productor.PowerProductor(status=status, powerGenerator=powerGenerator))

        self.powerProductDict = power_productor.PowerProductDict(self.powerProductors)
        self.powerProductList = power_productor.PowerProductList(self.powerProductors)

        return self.powerProductors

    """
	*******************************求解流程*****************************************
	"""

    def solve(self, debug=False):
        """
		整个流程函数
		:return:
		"""
        self.sortPath()
        for sceneNum in self.sceneNums:
            for month in range(1, 13):
                self.calculateDemand(sceneNum, month)
                self.assignOpenSchema(sceneNum, month)
                self.distributePower(sceneNum, month)
                self.checkPowerFlowOnArc(sceneNum, month)

                if debug:
                    break

    def calculateDemand(self, sceneNum, month):
        """
		第一步 算火电开发需求

		PTmax=max(PowerDemand-CleanProduct)
		Rl=max(PowerDemand)*kr
		C0=PTmax+Rl

		raise ValueError if the Status is empty list

		:return:中间结果 C0

		"""
        # logging.info("计算开机需求 ...")

        status = self.statusDict[sceneNum, month]
        # calculate the PTmax
        PTmax = max(
            [a - b for a, b in zip(status.powerDemand, status.totalCleanProduct)]
        )

        # calculate the Rl
        Rl = max(status.powerDemand) * self.kr

        # calculate the C0
        C0 = PTmax + Rl

        self.c0 = C0

        return C0

    def sortPath(self):
        """
		第二步 计算所有节点到外送节点的路径和其电阻，按照电阻从小到大排序 计算系数矩阵, 计算结果保留到中间状态字典
		:return:
		"""

        def _sortNodesByValidNodeID():
            """
			根据TOTAL_NODE_ID的顺序排序powerProductorList和nodeDict
			"""
            self.powerProductList.sortNodesByValidNodeIDs(self.TOTAL_NODE_ID)
            self.nodeDict.sortedNodesByValidNodeID(self.TOTAL_NODE_ID)

        # logging.info("路径排序 ...")
        self.graph = Graph(self, self.arcs)
        sortedValidNodeIDs = list(self.graph.sortedNodeIDs.keys())
        assert checkListSame(
            self.VALID_NODE_ID, sortedValidNodeIDs
        ), f"Expected nodeID result on VALID_NODE_ID is same to self.sortedValidNodeIDs,but got {self.VALID_NODE_ID}" \
           f" and {sortedValidNodeIDs} "
        assert checkListSame(
            self.VALID_NODE_ID + [self.OUTPUT_NODE_ID], self.nodeDict.totalNodeIDs
        ), f"Expected nodeID result on VALID_NODE_ID is same to nodeDict.totalNodeIDs,but got " \
           f"{self.VALID_NODE_ID + [self.OUTPUT_NODE_ID]} and {self.nodeDict.totalNodeIDs}"

        self.VALID_NODE_ID = sortedValidNodeIDs

        self.TOTAL_NODE_ID = [self.OUTPUT_NODE_ID] + self.VALID_NODE_ID
        _sortNodesByValidNodeID()

        logging.debug(self.graph)

        self._makeG(self.graph.adjacentBidirectional)

    def assignOpenSchema(self, sceneNum, month):
        """
		第三步 确定各个节点火电开机台数 N(j,i)

		:return:
		{"节点1": [(generator1, num1), (generator2, num2),...],...}

		"""

        # logging.info("分配火电开机 ...")

        def _assignOpenSchema(nodeID):
            """
			直接修改当前节点nodeID的所有火电电站开机的数量
			:param nodeID:
			:return:
			"""
            node = self.nodeDict[nodeID]
            commutationMin = node.commutationMin
            sumMinPV = sum(
                min(powerProductor.cleanProduct)
                for powerProductor in
                self.powerProductList.getCleanProductors(sceneNum=sceneNum, month=month, nodeID=nodeID)
            )

            sumMinPT = 0
            if sumMinPV < commutationMin:  # 如果新能源最小值小于换流容量最小值，则增加火电开机台数
                for stationID in node.thermalPowerStationIds:
                    powerProductor = self.powerProductDict[sceneNum, month, stationID]
                    generator = powerProductor.powerGenerator
                    techPower = powerProductor.techPowerMulCapacity
                    totalNums = generator.totalNums
                    if techPower * totalNums + sumMinPT >= commutationMin:  # 如果剩下的容量不够，则不用全开，只开一部分火电
                        powerProductor.open_schema = (commutationMin - sumMinPT) // techPower
                    else:
                        powerProductor.open_schema = totalNums

                    sumMinPT += techPower * powerProductor.open_schema

        def _getC(nodeID=None):
            """
			计算C(nodeID)，即火电电站目前已经开机的容量总和
			如果nodeID为None时，则计算所有的node的所有火电电站已经开机的容量之和
			否则，单独只计算节点nodeID的所有火电开机了的容量
			:param nodeID: 指定的某一个节点ID
			:return: 如果nodeID为None时，返回所有节点的电机容量之和，否则返回指定节点的容量
			"""

            if nodeID is None:
                return sum(self.powerProductList.getThermalCapacityUnderSchema(sceneNum, month, nodeID) for nodeID in
                           self.nodeDict.totalNodeIDs)
            else:
                return self.powerProductList.getThermalCapacityUnderSchema(sceneNum, month, nodeID)

        def _adjustNodeSchema(node, plus=False):

            """
			调整node的stationID节点的开机数量，可以加1（设置plus=True),也可以减1（设置plus=False)
			:param node:
			:param plus:如果为True，就在stationID的位置上新开机一个，否则新关机一个
			:return:
			如果当前节点node的所有火电站都调整到了极限一遍(nextStationID=None)，则下一步开始调整下一个节点
			如果当前节点node的当前火电站已经不可以继续开机或者关机（即达到了数量极限要求），则开始调整下一个火电站
			"""

            def _adjustStationSchema(stationID=None):
                """
				调整电站stationID的数量
				"""
                if stationID is None:
                    return False
                powerProductor = self.powerProductDict[sceneNum, month, stationID]
                canAdjust = powerProductor.adjustOneOpenSchema(plus)
                return canAdjust

            canAdjust = _adjustStationSchema(node.getCurrentThermalPowerStationID())
            if not canAdjust:
                # 当前火电站已经被调整到了极限，准备调整下一个火电站
                node.getNextThermalPowerStationID()

        # 检查一下是否火电电站全开满足火电需求
        self.canFetch, sumThermalProduct = self.checkSumThermalProductFetchDemand(month, sceneNum,
                                                                                  returnSumThermalProduct=True)
        if not self.canFetch:  # 如果不能满足需求，那只需要全部火电开满即可，不需要做调整
            self._setAllThermalProductorOpenSchemaToFull(sceneNum, month)
            logging.warning(
                f"场景{sceneNum}、月份{month}下,火电全开仍无法满足送电电力需求。火电总装机={sumThermalProduct},"
                f"电力需求={self.c0},需求偏差={sumThermalProduct - self.c0}")
            return

        # 确定各节点与新能源相匹配的最小开机台数
        for nodeID in self.TOTAL_NODE_ID:
            _assignOpenSchema(nodeID)

        # 开始调整，使得各节点开机容量之和尽可能接近（>=）开机需求
        _getDC = lambda: _getC() - self.c0

        nodeIndex = 0
        node = self.nodeDict[self.TOTAL_NODE_ID[nodeIndex]]
        Ci = _getC(self.TOTAL_NODE_ID[nodeIndex])
        dC = _getDC()

        countNoAssign = 0
        while dC < 0 or dC >= Ci:
            if dC < 0:
                if countNoAssign > len(self.TOTAL_NODE_ID):  # 一直在不断的开新的节点的电站，已经全部开完了，就不用继续死循环下去，开机容量已经最接近目标值了
                    break
                # 从1节点开始依次增加火电开机
                _adjustNodeSchema(
                    node,
                    plus=True
                )
                if self.powerProductList.isFull(sceneNum, month, self.TOTAL_NODE_ID[nodeIndex]):
                    # 当前节点已经更新完了，换到下一个节点
                    countNoAssign += 1
                    nodeIndex = (nodeIndex + 1) % len(self.TOTAL_NODE_ID)
                    node = self.nodeDict[self.TOTAL_NODE_ID[nodeIndex]]

            elif dC >= Ci:
                # 从5开始减少火电开机
                _adjustNodeSchema(
                    node,
                    plus=False
                )
                countNoAssign = 0
                if self.powerProductList.isEmpty(sceneNum, month, self.TOTAL_NODE_ID[nodeIndex]):
                    # 当前节点已经更新完了，换到下一个节点
                    nodeIndex = (nodeIndex - 1 + len(self.TOTAL_NODE_ID)) % len(
                        self.TOTAL_NODE_ID
                    )
                    node = self.nodeDict[self.TOTAL_NODE_ID[nodeIndex]]

            # 更新开关机状态，便于判断Ci和dC的结果
            Ci = _getC(self.TOTAL_NODE_ID[nodeIndex])
            dC = _getDC()

    def checkSumThermalProductFetchDemand(self, month, sceneNum, returnSumThermalProduct=True):
        # 检查火电全开的情况下的总出力
        sumThermalProduct = 0
        for nodeID in self.TOTAL_NODE_ID:
            maxThermalPower = self.powerProductList.maxThermalPower(sceneNum, month, nodeID)
            logging.debug(f'节点{nodeID}, 火电最大出力={maxThermalPower}')
            sumThermalProduct += maxThermalPower
        checkFetch = sumThermalProduct >= self.c0
        if returnSumThermalProduct:
            return checkFetch, sumThermalProduct
        else:
            return checkFetch

    def distributePower(self, sceneNum, month):
        """
		第四步，确定各节点的24小时可调功率
		:param sceneNum:
		:param month:
		:return:
		"""

        def _cumulateForcedPower() -> (float, Dict[int, float]):
            """
			根据火电开机信息计算当前status下的火电强迫出力之和
			使用到self.intermediate 中间结果的存储对象，包括了在当前status下的火电需求，各节点火电开机信息
			:return: 当前status下的火电强制出力之和, 各个节点的火电强制出力
			"""
            forcePowerTotal = 0.0
            forcePowerPerNode = {}
            for nodeId, products in self.powerProductList[sceneNum, month].items():
                powerPerNode = 0.
                for product in products:
                    powerPerNode += product.forceThermalPowerUnderSchema
                forcePowerPerNode[nodeId] = powerPerNode
                forcePowerTotal += powerPerNode

            return forcePowerTotal, forcePowerPerNode

        def _calculateTotalPowerToDistribute(currentStatus: Status, forcePower: float) -> List[float]:
            """
			计算逐小时待分配火电总功率
			DeltaPT = P_{外送} - PV_t - sum PT_min

			:param currentStatus: 总的分小时店里需求和
			:param forcePower:
			:return: len=24的列表
			"""
            return [
                a - b - forcePower
                for a, b in zip(currentStatus.powerDemand, currentStatus.totalCleanProduct)
            ]

        def _distributeRestPower(restPowerToDistribute: List[float]) -> NoReturn:
            """
			分配待分配功率，
			逐小时进行分配，首先获得当前小时需要分配的电力，按照节点顺序访问节点，如果当前节点能够继续增加分配
			:param restPowerToDistribute:
			:return:
			"""
            currentPowerProducts = self.powerProductList[sceneNum, month]
            for hour in range(24):
                restPower = restPowerToDistribute[hour]  # 待分配电力，<0就是False，说明弃电
                stepNoDistribute = 0  # 没有成功分配的记录
                dropPower = False  # restPower是否表示的是弃电
                if restPower < 0:
                    # 强迫出力超过需求，需要弃电
                    index = len(self.TOTAL_NODE_ID) - 1  # 从最后一个节点ID开始往前弃
                    restPower = -restPower
                    dropPower = True
                else:
                    # 强迫出力小于需求，需要分配剩余电力需求
                    index = 0
                while restPower > 0:
                    nodeId = self.TOTAL_NODE_ID[index]
                    node = self.nodeDict[nodeId]
                    if dropPower:
                        nodeCanDrop = min(restPower, node.canDrop(currentPowerProducts[nodeId], hour))
                        if nodeCanDrop > 0:
                            success = node.dropBy(self.powerProductDict, sceneNum, month, hour, nodeCanDrop)
                            assert success
                            restPower -= nodeCanDrop
                            stepNoDistribute = 0
                        else:
                            stepNoDistribute += 1
                    else:
                        # 增出力过程
                        maxIncrease = node.canIncrease(currentPowerProducts[nodeId], hour)
                        increaseStep = min(restPower, STEP_IN_POWER_DISTRIBUTION, maxIncrease)
                        if increaseStep > 0:
                            success = node.increaseBy(self.powerProductDict, sceneNum, month, hour, increaseStep)
                            assert success
                            restPower -= increaseStep
                            stepNoDistribute = 0
                        else:
                            stepNoDistribute += 1

                    if stepNoDistribute > len(currentPowerProducts):  # 遍历所有节点后仍不能分配任何出力
                        raise AssertionError(f"当前时刻无法分配出力 (场景：{sceneNum}, 月份：{month}), 时刻：{hour}")
                    # 弃电时从最远的开始递减，加电时从最近的开始增加
                    index = (index - 1 + len(self.TOTAL_NODE_ID)) % len(self.TOTAL_NODE_ID) if dropPower \
                        else (index + 1) % len(self.TOTAL_NODE_ID)

        systemInput = self.statusDict[sceneNum, month]
        logging.info(f"(场景{sceneNum}, 月份{month})下的出力分配...")
        # 计算火电总强制出力，火电各节点强制出力
        totalForcePower, nodeId2ForcePower = _cumulateForcedPower()
        # 根据当天的24小时外送需求-总的清洁能源出力和火电当天的总出力，得到24小时的待分配电力
        totalPowerToDistribute = _calculateTotalPowerToDistribute(systemInput, totalForcePower)
        # 实际分配待分配电力
        _distributeRestPower(totalPowerToDistribute)

    def checkPowerFlowOnArc(self, sceneNum, month):
        """
		第五步 校验线路潮流
		:return:
		"""

        def calculateGx0():
            gx = []
            outputNode = self.OUTPUT_NODE_ID
            for nodeId in sorted(self.VALID_NODE_ID):
                if self.arcCost[nodeId, outputNode] > 0:
                    gx.append(1 / self.arcCost[nodeId, outputNode])
                else:
                    gx.append(0)
            return gx

        def calculateVoltage(powerOnNode: List, gx) -> List[float]:
            """
			根据各节点电力输出计算各节点电压
			:param powerOnNode: 按节点的出力
			:param gx: x-0的电阻倒数
			:return: 1-5节点的电压
			"""
            powerOnNode = [powerOnNode[i] for i in sorted(self.VALID_NODE_ID)]  # 获得非输出节点从小到大排序, 问题中是取powerNode[1:]
            voltage = np.ones(5, np.float) * V0_IN_FLOW_CHECK
            powerVector = np.array(list(reversed(powerOnNode)))  # (P5, P4, P3, P2, P1)
            biasVector = np.array(list(reversed(gx)))
            gInv = np.linalg.inv(self.matrixG)
            for _ in range(3):
                voltage = gInv.dot((powerVector / (2 * voltage) + biasVector * voltage))
            voltage = voltage[::-1].tolist()
            return voltage

        def calPowerOnArc(voltageOnNode):
            """
			计算线路潮流
			:param voltageOnNode:
			:return:
			"""
            for a in self.arcs:
                vFirst = voltageOnNode[a.nodeFrom]
                vSecond = voltageOnNode[a.nodeTo]
                # 为负时表示反向
                a.powerFlow = max(vFirst, vSecond) * (vFirst - vSecond) / a.cost

        def checkOneHour(thisHour):
            """
			单个小时的边功率检查逻辑

			:param thisHour:
			:return:
			"""
            # 各节点功率
            powerPerNode = [self.powerProductList.getCurrentPowerByNode(nodeId, sceneNum, month, thisHour)
                            for nodeId in sorted(self.TOTAL_NODE_ID)]
            # 计算节点的电压
            voltage = calculateVoltage(powerPerNode, gx0)
            voltage.insert(self.OUTPUT_NODE_ID, V0_IN_FLOW_CHECK)
            # 计算边上的功率
            calPowerOnArc(voltage)
            # 遍历所有边，判断是否存在超限，存在超限时，从超限最少的节点开始调整
            minInvalidArc = None
            minInvalidValue = sys.maxsize
            # 找到越限的边，如果不存在返回检查完成，如果存在调整越限最小的边（从最少调整，最小消除后更大越限的可能减少)
            for a in self.arcs:
                exceed = abs(a.powerFlow) - a.capacity
                if 1e-3 < exceed < minInvalidValue:
                    minInvalidArc = a
                    minInvalidValue = exceed

            if minInvalidArc is not None:
                logging.info(
                    f"场景:({sceneNum},{month}月,{hour}时)存在功率越限线路, 受限线路ID: {minInvalidArc.arcID}, "
                    f"越限值: {minInvalidValue}")
                adjustPowerOnArc(minInvalidArc, minInvalidValue, thisHour)
                return False

            # 将未超限的边上功率存储到结构
            for a in self.arcs:
                a.savePowerOnArc(sceneNum, month, hour)

            return True

        def adjustPowerOnArc(minInvalidArc: graph.Arc, minInvalidValue: float, thisHour: int):
            """
			调整边上超限的电力，从通过超限边的最远的最不经济的电站开始减少出力，通过待调整的超限的边的节点最短路中所包含的
			边构成的边集与整个边集去差集，在这部分边中从最短路包含该边的节点数最多的边（也就是最近的边）开始调整，从最短路通过
			这条边的最近的节点的最经济的电站开始增加出力

			:param minInvalidArc:
			:param minInvalidValue:
			:param thisHour:
			:return:
			"""
            # 从最远节点的最不经济电站开始减少出力
            powerToDecrease = minInvalidValue
            for nodeId in reversed(minInvalidArc.minPathContained):
                if powerToDecrease <= 0:
                    break
                node = self.nodeDict[nodeId]
                nodeCanDecrease = min(node.canDecrease(self.powerProductList[sceneNum, month][nodeId], thisHour),
                                      powerToDecrease)
                if nodeCanDecrease > 0:
                    success = node.decreaseBy(self.powerProductDict, sceneNum, month, thisHour, nodeCanDecrease)
                    assert success
                    powerToDecrease -= nodeCanDecrease

            if powerToDecrease > 0:
                raise ValueError(f"无法降低线路潮流, 线路ID:{minInvalidArc.arcID}, 越限值: {powerToDecrease}")

            # 从没有超限的边中选择最近的最经济的电站开始增加
            powerToIncrease = minInvalidValue
            candidateArcs = []
            # 找到 排除节点最短路经过超限边的节点后，剩下的arc，这部分arc与超限边不共享前往出力节点的路径
            for a in self.arcs:
                if len(set(a.minPathContained).intersection(set(minInvalidArc.minPathContained))) == 0 \
                        and abs(a.powerFlow) < a.capacity:
                    candidateArcs.append(a)
            # 对这部分边按照通过节点数的多少进行排序，离外送节点越近的边，把这条边作为最短路径的节点越多，从通过这个边的节点开始增加更加经济
            candidateArcs.sort(key=lambda aa: len(aa.minPathContained), reverse=True)
            for a in candidateArcs:
                if powerToIncrease <= 0:
                    break
                for nodeId in a.minPathContained:
                    if powerToIncrease <= 0:
                        break
                    node = self.nodeDict[nodeId]
                    nodeCanIncrease = min(node.canIncrease(self.powerProductList[sceneNum, month][nodeId], thisHour),
                                          powerToIncrease)
                    if nodeCanIncrease > 0:
                        success = node.increaseBy(self.powerProductDict, sceneNum, month, thisHour, nodeCanIncrease)
                        assert success
                        powerToIncrease -= nodeCanIncrease

            if powerToIncrease > 0:
                raise ValueError(f"无法增加线路潮流, 需求量: {powerToIncrease}")

        # 计算各节点到外送节点的电阻的倒数，不直接连通的为0
        gx0 = calculateGx0()
        # 逐小时进行检查，当同一个小时调整三次时开始报提示，同一个小时调整5次依然无法通过时，报错
        for hour in range(24):
            # 计算节点的功率
            checkCount = 0
            while not checkOneHour(hour):
                checkCount += 1
                if checkCount == 3:
                    logging.warning(f"场景{sceneNum}, {month}月, {hour}时: 校验线路潮流  迭代次数：3")
                if checkCount == 5:
                    raise ValueError(f"场景{sceneNum}, {month}月, {hour}时校核出现不收敛")
        logging.debug(f"场景{sceneNum},{month}月潮流校核完成")

    def checkSolution(self, solution, sceneNum, month):
        """
		交叉验证解的合法性，各小时节点上层出力之和 = 电力需求，节点的火电+新能源-弃电 = 上层出力

		:param solution:
		:param sceneNum:
		:param month:
		:return:
		"""
        npPowerAbandon = np.array(solution.abandonOnNode)
        powerTotalAbandon = np.sum(npPowerAbandon[:, :-1], axis=0)
        npPowerOnStation = np.array(solution.thermalPowerOnStation)
        powerTotalThermal = np.sum(npPowerOnStation[:, :-1], axis=0)
        npPowerDemand = np.array(self.statusDict[sceneNum, month].powerDemand)
        npCleanPower = np.array(self.statusDict[sceneNum, month].totalCleanProduct)
        npPowerOnNode = np.array(solution.powerOnNode)
        powerProduce = np.sum(npPowerOnNode[:, :-1], axis=0)
        epsilon1 = sum(npCleanPower + powerTotalThermal - powerTotalAbandon - npPowerDemand)
        epsilon2 = sum(powerProduce - npPowerDemand)
        return epsilon1 == 0 and epsilon2 == 0

    def afterProcess(self):
        """
		后处理, 将powerProducts结构中的数据抽取成矩阵，如果当前为Debug模式，进行额外的结果检查
		:return:
		"""
        self.solutions = self._getSolutions()
        # self.solutions = self.matrixchange()
        if logging.getLogger(__name__).getEffectiveLevel() == logging.DEBUG:
            for sceneNum in self.sceneNums:
                for month in range(1, 13):
                    result = self.checkSolution(self.solutions[sceneNum, month], sceneNum, month)
                    if not result:
                        logging.debug(f"solutionCheckResult: ({sceneNum}, {month}) false")
        # logging.debug("after process solution check complete")

        return self.solutions

    """
	*******************************其他辅助方法*****************************************
	"""

    def getSolutionByStatus(self, sceneNum, month):
        """
		根据场景码和月份获取对应的解
		:param sceneNum:
		:param month:
		:return: [openSchema, thermalPowerOnStation, powerOnNode, abandon] 开机方案(12,), 火电电站的总出力(12, 25)，
			各节点的上层出力(25,6), 节点弃电(25,6)
		"""
        return self.solutions[sceneNum, month][:-1]

    def getPLineByStatus(self, sceneNum, month):
        sol = self.getSolutionByStatus(sceneNum, month)
        return sol[0:2], sol[4]

    def _makeG(self, graph: dict):
        """
		生成电导矩阵

		:param graph: 邻接表
		:return:
		"""
        logging.info("Calculate Matrix G...")
        g = lambda x, y: 1.0 / self.arcCost[x, y]

        G = np.ones((len(self.VALID_NODE_ID), len(self.VALID_NODE_ID)), dtype=float)
        for _i in range(len(self.VALID_NODE_ID)):
            for _j in range(len(self.VALID_NODE_ID)):
                i, j = len(self.VALID_NODE_ID) - _i, len(self.VALID_NODE_ID) - _j
                if i == j:
                    G[_i, _j] = sum(g(i, y.nodeTo) for y in graph[i])
                else:
                    i, j = max(i, j), min(i, j)
                    G[_i, _j] = -g(i, j) if (i, j) in self.arcCost else 0
        self.matrixG = G
        return G

    def _getSolutions(self):
        """
		实际从powerProducts中抽取结果矩阵
		:return:
		"""

        def getOneSolution():
            """

			:return:
			[
			openSchemaOnStation, 开机方案(12,),
			forcePowerOnStation, 强迫出力(12,),
			thermalPowerOnStation, 火电电站的总出力(12, 25)，
			cleanPowerOnStation，风光电站上层的总新能源出力(12,25),
			profitOnStation, 火电电站的盈亏（12,25）,
			powerOnNode, 各节点的上层出力(6,25),
			abandonOnNode,节点弃电(6,25),
			cleanPowerOnNode, 上层各节点的新能源出力(6,25),
			thermalPowerOnNode，上层各节点的火电出力(6,25),
			]
			"""
            openSchemaOnStation = [[0, 0]] * numOfThermalGenerator  # 开机方案
            forcePowerOnStation = [[0, 0]] * numOfThermalGenerator  # 强迫出力（最小出力）

            thermalPowerOnStation = []  # 电站的火电出力
            cleanPowerOnStation = []  # 电站的风光电出力
            profitOnStation = []  # 火电电站的盈亏

            powerOnNode = []  # 节点的出力
            abandonOnNode = []  # 节点弃电
            cleanPowerOnNode = []  # 节点弃电
            thermalPowerOnNode = []  # 节点的火电出力

            for nodeId, products in self.powerProductList[sceneNum, month].items():  # 根据场景和月份提取24h出力等状态量
                powerThisNode = [0] * 24
                abandonThisNode = [0] * 24
                cleanPowerThisNode = [0] * 24
                thermalPowerThisNode = [0] * 24
                for product in products:
                    for h in range(24):
                        powerThisNode[h] += product.currentOutputAtHour(h, highLevelOnly=True)
                        abandonThisNode[h] += product.abandonDistribute[h]
                        cleanPowerThisNode[h] += product.NewenergythisNode(h)
                        thermalPowerThisNode[h] += product.currentThermalAtHour(h)
                    if not product.isCleanProductor():
                        openSchemaOnStation[stationId2OpenSchemaIndex[product.powerGenerator.stationID]] = [
                            product.open_schema, product.powerGenerator.stationID]  # 开机方案
                        forcePowerOnStation[stationId2OpenSchemaIndex[product.powerGenerator.stationID]] = [
                            product.forceThermalPowerUnderSchema, product.powerGenerator.stationID]  # 开机方案

                        # 计算没有使用product提供的高层封装函数是希望通过临近访问命中缓存进行加速
                        forcePower = product.forceThermalPowerUnderSchema
                        powerThisThermalStation = [forcePower] * 24
                        profitOnThisStation = [0] * 24
                        for h in range(24):
                            powerThisThermalStation[h] += product.distributeSchema[h]
                            profitOnThisStation[h] += product.abandonDistribute[h]
                        powerThisThermalStation.append(product.powerGenerator.stationID)
                        profitOnThisStation.append(product.powerGenerator.stationID)
                        thermalPowerOnStation.append(powerThisThermalStation)
                        profitOnStation.append(profitOnThisStation)
                    else:
                        powerThisCleanStation = [0] * 24
                        for h in range(24):
                            powerThisCleanStation[h] += product.cleanProduct[h]
                        powerThisCleanStation.append(product.powerGenerator.stationID)
                        cleanPowerOnStation.append(powerThisCleanStation)

                powerThisNode.append(nodeId)
                abandonThisNode.append(nodeId)
                cleanPowerThisNode.append(nodeId)
                thermalPowerThisNode.append(nodeId)

                powerOnNode.append(powerThisNode)
                abandonOnNode.append(abandonThisNode)
                cleanPowerOnNode.append(cleanPowerThisNode)
                thermalPowerOnNode.append(thermalPowerThisNode)

            openSchemaOnStation.sort(key=lambda p: p[-1])
            forcePowerOnStation.sort(key=lambda p: p[-1])

            thermalPowerOnStation.sort(key=lambda p: p[-1])
            cleanPowerOnStation.sort(key=lambda p: p[-1])
            profitOnStation.sort(key=lambda p: p[-1])

            powerOnNode.sort(key=lambda p: p[-1])
            abandonOnNode.sort(key=lambda p: p[-1])
            cleanPowerOnNode.sort(key=lambda p: p[-1])
            thermalPowerOnNode.sort(key=lambda p: p[-1])

            powerOnArcs = []
            for a in self.arcs:
                powerOnThisArcInSolution = a.getPowerOnArc(sceneNum, month).copy()
                powerOnThisArcInSolution += [a.arcID, a.nodeFrom, a.nodeTo]
                powerOnArcs.append(powerOnThisArcInSolution)

            powerOnArcs.sort(key=lambda p: p[24])

            return Bunch(openSchemaOnStation=openSchemaOnStation,
                         forcePowerOnStation=forcePowerOnStation,

                         thermalPowerOnStation=thermalPowerOnStation,
                         profitOnStation=profitOnStation,
                         cleanPowerOnStation=cleanPowerOnStation,

                         powerOnNode=powerOnNode,
                         abandonOnNode=abandonOnNode,
                         cleanPowerOnNode=cleanPowerOnNode,
                         thermalPowerOnNode=thermalPowerOnNode,

                         powerOnArcs=powerOnArcs,
                         )

        solutions = {}
        stationId2OpenSchemaIndex = dict([(stationId, index) for index, stationId in
                                          enumerate(self.powerGeneratorDict.thermalGeneratorIds)])
        numOfThermalGenerator = len(self.powerGeneratorDict.thermalGeneratorIds)
        for sceneNum in self.sceneNums:
            for month in range(1, 13):
                solutions[(sceneNum, month)] = getOneSolution()

        return solutions

    def matrixchange(self):
        def onescene():
            openSchemaOnStation = [[0, 0]] * numOfThermalGenerator  # 开机方案
            forcePowerOnStation = [[0, 0]] * numOfThermalGenerator  # 强迫出力（最小出力）
            maxPowerOnStation = [[0, 0]] * numOfThermalGenerator  # 最大出力

            powerOnNode = []  # 节点的注入功率
            abandonOnNode = []  # 节点弃电
            cleanPowerOnNode = []  # 节点风光联合出力
            thermalPowerOnNode = []  # 节点的火电出力
            maxpower = []  # 节点火电最大功率
            minpower = []  # 节点火电最小功率

            for nodeId, products in self.powerProductList[sceneNum, month].items():  # 根据场景和月份提取24h出力等状态量
                powerThisNode = [0] * 24  # 节点注入功率
                abandonThisNode = [0] * 24  # 弃电
                cleanPowerThisNode = [0] * 24  # 风光联合出力
                thermalPowerThisNode = [0] * 24  # 火电出力
                powerThisNode.insert(0, nodeId)
                abandonThisNode.insert(0, nodeId)
                cleanPowerThisNode.insert(0, nodeId)
                thermalPowerThisNode.insert(0, nodeId)
                powerThisNode.insert(1, 'B注入功率')
                abandonThisNode.insert(1, 'A弃电')
                cleanPowerThisNode.insert(1, 'D风光出力')
                thermalPowerThisNode.insert(1, 'C火电出力')

                for product in products:
                    for h in range(24):
                        powerThisNode[h + 2] += product.currentOutputAtHour(h, highLevelOnly=True)
                        abandonThisNode[h + 2] += product.abandonDistribute[h]
                        cleanPowerThisNode[h + 2] += product.NewenergythisNode(h)
                        thermalPowerThisNode[h + 2] += product.currentThermalAtHour(h)
                    if not product.isCleanProductor():
                        openSchemaOnStation[stationId2OpenSchemaIndex[product.powerGenerator.stationID]] = [
                            product.open_schema, product.powerGenerator.stationID]  # 开机方案
                        forcePowerOnStation[stationId2OpenSchemaIndex[product.powerGenerator.stationID]] = [
                            product.forceThermalPowerUnderSchema, product.powerGenerator.stationID]  # 最小出力
                        maxPowerOnStation[stationId2OpenSchemaIndex[product.powerGenerator.stationID]] = [
                            product.maxthermalPower, product.powerGenerator.stationID]  # 最大出力

                powerThisNode.insert(2, max(powerThisNode[2:26]))
                powerThisNode.insert(3, min(powerThisNode[3:27]))
                powerThisNode.insert(4, mean(powerThisNode[4:28]))
                abandonThisNode.insert(2, max(abandonThisNode[2:26]))
                abandonThisNode.insert(3, min(abandonThisNode[3:27]))
                abandonThisNode.insert(4, mean(abandonThisNode[4:28]))
                cleanPowerThisNode.insert(2, max(cleanPowerThisNode[2:26]))
                cleanPowerThisNode.insert(3, min(cleanPowerThisNode[3:27]))
                cleanPowerThisNode.insert(4, mean(cleanPowerThisNode[4:28]))
                thermalPowerThisNode.insert(2, 0)
                thermalPowerThisNode.insert(3, 0)
                thermalPowerThisNode.insert(4, mean(thermalPowerThisNode[4:28]))

                powerOnNode.append(powerThisNode)
                abandonOnNode.append(abandonThisNode)
                cleanPowerOnNode.append(cleanPowerThisNode)
                thermalPowerOnNode.append(thermalPowerThisNode)

            openSchemaOnStation.sort(key=lambda p: p[-1])
            forcePowerOnStation.sort(key=lambda p: p[-1])
            maxPowerOnStation.sort(key=lambda p: p[-1])

            powerOnNode.sort(key=lambda p: p[0])
            abandonOnNode.sort(key=lambda p: p[0])
            cleanPowerOnNode.sort(key=lambda p: p[0])
            thermalPowerOnNode.sort(key=lambda p: p[0])

            powerOnArcs = []
            for a in self.arcs:
                powerOnThisArcInSolution = a.getPowerOnArc(sceneNum, month).copy()
                powerOnThisArcInSolution.insert(0, a.nodeFrom)
                powerOnThisArcInSolution.insert(1, 'L' + str(a.nodeFrom) + '-' + str(a.nodeTo))
                powerOnThisArcInSolution.insert(2, max(a.getPowerOnArc(sceneNum, month).copy()))
                powerOnThisArcInSolution.insert(3, min(a.getPowerOnArc(sceneNum, month).copy()))
                powerOnThisArcInSolution.insert(4, a.capacity)
                # powerOnThisArcInSolution += [a.arcID, a.nodeFrom, a.nodeTo]
                powerOnArcs.append(powerOnThisArcInSolution)

            powerOnArcs.sort(key=lambda p: p[24])

            return Bunch(openSchemaOnStation=openSchemaOnStation,
                         forcePowerOnStation=forcePowerOnStation,
                         maxPowerOnStation=maxPowerOnStation,

                         abandonOnNode=abandonOnNode,
                         powerOnNode=powerOnNode,
                         thermalPowerOnNode=thermalPowerOnNode,
                         cleanPowerOnNode=cleanPowerOnNode,

                         powerOnArcs=powerOnArcs,
                         )

        solutions = {}
        stationId2OpenSchemaIndex = dict([(stationId, index) for index, stationId in
                                          enumerate(self.powerGeneratorDict.thermalGeneratorIds)])
        numOfThermalGenerator = len(self.powerGeneratorDict.thermalGeneratorIds)
        for sceneNum in self.sceneNums:
            for month in range(1, 13):
                solutions[(sceneNum, month)] = onescene()
        return solutions

    def _setAllThermalProductorOpenSchemaToFull(self, sceneNum, month):
        """
		设置所有的火电站的开机数为满开
		"""
        for powerProductor in self.powerProductList.getThermalProductors(sceneNum, month):
            powerProductor.open_schema = powerProductor.powerGenerator.totalNums

    def _collectCleanProductsByNode(self, n: Node, sceneNum: int, month: int) -> List[List[int]]:
        """
		按照图上的节点收集该节点的所有新能源电站的24小时出力曲线，得到个电站24小时出力列表的列表
		:param n: 给定图上节点
		:param sceneNum: 场景号
		:param month: 月份
		:return: 24小时出力列表的列表
		"""
        cleanProducts = []
        for stationID in n.cleanPowerStationIds:
            # if (sceneNum, month, stationID) in self.powerProductDict:
            # cleanProduct出力曲线中已经考虑了station的数量
            cleanProduct = self.powerProductDict[sceneNum, month, stationID].cleanProduct
            cleanProducts.append(cleanProduct)
        return cleanProducts

    def _maxCleanProductByNode(self, n: Node, sceneNum: int, month: int) -> int:
        """
		按照节点查询该节点的24小时清洁能源出力中的最大值
		:param n: 图上节点
		:param sceneNum: 场景号
		:param month: 月份
		:return: 最大值
		"""
        cleanProducts = self._sumCleanProductByNode(n, sceneNum, month)
        return max(cleanProducts)

    def _sumCleanProductByNode(self, n: Node, sceneNum: int, month: int) -> List[int]:
        """
		计算节点下的所有清洁发电站的每小时出力之和，得到该节点的24小时的清洁能源出力
		:param n: 图上节点
		:param sceneNum: 场景号
		:param month: 月份
		:return: 长度24的int列表，该节点的24小时清洁能源出力曲线
		"""
        cleanProducts = self._collectCleanProductsByNode(n, sceneNum, month)
        return np.array(cleanProducts).sum(axis=0).tolist()


def get_result(debug=False):
    from results.power_result import PowerResult
    if os.path.exists(power_result_filename):
        power_result = load_object(power_result_filename)
    else:
        power_result = PowerResult(solver, debug)
        power_result.build_data(date_type=False)
        save_object(power_result, power_result_filename)

    power_result.to_excel(0, 1, False, power_result_excel)
    print(power_result)

    if os.path.exists(electricity_result_filename):
        electricity_result = load_object(electricity_result_filename)
    else:
        electricity_result = ElectricityResult(solver, power_result, debug)
        save_object(electricity_result, electricity_result_filename)

    electricity_result.to_excel(False, electricity_result_excel)
    print(electricity_result)

    if os.path.exists(technical_statics_result_filename):
        technical_statics_result = load_object(technical_statics_result_filename)
    else:
        technical_statics_result = TechnicalStaticsResult(solver, power_result, electricity_result, debug)
        save_object(technical_statics_result, technical_statics_result_filename)

    technical_statics_result.to_excel(False, technical_statics_result_excel)
    print(technical_statics_result)


if __name__ == "__main__":
    debug = False
    pkl_dir = 'pkl/'
    os.makedirs(pkl_dir, exist_ok=True)
    solver_filename = os.path.join(pkl_dir, 'solver.pkl')
    power_result_filename = os.path.join(pkl_dir, 'power_result.pkl')
    electricity_result_filename = os.path.join(pkl_dir, 'electricity_result.pkl')
    technical_statics_result_filename = os.path.join(pkl_dir, 'technical_statics_result.pkl')

    excel_dir = 'results/'
    os.makedirs(excel_dir, exist_ok=True)
    power_result_excel = os.path.join(excel_dir, 'power_result.csv')
    electricity_result_excel = os.path.join(excel_dir, 'electricity_result.csv')
    technical_statics_result_excel = os.path.join(excel_dir, 'technical_statics_result.csv')

    if os.path.exists(solver_filename):
        solver = load_object(solver_filename)
    else:
        solver = Solver()
        solver.solve(debug=False)
    save_object(solver, solver_filename)

    get_result(debug)
