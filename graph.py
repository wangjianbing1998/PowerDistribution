# -*- coding:utf-8 -*-
# @Time: 2020/12/16 12:21 下午
# @Author:
# @File: graph.py
from collections import OrderedDict
from collections import defaultdict
from typing import List
from typing import NoReturn


class Arc:
    def __init__(self, ID, nodeFrom, nodeTo, cost, capacity, flow=0):
        # arcID为线路ID，来自算例.xlsx-线路表-线路ID，已经转为int型
        # nodeFrom为起始点，来自算例.xlsx-线路表-首端ID，已经转为int型
        # nodeTo为终止点，来自算例.xlsx-线路表-末端ID，已经转为int型
        # cost为电阻，来自算例.xlsx-线路表，回路数×长度×电阻
        # capacity为最大容量，来自算例.xlsx-线路表，回路数×额定容量
        # flow为当前流量
        # gradient为当前电阻cost求电导逆阵 g=1/cost
        # minPathContained记录了将该边作为最短路的一部分的节点列表
        self.arcID = ID
        self.nodeFrom = nodeFrom
        self.nodeTo = nodeTo
        self.cost = cost
        self.capacity = capacity
        self.minPathContained = []
        self.powerOnArcDict = {}

        # 有状态属性
        self.powerFlow = flow

        self.gradient = 1.0 / self.cost

    def __repr__(self):
        return f"arcID={self.arcID}, nodeFrom={self.nodeFrom}, nodeTo={self.nodeTo}, cost={self.cost}, flow={self.powerFlow}, capacity={self.capacity}\n"

    def reInit(self):
        self.powerFlow = 0

    def __reversed__(self):
        return Arc(
            ID=self.arcID,
            nodeFrom=self.nodeTo,
            nodeTo=self.nodeFrom,
            cost=self.cost,
            capacity=self.capacity,
            flow=self.powerFlow,
        )

    def savePowerOnArc(self, sceneNum, month, hour):
        if (sceneNum, month) not in self.powerOnArcDict:
            self.powerOnArcDict[sceneNum, month] = [0] * 24
        self.powerOnArcDict[sceneNum, month][hour] = self.powerFlow

    def getPowerOnArc(self, sceneNum, month):
        if (sceneNum, month) in self.powerOnArcDict:
            return self.powerOnArcDict[sceneNum, month]

        return None


class Path:
    """
    描述电网中的两个节点之间的路径，可以由多条边共同构成，路径的最大容量由构成边中的最小容量决定，
    路径用于排序
    """

    def __init__(self, arcs: List[Arc] = None):
        self.arcs = arcs  # 路径对应的组成边
        self.capacity = 0
        self.nodeFrom = None
        self.nodeTo = None
        self.cost = 0

        self.updatePath()

    def __repr__(self):
        s = f"cost={self.cost}\t"
        for arc in self.arcs:
            s += str(arc.nodeFrom) + "->"
        if len(self.arcs):
            s += str(self.arcs[-1].nodeTo)
        return s

    def __lt__(self, other):
        return self.cost < other.cost

    def updatePath(self, arcs=None):
        """
        当路径中的arcs发生变化的时候，需要重新更新一下所有变量
        :param arcs: 如果指定了新的arcs，那么self.arcs就会被替换
        :return:
        """
        if arcs:
            self.arcs = arcs
        if self.arcs:
            self.capacity = min(a.capacity for a in self.arcs)
            self.nodeFrom = self.arcs[0].nodeFrom  # 路径起始点
            self.nodeTo = self.arcs[-1].nodeTo  # 路径结束点
            self.cost = sum(a.cost for a in self.arcs)  # 路径总距离 cost

    def updatePathToArc(self) -> NoReturn:
        """
        将path的起点注册到arc的minPathContained当中，用于校核步骤中，在超限时调整节点出力

        :return:
        """
        for arc in self.arcs:
            arc.minPathContained.append(self.nodeFrom)

    def addArc(self, arc) -> bool:
        """
        在最后添加一条边
        :return: 是否添加成功
        """
        if arc.nodeFrom != self.nodeTo:
            return False
        self.arcs.append(arc)
        self.updatePath()
        return True


class Graph:
    def __init__(self, config, arcs):
        self.config = config
        self.arcs = arcs
        self.adjacentList = self._buildGraph(bidirectional=False)  # 邻接表
        self.adjacentBidirectional = self._buildGraph(bidirectional=True)  # 双向邻接表
        self.__paths_dict = dict(
            [
                (x, self._getPathsTo(x, 0))
                for x in config.VALID_NODE_ID + [config.OUTPUT_NODE_ID]
            ]
        )  # 所有节点到出力节点之间的所有路径
        self.paths = [p for pp in self.__paths_dict.values() for p in pp]
        self.paths = sorted(self.paths)  # 所有非出力节点到出力节点的所有路径，并按照路径电阻升序排序
        valid_adjustable_nodes = [
            (x, min(self._getPathsTo(x, config.OUTPUT_NODE_ID))) for x in config.VALID_NODE_ID
        ]
        valid_adjustable_nodes = sorted(valid_adjustable_nodes, key=lambda x: x[1])
        self.sortedNodeIDs = OrderedDict(
            list(valid_adjustable_nodes)
        )  # 所有非出力节点到节点的最短路径，并升序排序

        '''
        将节点选择的最短路信息，更新到边当中, 这里valid_adjustable_nodes已经按照cost从小到大排序，因此append到arc中的顺序也是从小到大
        保证了arc中的minPathContained中的node顺序是按照cost的从小到大排列的
        '''
        for _, path in valid_adjustable_nodes:
            path.updatePathToArc()

    def _buildGraph(self, bidirectional=False):
        """
        Build the Graph
        :return:
        """
        graph = defaultdict(list)
        for arc in self.arcs:
            graph[arc.nodeFrom].append(arc)
        if bidirectional:
            for arc in self.arcs:
                graph[arc.nodeTo].append(reversed(arc))
        return graph

    def _getPathsTo(self, nodeFrom, nodeTo) -> [Path]:
        """
        Calculate the __paths_dict from Node `nodeFrom` to Node `nodeTo`,
        Two method for path searching, both DFS or BFS are ok.
        0->0:[]
        1->0:[[(1,0),],]
        2->0:[[(2,1),(1,0)],
              [(2,0),]]
        3->0:[[(3,2),(2,1),(1,0)],
              [(3,2),(2,0),]]
        4->0:[[(4,0),]]
        :param nodeFrom:
        :param nodeTo:
        :return: (PATH1,[,PATH2,...]), the list of `PATH`,
                PATH=(Arc1,[,Arc2,...]), represents the path from Node `nodeFrom` to Node `nodeTo`, previous Arc.nodeTo is supposed to be equally to the next Arc.nodeFrom
                Arc=(ID,cost,nodeFrom,nodeTo,capacity,flow)
        """

        def bfs(nodeFrom, nodeTo):
            queue = [(nodeFrom, [])]
            paths = []
            while queue:
                top_node, top_path = queue.pop(0)
                if top_node == nodeTo:
                    paths.append(Path(top_path))
                for arc in self.adjacentList[top_node]:
                    queue.append((arc.nodeTo, top_path + [arc]))
            return paths

        def dfs(nodeFrom, nodeTo):
            def helper(current, path: [Arc]):
                if current == nodeTo:
                    paths.append(Path(path))
                    return
                for arc in self.adjacentList[current]:
                    helper(arc.nodeTo, path + [arc])

            paths = []
            helper(nodeFrom, [])
            return paths

        return dfs(nodeFrom, nodeTo)

    def __repr__(self):
        s = ""
        for x in self.config.VALID_NODE_ID:
            s += f"Node={x},miniest_cost={min(path.cost for path in self.__paths_dict[x])}\n"
            for index, path in enumerate(self.__paths_dict[x]):
                s += f"Path {index + 1}/{len(self.__paths_dict[x])}: {path}\n"
        return s

