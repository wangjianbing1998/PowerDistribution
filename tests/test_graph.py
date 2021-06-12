# encoding=utf-8
"""
@Author: jianbingxia 
@Time: 2020/12/3 15:37
@File: test_graph.py
@Description: 



"""
import better_exceptions
from beeprint import pp

from arc import Graph
from solver import Solver

better_exceptions.hook()


class GraphTester:
    def __init__(self):
        solver = Solver()
        self.graph = Graph(solver.arcs)

    def test_get_paths_to(self):
        pp(self.graph._getPathsTo(3, 0))

    def test_get_cost_to0(self):
        pass
        # pp(self.graph.getMinCostTo0(3))


if __name__ == "__main__":
    GraphTester().test_get_paths_to()
    GraphTester().test_get_cost_to0()
