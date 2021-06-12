# encoding=utf-8
"""
@Author: jianbingxia 
@Time: 2020/12/1 16:14
@File: test_solver.py
@Description:
"""
import logging

import matplotlib.pyplot as plt

plt.rcParams["font.sans-serif"] = "SimHei"
plt.rcParams["axes.unicode_minus"] = False

plt.rcParams["font.sans-serif"] = "SimHei"
plt.rcParams["axes.unicode_minus"] = False

import solver


class SolverTester:
    def get_solver(self):
        return solver.Solver()

    def test_calculate_demand(self):
        logging.debug("test calculateDemand ...")
        solver = self.get_solver()
        solver.calculateDemand(sceneNum=0, month=1)
        print(solver.c0)

    def test_sort_arc(self):
        logging.debug("test sortPath ...")
        solver = self.get_solver()
        solver.sortPath()
        print(solver.arcs)

    def test_assignOpenSchema(self):
        logging.debug("test assignOpenSchema ...")
        solver = self.get_solver()
        solver.assignOpenSchema()
        print(solver)


if __name__ == "__main__":
    # SolverTester().test_calculate_demand()
    # SolverTester().test_sort_arc()
    SolverTester().test_assignOpenSchema()
