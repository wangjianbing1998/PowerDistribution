import numpy as np



def Unit_Commitment(system,station,character,waisong,PV_joint,PV_station,WS):
    inject = []

    # TODO
    UC = np.mat(np.ones((1,12)))#返回各火电站开机矩阵，1*12
    PT = 800*np.mat(np.ones((12,24)))#返回各火电站出力矩阵，12*（24+最后一列火电站ID）
    PT = np.column_stack((PT, [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]))
    inject = 1600*np.mat(np.ones((6,25)))#返回各节点注入功率矩阵，6*（24+最后一列节点ID）
    inject = np.column_stack((inject, [0, 1, 2, 3, 4, 5]))
    abandon = 1600 * np.mat(np.ones((6, 25)))  # 返回各节点弃电，6*（24+最后一列节点ID）
    abandon = np.column_stack((abandon, [0, 1, 2, 3, 4, 5]))
    return (UC,PT,inject,abandon)

