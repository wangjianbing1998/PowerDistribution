import numpy as np
def findNum(A, B):#数组A当中哪些不在数组B里
    C = []
    if len(A) == 0:
        return A
    if len(B) == 0:
        return A
    for i in range(len(A)):  # i=0,1,2,3...
        if A[i] not in B:
            C.append(A[i])
    return C

#TODO
def PLine(g,matrix,system,line,station,UCT,PT,inject,):
    DG_capacity=[]#读各电站单机容量
    Nmax = []#各电站机组台数
    percent=[]#技术出力百分比
    for i in range(station.shape[0]):#station.shape输出矩阵的行数和列数，[0]是第一个元素代表行数；[1]是第2个元素代表列数
        if station[i, 5] == 101:#选取标识为101的，即所有火电站
            DG_capacity.append(station[i, 6])
            Nmax.append(station[i, 7])
            percent.append(station[i, 8])
    PTmin=np.array(DG_capacity)*np.array(percent)#最小技术出力
    ZG_capacity=np.array(DG_capacity)*np.array(Nmax)#装机
    KJ_Num=[2,2,1,1,2,1,1,1,1,1,1,1]#我假设的各电站的开机台数
    KJ_capacity=np.array(KJ_Num)*np.array(PTmin)#开机容量
    Reserve=50*np.ones((12,))#旋转备用
    Repair=np.zeros((12,))#检修容量




    UC=KJ_Num#返回各火电站修正后的开机矩阵，1*12
    PT=700*np.mat(np.ones((12,24)))#返回各火电站修正后出力+电站ID+装机+开机容量+最小出力+旋转备用+检修容量
    PT = np.column_stack((PT, [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11],ZG_capacity,KJ_capacity,PTmin,Reserve,Repair))
    yk=300*np.mat(np.ones((6,24)))#返回各节点的电力盈亏+系统ID
    yk = np.column_stack((yk, [0, 1, 2, 3, 4, 5]))
    return (UC,PT,yk)