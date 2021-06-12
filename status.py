# encoding=utf-8
from util import checkIsInteger


class Status(object):#外部输入，即新能源出力及负荷（送电曲线），对于固定的新能源和负荷曲线，可视作一个场景的状态Status
    def __init__(self,
                 sceneNum: int,  # 场景ID
                 month: int,  # 月份
                 powerDemand: list,  # 送电曲线 P外送(t),来自中间数据-场景曲线.xlsx-外送曲线-C2~Z13
                 totalCleanProduct: list,  # 上层汇总风光联合出力 PV(t)
                 # 来自中间数据-场景曲线.xlsx-新能源-电站ID为-99的部分
                 ):
        self._sceneNum = sceneNum
        self._month = month
        self.powerDemand = powerDemand
        self.totalCleanProduct = totalCleanProduct

    def __repr__(self):
        return (
            f" 场景ID={self._sceneNum},"
            f" 月份={self._month},"
            f" 电力需求={self.powerDemand},"
            f" 风光联合出力={self.totalCleanProduct},"
            f"\n"
        )

    @property
    def month(self):
        return self._month

    @month.setter#*.setter可读可写
    def month(self, value):
        if 1 <= value <= 12:
            self._month = value
            return
        raise ValueError('月份数值必须在1到12之间')

    @property#只读属性
    def sceneNum(self):
        return self._sceneNum

    @sceneNum.setter#*.setter可读可写
    def sceneNum(self, value):
        self._sceneNum=value

class StatusDict(object):
    """
    statusDict[sceneNum, month] = status
    """
    def __init__(self, statusList: [Status]):
        self._statusList = {}
        for status in statusList:
            self._statusList[status.sceneNum, status.month] = status

    def __getitem__(self, item):
        assert len(item) == 2 and checkIsInteger(item[0]) and type(
            item[1]) == int, f'搜索关键字格式必须为(场景ID,月份), 错误内容：{item}'
        return self._statusList[item]
