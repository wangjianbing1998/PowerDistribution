# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'MWin.ui'
#
# Created by: PyQt5 UI code generator 5.13.0
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 18))
        self.menubar.setObjectName("menubar")
        self.menu = QtWidgets.QMenu(self.menubar)
        self.menu.setObjectName("menu")
        self.menu_2 = QtWidgets.QMenu(self.menubar)
        self.menu_2.setMaximumSize(QtCore.QSize(16777215, 16777214))
        self.menu_2.setObjectName("menu_2")
        self.menu_3 = QtWidgets.QMenu(self.menubar)
        self.menu_3.setObjectName("menu_3")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.action1 = QtWidgets.QAction(MainWindow)
        self.action1.setObjectName("action1")
        self.action2 = QtWidgets.QAction(MainWindow)
        self.action2.setObjectName("action2")
        self.action3 = QtWidgets.QAction(MainWindow)
        self.action3.setObjectName("action3")
        self.action4 = QtWidgets.QAction(MainWindow)
        self.action4.setObjectName("action4")
        self.action5 = QtWidgets.QAction(MainWindow)
        self.action5.setObjectName("action5")
        self.action3_2 = QtWidgets.QAction(MainWindow)
        self.action3_2.setObjectName("action3_2")
        self.action5_2 = QtWidgets.QAction(MainWindow)
        self.action5_2.setObjectName("action5_2")
        self.menu.addAction(self.action1)
        self.menu.addAction(self.action3_2)
        self.menu.addAction(self.action2)
        self.menu_2.addAction(self.action3)
        self.menu_2.addAction(self.action4)
        self.menu_3.addAction(self.action5_2)
        self.menubar.addAction(self.menu.menuAction())
        self.menubar.addAction(self.menu_2.menuAction())
        self.menubar.addAction(self.menu_3.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.menu.setTitle(_translate("MainWindow", "开始"))
        self.menu_2.setTitle(_translate("MainWindow", "计算"))
        self.menu_3.setTitle(_translate("MainWindow", "结果"))
        self.action1.setText(_translate("MainWindow", "打开文件"))
        self.action2.setText(_translate("MainWindow", "退出"))
        self.action2.setIconText(_translate("MainWindow", "退出"))
        self.action3.setText(_translate("MainWindow", "初始送电曲线生成"))
        self.action4.setText(_translate("MainWindow", "直流运行模拟"))
        self.action5.setText(_translate("MainWindow", "当地运行模拟"))
        self.action3_2.setText(_translate("MainWindow", "设置"))
        self.action5_2.setText(_translate("MainWindow", "指标分析"))
