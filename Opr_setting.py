from PyQt5.QtWidgets import *
import Path
from setting import Ui_MainWindow
import sys
import os
class SWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("设置")
        self.pushButton.clicked.connect(self.open0)
        self.pushButton_2.clicked.connect(self.save0)

    def open0(self):
        Path.Hust_path=[]
        exe_path, ok = QFileDialog.getOpenFileName(self, "打开", "D:/")
        if exe_path == '':
            return
        Path.Hust_path = os.path.dirname(exe_path) + '/'
        filename = os.path.basename(exe_path)
        if filename!='HUST_ProS.exe':
            QMessageBox.information(self, "错误", "你打开了不正确的exe文件", QMessageBox.Yes)
        else:
            self.lineEdit.setText(Path.Hust_path + filename)

    def save0(self):
        self.close()
    def Visi(self):
        if not self.isVisible():
            self.show()

if __name__ == "__main__":
    app=QApplication(sys.argv)
    myWin=SWindow()
    myWin.show()
    sys.exit(app.exec())