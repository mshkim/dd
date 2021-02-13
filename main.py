import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from GrayIsland import onlineShopTask

form_class = uic.loadUiType("ver01.ui")[0]

class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.orderListButton.clicked.connect(self.orderListButton_clicked)
        self.packageNumberButton.clicked.connect(self.packageNumberButton_clicked)
        self.stockManageButton.clicked.connect(self.stockManageButton_clicked)
        self.executeButton.clicked.connect(self.executeButton_clicked)
        self.task = onlineShopTask.OnlineShopTask()

        self.label1 = QLabel()
        self.label2 = QLabel()
        self.label3 = QLabel()
        self.label1.setText('None')
        self.label2.setText('None')
        self.label3.setText('None')
        self.reStockingListTable.setRowCount(5)
        self.reStockingListTable.setColumnCount(3)
        self.setTableWidgetData()

    def setTableWidgetData(self):
        self.reStockingListTable.setItem(0, 0, QTableWidgetItem("(0,0)"))
        self.reStockingListTable.setItem(0, 1, QTableWidgetItem("(0,1)"))
        self.reStockingListTable.setItem(1, 0, QTableWidgetItem("(1,0)"))
        self.reStockingListTable.setItem(1, 1, QTableWidgetItem("(1,1)"))

    def orderListButton_clicked(self):
        fname = QFileDialog.getOpenFileName(self)
        self.label1.setText(fname[0])
        print('Selected file name :', fname[0])

    def packageNumberButton_clicked(self):
        fname = QFileDialog.getOpenFileName(self)
        self.label2.setText(fname[0])
        print('Selected file name :', fname[0])

    def stockManageButton_clicked(self):
        fname = QFileDialog.getOpenFileName(self)
        self.label3.setText(fname[0])
        print('Selected file name :', fname[0])

    def executeButton_clicked(self):
        orderEexl = self.label1.text()
        packageNumExel = self.label2.text()
        stockManageExel = self.label3.text()
        print(orderEexl, packageNumExel, stockManageExel)
        if orderEexl != 'None' and packageNumExel != 'None' and stockManageExel != 'None':
            print('<3개 파일 입력 완료>')
            self.task.NaverPackageDelivery(orderEexl, packageNumExel)
            self.task.ManageItems(orderEexl, stockManageExel)
        elif orderEexl == 'None':
            QMessageBox.about(self, "message", "주문리스트파일 업로드안함^^")
        elif packageNumExel == 'None':
            QMessageBox.about(self, "message", "송장번호파일 업로드안함^^")
        elif stockManageExel == 'None':
            QMessageBox.about(self, "message", "재고관리파일 업로드안함^^")

        #self.task.NeedReStockList()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()