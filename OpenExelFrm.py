# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'D:\Users\Sniwt-home\Desktop\работа с excel и БД в pyton\OpenExelFrm.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(647, 591)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.OpenBtn = QtWidgets.QPushButton(self.centralwidget)
        self.OpenBtn.setObjectName("OpenBtn")
        self.verticalLayout.addWidget(self.OpenBtn, 0, QtCore.Qt.AlignLeft)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setContentsMargins(-1, 0, -1, -1)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.SaveBtn = QtWidgets.QPushButton(self.centralwidget)
        self.SaveBtn.setObjectName("SaveBtn")
        self.horizontalLayout.addWidget(self.SaveBtn)
        self.SaveAllBtn = QtWidgets.QPushButton(self.centralwidget)
        self.SaveAllBtn.setObjectName("SaveAllBtn")
        self.horizontalLayout.addWidget(self.SaveAllBtn)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.SelFileName = QtWidgets.QLabel(self.centralwidget)
        self.SelFileName.setText("")
        self.SelFileName.setObjectName("SelFileName")
        self.verticalLayout.addWidget(self.SelFileName)
        self.tableWd = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWd.setRowCount(0)
        self.tableWd.setColumnCount(0)
        self.tableWd.setObjectName("tableWd")
        self.verticalLayout.addWidget(self.tableWd)
        self.ChoosSheetL = QtWidgets.QHBoxLayout()
        self.ChoosSheetL.setContentsMargins(150, -1, 150, 10)
        self.ChoosSheetL.setObjectName("ChoosSheetL")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName("label")
        self.ChoosSheetL.addWidget(self.label)
        self.sheetSpin = QtWidgets.QSpinBox(self.centralwidget)
        self.sheetSpin.setEnabled(False)
        self.sheetSpin.setSpecialValueText("")
        self.sheetSpin.setMinimum(1)
        self.sheetSpin.setProperty("value", 1)
        self.sheetSpin.setObjectName("sheetSpin")
        self.ChoosSheetL.addWidget(self.sheetSpin)
        self.SheetLb = QtWidgets.QLabel(self.centralwidget)
        self.SheetLb.setObjectName("SheetLb")
        self.ChoosSheetL.addWidget(self.SheetLb)
        self.verticalLayout.addLayout(self.ChoosSheetL)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setContentsMargins(-1, 0, -1, 10)
        self.horizontalLayout_2.setSpacing(10)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.LoadDBBtn = QtWidgets.QPushButton(self.centralwidget)
        self.LoadDBBtn.setObjectName("LoadDBBtn")
        self.horizontalLayout_2.addWidget(self.LoadDBBtn)
        self.DBbtn = QtWidgets.QPushButton(self.centralwidget)
        self.DBbtn.setObjectName("DBbtn")
        self.horizontalLayout_2.addWidget(self.DBbtn)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.EditeDBBt = QtWidgets.QPushButton(self.centralwidget)
        self.EditeDBBt.setObjectName("EditeDBBt")
        self.verticalLayout.addWidget(self.EditeDBBt)
        self.AllPB = QtWidgets.QProgressBar(self.centralwidget)
        self.AllPB.setEnabled(True)
        self.AllPB.setMinimum(0)
        self.AllPB.setProperty("value", 0)
        self.AllPB.setAlignment(QtCore.Qt.AlignBottom|QtCore.Qt.AlignHCenter)
        self.AllPB.setTextVisible(True)
        self.AllPB.setInvertedAppearance(False)
        self.AllPB.setTextDirection(QtWidgets.QProgressBar.TopToBottom)
        self.AllPB.setObjectName("AllPB")
        self.verticalLayout.addWidget(self.AllPB)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 647, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Работа с файлами и таблицами"))
        self.OpenBtn.setText(_translate("MainWindow", "Открыть файл Excel"))
        self.SaveBtn.setText(_translate("MainWindow", "Сохранить текущую страницу в файл"))
        self.SaveAllBtn.setText(_translate("MainWindow", "Сохранить всю книгу в файл"))
        self.label.setText(_translate("MainWindow", "Выберете страницу"))
        self.SheetLb.setText(_translate("MainWindow", "Not select"))
        self.LoadDBBtn.setText(_translate("MainWindow", "Загрузить из базы данных"))
        self.DBbtn.setText(_translate("MainWindow", "Сформировать базу данных"))
        self.EditeDBBt.setText(_translate("MainWindow", "Работа с базой данных"))
        self.AllPB.setFormat(_translate("MainWindow", "%v/100"))