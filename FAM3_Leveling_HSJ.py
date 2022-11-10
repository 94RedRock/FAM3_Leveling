import datetime
import glob
import logging
from logging.handlers import RotatingFileHandler
from operator import index
import os
import re
import math
from termios import B0
import time
import numpy as np
import openpyxl as xl
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtCore import Qt, QCoreApplication
from PyQt5.QtGui import QDoubleValidator, QStandardItemModel, QIcon, QStandardItem, QIntValidator, QFont
from PyQt5.QtWidgets import QMainWindow, QMessageBox, QProgressBar, QPlainTextEdit, QWidget, QGridLayout, QGroupBox, QLineEdit, QSizePolicy, QToolButton, QLabel, QFrame, QListView, QMenuBar, QStatusBar, QPushButton, QApplication, QCalendarWidget, QVBoxLayout, QFileDialog, QCheckBox
from PyQt5.QtCore import pyqtSlot, pyqtSignal, QObject, QThread, QRect, QSize, QDate
import pandas as pd
import cx_Oracle
from collections import OrderedDict
from collections import defaultdict 

class CustomFormatter(logging.Formatter):
    FORMATS = {
        logging.ERROR:   ('[%(asctime)s] %(levelname)s:%(message)s','white'),
        logging.DEBUG:   ('[%(asctime)s] %(levelname)s:%(message)s','white'),
        logging.INFO:    ('[%(asctime)s] %(levelname)s:%(message)s','white'),
        logging.WARNING: ('[%(asctime)s] %(levelname)s:%(message)s', 'yellow')
    }

    def format( self, record ):
        last_fmt = self._style._fmt
        opt = CustomFormatter.FORMATS.get(record.levelno)
        if opt:
            fmt, color = opt
            self._style._fmt = "<font color=\"{}\">{}</font>".format(QtGui.QColor(color).name(),fmt)
        res = logging.Formatter.format( self, record )
        self._style._fmt = last_fmt
        return res

class QTextEditLogger(logging.Handler):
    def __init__(self, parent=None):
        super().__init__()
        self.widget = QPlainTextEdit(parent)
        self.widget.setReadOnly(True)    
        self.widget.setGeometry(QRect(10, 260, 661, 161))
        self.widget.setStyleSheet('background-color: rgb(53, 53, 53);\ncolor: rgb(255, 255, 255);')
        self.widget.setObjectName('logBrowser')
        font = QFont()
        font.setFamily('Nanum Gothic')
        font.setBold(False)
        font.setPointSize(9)
        self.widget.setFont(font)

    def emit(self, record):
        msg = self.format(record)
        self.widget.appendHtml(msg) 
        # move scrollbar
        scrollbar = self.widget.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

class CalendarWindow(QWidget):
    submitClicked = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        cal = QCalendarWidget(self)
        cal.setGridVisible(True)
        cal.clicked[QDate].connect(self.showDate)
        self.lb = QLabel(self)
        date = cal.selectedDate()
        self.lb.setText(date.toString("yyyy-MM-dd"))
        vbox = QVBoxLayout()
        vbox.addWidget(cal)
        vbox.addWidget(self.lb)
        self.submitBtn = QToolButton(self)
        sizePolicy = QSizePolicy(QSizePolicy.Ignored, 
                                    QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        self.submitBtn.setSizePolicy(sizePolicy)
        self.submitBtn.setMinimumSize(QSize(0, 35))
        self.submitBtn.setStyleSheet('background-color: rgb(63, 63, 63);\ncolor: rgb(255, 255, 255);')
        self.submitBtn.setObjectName('submitBtn')
        self.submitBtn.setText('착공지정일 결정')
        self.submitBtn.clicked.connect(self.confirm)
        vbox.addWidget(self.submitBtn)

        self.setLayout(vbox)
        self.setWindowTitle('캘린더')
        self.setGeometry(500,500,500,400)
        self.show()

    def showDate(self, date):
        self.lb.setText(date.toString("yyyy-MM-dd"))

    @pyqtSlot()
    def confirm(self):
        self.submitClicked.emit(self.lb.text())
        self.close()

class UISubWindow(QMainWindow):
    submitClicked = pyqtSignal(list)
    status = ''

    def __init__(self):
        super().__init__()
        self.setupUi()

    def setupUi(self):
        self.setObjectName('SubWindow')
        self.resize(600, 600)
        self.setStyleSheet('background-color: rgb(252, 252, 252);')
        self.centralwidget = QWidget(self)
        self.centralwidget.setObjectName('centralwidget')
        self.gridLayout2 = QGridLayout(self.centralwidget)
        self.gridLayout2.setObjectName('gridLayout2')
        self.gridLayout = QGridLayout()
        self.gridLayout.setObjectName('gridLayout')
        self.groupBox = QGroupBox(self.centralwidget)
        self.groupBox.setTitle('')
        self.groupBox.setObjectName('groupBox')
        self.gridLayout4 = QGridLayout(self.groupBox)
        self.gridLayout4.setObjectName('gridLayout4')
        self.gridLayout3 = QGridLayout()
        self.gridLayout3.setObjectName('gridLayout3')
        self.linkageInput = QLineEdit(self.groupBox)
        self.linkageInput.setMinimumSize(QSize(0, 25))
        self.linkageInput.setObjectName('linkageInput')
        self.linkageInput.setValidator(QDoubleValidator(self))
        self.gridLayout3.addWidget(self.linkageInput, 0, 1, 1, 3)
        self.linkageInputBtn = QPushButton(self.groupBox)
        self.linkageInputBtn.setMinimumSize(QSize(0,25))
        self.gridLayout3.addWidget(self.linkageInputBtn, 0, 4, 1, 2)
        self.linkageAddExcelBtn = QPushButton(self.groupBox)
        self.linkageAddExcelBtn.setMinimumSize(QSize(0,25))
        self.gridLayout3.addWidget(self.linkageAddExcelBtn, 0, 6, 1, 2)
        self.mscodeInput = QLineEdit(self.groupBox)
        self.mscodeInput.setMinimumSize(QSize(0, 25))
        self.mscodeInput.setObjectName('mscodeInput')
        self.mscodeInputBtn = QPushButton(self.groupBox)
        self.mscodeInputBtn.setMinimumSize(QSize(0,25))
        self.gridLayout3.addWidget(self.mscodeInput, 1, 1, 1, 3)
        self.gridLayout3.addWidget(self.mscodeInputBtn, 1, 4, 1, 2)
        self.mscodeAddExcelBtn = QPushButton(self.groupBox)
        self.mscodeAddExcelBtn.setMinimumSize(QSize(0,25))
        self.gridLayout3.addWidget(self.mscodeAddExcelBtn, 1, 6, 1, 2)
        sizePolicy = QSizePolicy(QSizePolicy.Ignored,
                                    QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        self.submitBtn = QToolButton(self.groupBox)
        sizePolicy.setHeightForWidth(self.submitBtn.sizePolicy().hasHeightForWidth())
        self.submitBtn.setSizePolicy(sizePolicy)
        self.submitBtn.setMinimumSize(QSize(100, 35))
        self.submitBtn.setStyleSheet('background-color: rgb(63, 63, 63);\ncolor: rgb(255, 255, 255);')
        self.submitBtn.setObjectName('submitBtn')
        self.gridLayout3.addWidget(self.submitBtn, 3, 5, 1, 2)
        
        self.label = QLabel(self.groupBox)
        self.label.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label.setObjectName('label')
        self.gridLayout3.addWidget(self.label, 0, 0, 1, 1)
        self.label2 = QLabel(self.groupBox)
        self.label2.setAlignment(Qt.AlignRight | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label2.setObjectName('label2')
        self.gridLayout3.addWidget(self.label2, 1, 0, 1, 1)
        self.line = QFrame(self.groupBox)
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)
        self.line.setObjectName('line')
        self.gridLayout3.addWidget(self.line, 2, 0, 1, 10)
        self.gridLayout4.addLayout(self.gridLayout3, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.groupBox, 0, 0, 1, 1)
        self.groupBox2 = QGroupBox(self.centralwidget)
        self.groupBox2.setTitle('')
        self.groupBox2.setObjectName('groupBox2')
        self.gridLayout6 = QGridLayout(self.groupBox2)
        self.gridLayout6.setObjectName('gridLayout6')
        self.gridLayout5 = QGridLayout()
        self.gridLayout5.setObjectName('gridLayout5')
        listViewModelLinkage = QStandardItemModel()
        self.listViewLinkage = QListView(self.groupBox2)
        self.listViewLinkage.setModel(listViewModelLinkage)
        self.gridLayout5.addWidget(self.listViewLinkage, 1, 0, 1, 1)
        self.label3 = QLabel(self.groupBox2)
        self.label3.setAlignment(Qt.AlignLeft | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label3.setObjectName('label3')
        self.gridLayout5.addWidget(self.label3, 0, 0, 1, 1)

        self.vline = QFrame(self.groupBox2)
        self.vline.setFrameShape(QFrame.VLine)
        self.vline.setFrameShadow(QFrame.Sunken)
        self.vline.setObjectName('vline')
        self.gridLayout5.addWidget(self.vline, 1, 1, 1, 1)
        listViewModelmscode = QStandardItemModel()
        self.listViewmscode = QListView(self.groupBox2)
        self.listViewmscode.setModel(listViewModelmscode)
        self.gridLayout5.addWidget(self.listViewmscode, 1, 2, 1, 1)
        self.label4 = QLabel(self.groupBox2)
        self.label4.setAlignment(Qt.AlignLeft | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label4.setObjectName('label4')
        self.gridLayout5.addWidget(self.label4, 0, 2, 1, 1)
        self.label5 = QLabel(self.groupBox2)
        self.label5.setAlignment(Qt.AlignLeft | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label5.setObjectName('label5')       
        self.gridLayout5.addWidget(self.label5, 0, 3, 1, 1) 
        self.linkageDelBtn = QPushButton(self.groupBox2)
        self.linkageDelBtn.setMinimumSize(QSize(0,25))
        self.gridLayout5.addWidget(self.linkageDelBtn, 2, 0, 1, 1)
        self.mscodeDelBtn = QPushButton(self.groupBox2)
        self.mscodeDelBtn.setMinimumSize(QSize(0,25))
        self.gridLayout5.addWidget(self.mscodeDelBtn, 2, 2, 1, 1)
        self.gridLayout6.addLayout(self.gridLayout5, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.groupBox2, 1, 0, 1, 1)
        self.gridLayout2.addLayout(self.gridLayout, 0, 0, 1, 1)
        self.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(self)
        self.menubar.setGeometry(QRect(0, 0, 653, 21))
        self.menubar.setObjectName('menubar')
        self.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(self)
        self.statusbar.setObjectName('statusbar')
        self.setStatusBar(self.statusbar)
        self.retranslateUi(self)
        self.mscodeInput.returnPressed.connect(self.addmscode)
        self.linkageInput.returnPressed.connect(self.addLinkage)
        self.linkageInputBtn.clicked.connect(self.addLinkage)
        self.mscodeInputBtn.clicked.connect(self.addmscode)
        self.linkageDelBtn.clicked.connect(self.delLinkage)
        self.mscodeDelBtn.clicked.connect(self.delmscode)
        self.submitBtn.clicked.connect(self.confirm)
        self.linkageAddExcelBtn.clicked.connect(self.addLinkageExcel)
        self.mscodeAddExcelBtn.clicked.connect(self.addmscodeExcel)
        self.retranslateUi(self)
        self.show()
    
    def retranslateUi(self, MainWindow):
        _translate = QCoreApplication.translate
        MainWindow.setWindowTitle(_translate('SubWindow', '긴급/홀딩오더 입력'))
        MainWindow.setWindowIcon(QIcon('.\\Logo\\logo.png'))
        self.label.setText(_translate('SubWindow', 'Linkage No 입력 :'))
        self.linkageInputBtn.setText(_translate('SubWindow', '추가'))
        self.label2.setText(_translate('SubWindow', 'MS-CODE 입력 :'))
        self.mscodeInputBtn.setText(_translate('SubWindow', '추가'))
        self.submitBtn.setText(_translate('SubWindow','추가 완료'))
        self.label3.setText(_translate('SubWindow', 'Linkage No List'))
        self.label4.setText(_translate('SubWindow', 'MS-Code List'))
        self.linkageDelBtn.setText(_translate('SubWindow', '삭제'))
        self.mscodeDelBtn.setText(_translate('SubWindow', '삭제'))
        self.linkageAddExcelBtn.setText(_translate('SubWindow', '엑셀 입력'))
        self.mscodeAddExcelBtn.setText(_translate('SubWindow', '엑셀 입력'))

    @pyqtSlot()
    def addLinkage(self):
        linkageNo = self.linkageInput.text()
        if len(linkageNo) == 16:
            if linkageNo.isdigit():
                model = self.listViewLinkage.model()
                linkageItem = QStandardItem()
                linkageItemModel = QStandardItemModel()
                dupFlag = False
                for i in range(model.rowCount()):
                    index = model.index(i,0)
                    item = model.data(index)
                    if item == linkageNo:
                        dupFlag = True
                    linkageItem = QStandardItem(item)
                    linkageItemModel.appendRow(linkageItem)
                if not dupFlag:
                    linkageItem = QStandardItem(linkageNo)
                    linkageItemModel.appendRow(linkageItem)
                    self.listViewLinkage.setModel(linkageItemModel)
                else:
                    QMessageBox.information(self, 'Error', '중복된 데이터가 있습니다.')
            else:
                QMessageBox.information(self, 'Error', '숫자만 입력해주세요.')
        elif len(linkageNo) == 0: 
            QMessageBox.information(self, 'Error', 'Linkage Number 데이터가 입력되지 않았습니다.')
        else:
            QMessageBox.information(self, 'Error', '16자리의 Linkage Number를 입력해주세요.')
    
    @pyqtSlot()
    def delLinkage(self):
        model = self.listViewLinkage.model()
        linkageItem = QStandardItem()
        linkageItemModel = QStandardItemModel()
        for index in self.listViewLinkage.selectedIndexes():
            selected_item = self.listViewLinkage.model().data(index)
            for i in range(model.rowCount()):
                index = model.index(i,0)
                item = model.data(index)
                linkageItem = QStandardItem(item)
                if selected_item != item:
                    linkageItemModel.appendRow(linkageItem)
            self.listViewLinkage.setModel(linkageItemModel)

    @pyqtSlot()
    def addmscode(self):
        mscode = self.mscodeInput.text()
        if len(mscode) > 0:
            model = self.listViewmscode.model()
            mscodeItem = QStandardItem()
            mscodeItemModel = QStandardItemModel()
            dupFlag = False
            for i in range(model.rowCount()):
                index = model.index(i,0)
                item = model.data(index)
                if item == mscode:
                    dupFlag = True
                mscodeItem = QStandardItem(item)
                mscodeItemModel.appendRow(mscodeItem)
            if not dupFlag:
                mscodeItem = QStandardItem(mscode)
                mscodeItemModel.appendRow(mscodeItem)
                self.listViewmscode.setModel(mscodeItemModel)
            else:
                QMessageBox.information(self, 'Error', '중복된 데이터가 있습니다.')
        else: 
            QMessageBox.information(self, 'Error', 'MS-CODE 데이터가 입력되지 않았습니다.')

    @pyqtSlot()
    def delmscode(self):
        model = self.listViewmscode.model()
        mscodeItem = QStandardItem()
        mscodeItemModel = QStandardItemModel()
        for index in self.listViewmscode.selectedIndexes():
            selected_item = self.listViewmscode.model().data(index)
            for i in range(model.rowCount()):
                index = model.index(i,0)
                item = model.data(index)
                mscodeItem = QStandardItem(item)
                if selected_item != item:
                    mscodeItemModel.appendRow(mscodeItem)
            self.listViewmscode.setModel(mscodeItemModel)
    @pyqtSlot()
    def addLinkageExcel(self):
        try:
            fileName = QFileDialog.getOpenFileName(self, 'Open File', './', 'Excel Files (*.xlsx)')[0]
            if fileName != "":
                df = pd.read_excel(fileName)
                for i in df.index:
                    linkageNo = str(df[df.columns[0]][i])
                    if len(linkageNo) == 16:
                        if linkageNo.isdigit():
                            model = self.listViewLinkage.model()
                            linkageItem = QStandardItem()
                            linkageItemModel = QStandardItemModel()
                            dupFlag = False
                            for i in range(model.rowCount()):
                                index = model.index(i,0)
                                item = model.data(index)
                                if item == linkageNo:
                                    dupFlag = True
                                linkageItem = QStandardItem(item)
                                linkageItemModel.appendRow(linkageItem)
                            if not dupFlag:
                                linkageItem = QStandardItem(linkageNo)
                                linkageItemModel.appendRow(linkageItem)
                                self.listViewLinkage.setModel(linkageItemModel)
                            else:
                                QMessageBox.information(self, 'Error', '중복된 데이터가 있습니다.')
                        else:
                            QMessageBox.information(self, 'Error', '숫자만 입력해주세요.')
                    elif len(linkageNo) == 0: 
                        QMessageBox.information(self, 'Error', 'Linkage Number 데이터가 입력되지 않았습니다.')
                    else:
                        QMessageBox.information(self, 'Error', '16자리의 Linkage Number를 입력해주세요.')
        except Exception as e:
            QMessageBox.information(self, 'Error', '에러발생 : ' + e)
    @pyqtSlot()
    def addmscodeExcel(self):
        try:
            fileName = QFileDialog.getOpenFileName(self, 'Open File', './', 'Excel Files (*.xlsx)')[0]
            if fileName != "":
                df = pd.read_excel(fileName)
                for i in df.index:
                    mscode = str(df[df.columns[0]][i])
                    if len(mscode) > 0:
                        model = self.listViewmscode.model()
                        mscodeItem = QStandardItem()
                        mscodeItemModel = QStandardItemModel()
                        dupFlag = False
                        for i in range(model.rowCount()):
                            index = model.index(i,0)
                            item = model.data(index)
                            if item == mscode:
                                dupFlag = True
                            mscodeItem = QStandardItem(item)
                            mscodeItemModel.appendRow(mscodeItem)
                        if not dupFlag:
                            mscodeItem = QStandardItem(mscode)
                            mscodeItemModel.appendRow(mscodeItem)
                            self.listViewmscode.setModel(mscodeItemModel)
                        else:
                            QMessageBox.information(self, 'Error', '중복된 데이터가 있습니다.')
                    else: 
                        QMessageBox.information(self, 'Error', 'MS-CODE 데이터가 입력되지 않았습니다.')
        except Exception as e:
            QMessageBox.information(self, 'Error', '에러발생 : ' + e)
    @pyqtSlot()
    def confirm(self):
        self.submitClicked.emit([self.listViewLinkage.model(), self.listViewmscode.model()])
        self.close()

class Ui_MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi()
        
    def setupUi(self):
        logger = logging.getLogger(__name__)
        rfh = RotatingFileHandler(filename='./Log.log', 
                                    mode='a',
                                    maxBytes=5*1024*1024,
                                    backupCount=2,
                                    encoding=None,
                                    delay=0
                                    )
        logging.basicConfig(level=logging.DEBUG, 
                            format = '%(asctime)s:%(levelname)s:%(message)s', 
                            datefmt = '%m/%d/%Y %H:%M:%S',
                            handlers=[rfh])
        self.setObjectName('MainWindow')
        self.resize(900, 1000)
        self.setStyleSheet('background-color: rgb(252, 252, 252);')
        self.centralwidget = QWidget(self)
        self.centralwidget.setObjectName('centralwidget')
        self.gridLayout2 = QGridLayout(self.centralwidget)
        self.gridLayout2.setObjectName('gridLayout2')
        self.gridLayout = QGridLayout()
        self.gridLayout.setObjectName('gridLayout')
        self.groupBox = QGroupBox(self.centralwidget)
        self.groupBox.setTitle('')
        self.groupBox.setObjectName('groupBox')
        self.gridLayout4 = QGridLayout(self.groupBox)
        self.gridLayout4.setObjectName('gridLayout4')
        self.gridLayout3 = QGridLayout()
        self.gridLayout3.setObjectName('gridLayout3')
        self.mainOrderinput = QLineEdit(self.groupBox)
        self.mainOrderinput.setMinimumSize(QSize(0, 25))
        self.mainOrderinput.setObjectName('mainOrderinput')
        self.mainOrderinput.setValidator(QIntValidator(self))
        self.gridLayout3.addWidget(self.mainOrderinput, 0, 1, 1, 1)
        self.spOrderinput = QLineEdit(self.groupBox)
        self.spOrderinput.setMinimumSize(QSize(0, 25))
        self.spOrderinput.setObjectName('spOrderinput')
        self.spOrderinput.setValidator(QIntValidator(self))
        self.gridLayout3.addWidget(self.spOrderinput, 1, 1, 1, 1)
        self.powerOrderinput = QLineEdit(self.groupBox)
        self.powerOrderinput.setMinimumSize(QSize(0, 25))
        self.powerOrderinput.setObjectName('powerOrderinput')
        self.powerOrderinput.setValidator(QIntValidator(self))
        self.gridLayout3.addWidget(self.powerOrderinput, 2, 1, 1, 1)
        self.dateBtn = QToolButton(self.groupBox)
        self.dateBtn.setMinimumSize(QSize(0,25))
        self.dateBtn.setObjectName('dateBtn')
        self.gridLayout3.addWidget(self.dateBtn, 3, 1, 1, 1)
        self.emgFileInputBtn = QPushButton(self.groupBox)
        self.emgFileInputBtn.setMinimumSize(QSize(0,25))
        self.gridLayout3.addWidget(self.emgFileInputBtn, 4, 1, 1, 1)
        self.holdFileInputBtn = QPushButton(self.groupBox)
        self.holdFileInputBtn.setMinimumSize(QSize(0,25))
        self.gridLayout3.addWidget(self.holdFileInputBtn, 7, 1, 1, 1)
        self.label4 = QLabel(self.groupBox)
        self.label4.setAlignment(Qt.AlignRight | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label4.setObjectName('label4')
        self.gridLayout3.addWidget(self.label4, 5, 1, 1, 1)
        self.label5 = QLabel(self.groupBox)
        self.label5.setAlignment(Qt.AlignRight | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label5.setObjectName('label5')
        self.gridLayout3.addWidget(self.label5, 5, 2, 1, 1)
        self.label6 = QLabel(self.groupBox)
        self.label6.setAlignment(Qt.AlignRight | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label6.setObjectName('label6')
        self.gridLayout3.addWidget(self.label6, 8, 1, 1, 1)
        self.label7 = QLabel(self.groupBox)
        self.label7.setAlignment(Qt.AlignRight | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label7.setObjectName('label7')
        self.gridLayout3.addWidget(self.label7, 8, 2, 1, 1)
        listViewModelEmgLinkage = QStandardItemModel()
        self.listViewEmgLinkage = QListView(self.groupBox)
        self.listViewEmgLinkage.setModel(listViewModelEmgLinkage)
        self.gridLayout3.addWidget(self.listViewEmgLinkage, 6, 1, 1, 1)
        listViewModelEmgmscode = QStandardItemModel()
        self.listViewEmgmscode = QListView(self.groupBox)
        self.listViewEmgmscode.setModel(listViewModelEmgmscode)
        self.gridLayout3.addWidget(self.listViewEmgmscode, 6, 2, 1, 1)
        listViewModelHoldLinkage = QStandardItemModel()
        self.listViewHoldLinkage = QListView(self.groupBox)
        self.listViewHoldLinkage.setModel(listViewModelHoldLinkage)
        self.gridLayout3.addWidget(self.listViewHoldLinkage, 9, 1, 1, 1)
        listViewModelHoldmscode = QStandardItemModel()
        self.listViewHoldmscode = QListView(self.groupBox)
        self.listViewHoldmscode.setModel(listViewModelHoldmscode)
        self.gridLayout3.addWidget(self.listViewHoldmscode, 9, 2, 1, 1)
        self.labelBlank = QLabel(self.groupBox)
        self.labelBlank.setObjectName('labelBlank')
        self.gridLayout3.addWidget(self.labelBlank, 2, 4, 1, 1)
        self.progressbar = QProgressBar(self.groupBox)
        self.progressbar.setObjectName('progressbar')
        self.gridLayout3.addWidget(self.progressbar, 10, 1, 1, 2)
        self.runBtn = QToolButton(self.groupBox)
        sizePolicy = QSizePolicy(QSizePolicy.Ignored, 
                                    QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.runBtn.sizePolicy().hasHeightForWidth())
        self.runBtn.setSizePolicy(sizePolicy)
        self.runBtn.setMinimumSize(QSize(0, 35))
        self.runBtn.setStyleSheet('background-color: rgb(63, 63, 63);\ncolor: rgb(255, 255, 255);')
        self.runBtn.setObjectName('runBtn')
        self.gridLayout3.addWidget(self.runBtn, 10, 3, 1, 2)
        self.label = QLabel(self.groupBox)
        self.label.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label.setObjectName('label')
        self.gridLayout3.addWidget(self.label, 0, 0, 1, 1)
        self.label9 = QLabel(self.groupBox)
        self.label9.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label9.setObjectName('label9')
        self.gridLayout3.addWidget(self.label9, 1, 0, 1, 1)
        self.label10 = QLabel(self.groupBox)
        self.label10.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label10.setObjectName('label10')
        self.gridLayout3.addWidget(self.label10, 2, 0, 1, 1)
        self.label8 = QLabel(self.groupBox)
        self.label8.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label8.setObjectName('label8')
        self.gridLayout3.addWidget(self.label8, 3, 0, 1, 1) 
        self.labelDate = QLabel(self.groupBox)
        self.labelDate.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.labelDate.setObjectName('labelDate')
        self.gridLayout3.addWidget(self.labelDate, 3, 2, 1, 1) 
        self.label2 = QLabel(self.groupBox)
        self.label2.setAlignment(Qt.AlignRight | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label2.setObjectName('label2')
        self.gridLayout3.addWidget(self.label2, 4, 0, 1, 1)
        self.label3 = QLabel(self.groupBox)
        self.label3.setAlignment(Qt.AlignRight | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label3.setObjectName('label3')
        self.gridLayout3.addWidget(self.label3, 7, 0, 1, 1)
        self.line = QFrame(self.groupBox)
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)
        self.line.setObjectName('line')
        self.gridLayout3.addWidget(self.line, 11, 0, 1, 10)
        self.gridLayout4.addLayout(self.gridLayout3, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.groupBox, 0, 0, 1, 1)
        self.groupBox2 = QGroupBox(self.centralwidget)
        self.groupBox2.setTitle('')
        self.groupBox2.setObjectName('groupBox2')
        self.gridLayout6 = QGridLayout(self.groupBox2)
        self.gridLayout6.setObjectName('gridLayout6')
        self.gridLayout5 = QGridLayout()
        self.gridLayout5.setObjectName('gridLayout5')
        self.logBrowser = QTextEditLogger(self.groupBox2)
        # self.logBrowser.setFormatter(
        #                             logging.Formatter('[%(asctime)s] %(levelname)s:%(message)s', 
        #                                                 datefmt='%Y-%m-%d %H:%M:%S')
        #                             )
        self.logBrowser.setFormatter(CustomFormatter())
        logging.getLogger().addHandler(self.logBrowser)
        logging.getLogger().setLevel(logging.INFO)
        self.gridLayout5.addWidget(self.logBrowser.widget, 0, 0, 1, 1)
        self.gridLayout6.addLayout(self.gridLayout5, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.groupBox2, 1, 0, 1, 1)
        self.gridLayout2.addLayout(self.gridLayout, 0, 0, 1, 1)
        self.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(self)
        self.menubar.setGeometry(QRect(0, 0, 653, 21))
        self.menubar.setObjectName('menubar')
        self.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(self)
        self.statusbar.setObjectName('statusbar')
        self.setStatusBar(self.statusbar)
        self.retranslateUi(self)
        self.dateBtn.clicked.connect(self.selectStartDate)
        self.emgFileInputBtn.clicked.connect(self.emgWindow)
        self.holdFileInputBtn.clicked.connect(self.holdWindow)
        self.runBtn.clicked.connect(self.startLeveling)

        #디버그용 플래그
        self.isDebug = True
        if self.isDebug:
            self.debugDate = QLineEdit(self.groupBox)
            self.debugDate.setObjectName('debugDate')
            self.gridLayout3.addWidget(self.debugDate, 10, 0, 1, 1)
            self.debugDate.setPlaceholderText('디버그용 날짜입력')
        self.show()

    def retranslateUi(self, MainWindow):
        _translate = QCoreApplication.translate
        MainWindow.setWindowTitle(_translate('MainWindow', 'FA-M3 착공 평준화 자동화 프로그램 Rev0.00'))
        MainWindow.setWindowIcon(QIcon('.\\Logo\\logo.png'))
        self.label.setText(_translate('MainWindow', '메인 생산대수:'))
        self.label9.setText(_translate('MainWindow', '특수 생산대수:'))
        self.label10.setText(_translate('MainWindow', '전원 생산대수:'))
        self.runBtn.setText(_translate('MainWindow', '실행'))
        self.label2.setText(_translate('MainWindow', '긴급오더 입력 :'))
        self.label3.setText(_translate('MainWindow', '홀딩오더 입력 :'))
        self.label4.setText(_translate('MainWindow', 'Linkage No List'))
        self.label5.setText(_translate('MainWindow', 'MSCode List'))
        self.label6.setText(_translate('MainWindow', 'Linkage No List'))
        self.label7.setText(_translate('MainWindow', 'MSCode List'))
        self.label8.setText(_translate('MainWndow', '착공지정일 입력 :'))
        self.labelDate.setText(_translate('MainWndow', '미선택'))
        self.dateBtn.setText(_translate('MainWindow', ' 착공지정일 선택 '))
        self.emgFileInputBtn.setText(_translate('MainWindow', '리스트 입력'))
        self.holdFileInputBtn.setText(_translate('MainWindow', '리스트 입력'))
        self.labelBlank.setText(_translate('MainWindow', '            '))

        # try:
        #     self.df_productTime = self.loadProductTimeDb()
        #     # self.df_productTime.to_excel(r'.\result.xlsx')
        # except Exception as e:
        #     logging.error('검사시간DB 불러오기에 실패했습니다. 관리자에게 문의해주세요.')
        #     logging.exception(e, exc_info=True)      
        # try:
        #     self.df_smt = self.loadSmtDb
        # except Exception as e:
        #     logging.error('SMT Assy 재고량 DB 불러오기에 실패했습니다. 관리자에게 문의해주세요.')
        #     logging.exception(e, exc_info=True)   

        logging.info('프로그램이 정상 기동했습니다')

    # #생산시간 DB로부터 불러오기
    # def loadProductTimeDb(self):
    #     location = r'.\\instantclient_21_6'
    #     os.environ["PATH"] = location + ";" + os.environ["PATH"]
    #     dsn = cx_Oracle.makedsn("ymzn-bdv19az029-rds.cgbtxsdj6fjy.ap-northeast-1.rds.amazonaws.com", 1521, "tprod")
    #     db = cx_Oracle.connect("TEST_SCM","test_scm", dsn)

    #     cursor= db.cursor()
    #     cursor.execute("SELECT MODEL, COMPONENT_SET, MAEDZUKE, MAUNT, LEAD_CUTTING, VISUAL_EXAMINATION, PICKUP, ASSAMBLY, M_FUNCTION_CHECK, A_FUNCTION_CHECK, PERSON_EXAMINE, INSPECTION_EQUIPMENT FROM FAM3_PRODUCT_TIME_TB")
    #     out_data = cursor.fetchall()
    #     df_productTime = pd.DataFrame(out_data)
    #     df_productTime.columns = ["MODEL", "COMPONENT_SET", "MAEDZUKE", "MAUNT", "LEAD_CUTTING", "VISUAL_EXAMINATION", "PICKUP", "ASSAMBLY", "M_FUNCTION_CHECK", "A_FUNCTION_CHECK", "PERSON_EXAMINE", "INSPECTION_EQUIPMENT"]
    #     return df_productTime

    # #SMT Assy 재고 DB로부터 불러오기
    # def loadSmtDb(self):
    #     location = r'.\\instantclient_21_6'
    #     os.environ["PATH"] = location + ";" + os.environ["PATH"]
    #     dsn = cx_Oracle.makedsn("10.36.15.42", 1521, "NEURON")
    #     db = cx_Oracle.connect("ymi_user","ymi123!", dsn)

    #     cursor= db.cursor()
    #     cursor.execute("SELECT INV_D, PARTS_NO, CURRENT_INV_QTY FROM pdsg0040 where INV_D = TO_DATE(TO_CHAR(SYSDATE-1,'YYYYMMDD'),'YYYYMMDD')")
    #     out_data = cursor.fetchall()
    #     df_smt = pd.DataFrame(out_data)
    #     df_smt.columns = ["출력일", "PARTS NO", "TOTAL 재고"]
    #     return df_smt

    #착공지정일 캘린더 호출
    def selectStartDate(self):
        self.w = CalendarWindow()
        self.w.submitClicked.connect(self.getStartDate)
        self.w.show()
    
    #긴급오더 윈도우 호출
    @pyqtSlot()
    def emgWindow(self):
        self.w = UISubWindow()
        self.w.submitClicked.connect(self.getEmgListview)
        self.w.show()

    #홀딩오더 윈도우 호출
    @pyqtSlot()
    def holdWindow(self):
        self.w = UISubWindow()
        self.w.submitClicked.connect(self.getHoldListview)
        self.w.show()

    #긴급오더 리스트뷰 가져오기
    def getEmgListview(self, list):
        if len(list) > 0 :
            self.listViewEmgLinkage.setModel(list[0])
            self.listViewEmgmscode.setModel(list[1])
            logging.info('긴급오더 리스트를 정상적으로 불러왔습니다.')
        else:
            logging.error('긴급오더 리스트가 없습니다. 다시 한번 확인해주세요')
    
    #홀딩오더 리스트뷰 가져오기
    def getHoldListview(self, list):
        if len(list) > 0 :
            self.listViewHoldLinkage.setModel(list[0])
            self.listViewHoldmscode.setModel(list[1])
            logging.info('홀딩오더 리스트를 정상적으로 불러왔습니다.')
        else:
            logging.error('긴급오더 리스트가 없습니다. 다시 한번 확인해주세요')
    
    #프로그레스바 갱신
    def updateProgressbar(self, val):
        self.progressbar.setValue(val)

    #착공지정일 가져오기
    def getStartDate(self, date):
        if len(date) > 0 :
            self.labelDate.setText(date)
            logging.info('착공지정일이 %s 로 정상적으로 지정되었습니다.', date)
        else:
            logging.error('착공지정일이 선택되지 않았습니다.')

    @pyqtSlot()
    def startLeveling(self):
        #마스터 데이터 불러오기 내부함수
        def loadMasterFile():
            checkFlag = True
            masterFileList = []
            date = datetime.datetime.today().strftime('%Y%m%d')
            if self.isDebug:
                date = self.debugDate.text()

            sosFilePath = r'd:\\FAM3_Leveling\\input\\Master_File\\' + date +r'\\SOS2.xlsx'    
            progressFilePath = r'd:\\FAM3_Leveling\\input\\Master_File\\' + date +r'\\진척.xlsx'
            mainFilePath = r'd:\\FAM3_Leveling\\input\\Master_File\\' + date +r'\\MAIN.xlsx'
            spFilePath = r'd:\\FAM3_Leveling\\input\\Master_File\\' + date +r'\\OTHER.xlsx'
            powerFilePath = r'd:\\FAM3_Leveling\\input\\Master_File\\' + date +r'\\POWER.xlsx'
            calendarFilePath = r'd:\\FAM3_Leveling\\Input\\Calendar_File\\FY' + date[2:4] + '_Calendar.xlsx'
            smtAssyFilePath = r'd:\\FAM3_Leveling\\input\\DB\\MSCode_SMT_Assy.xlsx'
            # usedSmtAssyFilePath = r'.\\input\\DB\\MSCode_SMT_Assy.xlsx'
            secMainListFilePath = r'd:\\FAM3_Leveling\\input\\Master_File\\' + date +r'\\100L1311('+date[4:8]+')MAIN_2차.xlsx'
            inspectFacFilePath = r'd:\\FAM3_Leveling\\input\\DB\\Inspect_Fac.xlsx'
            conditionFilePath = r'd:\\FAM3_Leveling\\input\\MSCODE_Table\\FAM3기종분류표.xlsx'

            pathList = [sosFilePath, 
                        progressFilePath, 
                        mainFilePath, 
                        spFilePath, 
                        powerFilePath, 
                        calendarFilePath, 
                        smtAssyFilePath, 
                        secMainListFilePath, 
                        inspectFacFilePath,
                        conditionFilePath]

            for path in pathList:
                if os.path.exists(path):
                    file = glob.glob(path)[0]
                    masterFileList.append(file)
                else:
                    logging.error('%s 파일이 없습니다. 확인해주세요.', sosFilePath)
                    self.runBtn.setEnabled(True)
                    checkFlag = False
            if checkFlag :
                logging.info('마스터 파일 및 캘린더 파일을 정상적으로 불러왔습니다.')
            return masterFileList
        
        #워킹데이 체크 내부함수
        def checkWorkDay(df, today, compDate):
            dtToday = pd.to_datetime(datetime.datetime.strptime(today, '%Y%m%d'))
            dtComp = pd.to_datetime(compDate, unit='s')
            workDay = 0
            for i in df.index:
                dt = pd.to_datetime(df['Date'][i], unit='s')
                if dtToday < dt and dt <= dtComp:
                    if df['WorkingDay'][i] == 1:
                        workDay += 1
            return workDay

        #콤마 삭제용 내부함수
        def delComma(value):
            return str(value).split('.')[0]

        #디비 불러오기 공통내부함수
        def readDB(ip, port, sid, userName, password, sql):
            location = r'C:\\instantclient_21_6'
            os.environ["PATH"] = location + ";" + os.environ["PATH"]
            dsn = cx_Oracle.makedsn(ip, port, sid)
            db = cx_Oracle.connect(userName, password, dsn)

            cursor= db.cursor()
            cursor.execute(sql)
            out_data = cursor.fetchall()
            df_oracle = pd.DataFrame(out_data)
            col_names = [row[0] for row in cursor.description]
            df_oracle.columns = col_names
            return df_oracle

        #생산시간 합계용 내부함수
        def getSec(time_str):
            time_str = re.sub(r'[^0-9:]', '', str(time_str))
            if len(time_str) > 0:
                h, m, s = time_str.split(':')
                return int(h) * 3600 + int(m) * 60 + int(s)
            else:
                return 0
        # 알람 발생 함수 ksm
        def Alarm_all(df_sum,df_det,div,msc,smt,amo,ate,niz_a,niz_m,msg,ln,oq,sq,pt,nt,ecd):
            if str(div) == '1':
                df_sum = df_sum.append({
                    '분류' : str(div),
                    'MS CODE' : '-',
                    'SMT ASSY' : str(smt),
                    '수량' : int(amo),
                    '검사호기' : '-',
                    '부족 대수(특수,Power)' : 0,
                    '부족 시간(Main)' : 0,
                    'Message' : str(msg)
                    },ignore_index=True)
            elif str(div) == '2':
                df_sum = df_sum.append({
                    '분류' : str(div),
                    'MS CODE' : '-',
                    'SMT ASSY' : '-',
                    '수량' : '-',
                    '검사호기' : str(ate),
                    '부족 대수(특수,Power)' : int(niz_a),
                    '부족 시간(Main)' : 0,
                    'Message' : str(msg)
                    },ignore_index=True)
            elif str(div) == '기타':
                df_sum = df_sum.append({
                    '분류' : str(div),
                    'MS CODE' : str(msc),
                    'SMT ASSY' : '-',
                    '수량' : '-',
                    '검사호기' : '-',
                    '부족 대수(특수,Power)' : 0,
                    '부족 시간(Main)' : 0,
                    'Message' : str(msg)
                    },ignore_index=True)
            df_det = df_det.append({
                '분류':str(div),
                'L/N': str(ln), 
                'MS CODE' : str(msc), 
                'SMT ASSY' : str(smt), 
                '수주수량' : int(oq),
                '부족수량' : int(sq), 
                '검사호기' : str(ate), 
                '대상 검사시간(초)' : int(pt), 
                '필요시간(초)' : int(nt), 
                '완성예정일' : ecd
            },ignore_index=True)
            return(df_sum,df_det)

        def count_emg(model,undo,confirm,Gr): #model=df_addSmtAssy['MODEL'][i], undo=df_addSmtAssy['미착공수주잔'][i], confirm=df_addSmtAssy['착공 확정수량'][i], Gr=1,2차 그룹
            if str(Gr)  == '2':
                dict_secondMaxCnt[dict_modelSecondGr[model]] -= undo #2차 max -= 미착공수주량
                dict_firstMaxCnt[dict_modelFirstGr[model]] -= undo #1차 max -= 미착공수주량
            elif str(Gr) =='1':
                dict_firstMaxCnt[dict_modelFirstGr[model]] -= undo
            module_loading -= undo * dict_capableCnt[model] #최대 착공량 -= 미착공수주량[i] * 공수
            confirm = undo #확정수량[i]=미착공수주량[i]
        
        def count_emg2(model,undo,confirm):
            module_loading -= undo
            confirm = undo
            if module_loading - confirm > 0:
                return (0)
            else:
                return (1)
                #df_SMT_Alarm,df_Spcf_Alarm = Alarm_all(df_SMT_Alarm,df_Spcf_Alarm,'기타2',model,)  알람 처리

        self.runBtn.setEnabled(False)   
        #pandas 경고없애기 옵션 적용
        pd.set_option('mode.chained_assignment', None)

        try:
            list_masterFile = loadMasterFile()
            if len(list_masterFile) > 0 :
                mainOrderCnt = 0.0
                spOrderCnt = 0.0
                powerOrderCnt = 0.0

                #착공량 미입력시의 처리 (추후 멀티프로세싱 적용 시를 위한 처리)
                if len(self.mainOrderinput.text()) <= 0:
                    logging.info('메인기종 착공량이 입력되지 않아 메인기종 착공은 미실시 됩니다.')
                else:
                    mainOrderCnt = float(self.mainOrderinput.text())
                if len(self.spOrderinput.text()) <= 0:
                    logging.info('특수기종 착공량이 입력되지 않아 특수기종 착공은 미실시 됩니다.')
                else:
                    spOrderCnt = float(self.spOrderinput.text())
                if len(self.powerOrderinput.text()) <= 0:
                    logging.info('전원기종 착공량이 입력되지 않아 전원기종 착공은 미실시 됩니다.')            
                else:
                    powerOrderCnt = float(self.powerOrderinput.text())

                #긴급오더, 홀딩오더 불러오기
                emgLinkage = [str(self.listViewEmgLinkage.model().data(self.listViewEmgLinkage.model().index(x,0))) for x in range(self.listViewEmgLinkage.model().rowCount())]
                emgmscode = [self.listViewEmgmscode.model().data(self.listViewEmgmscode.model().index(x,0)) for x in range(self.listViewEmgmscode.model().rowCount())]
                holdLinkage = [str(self.listViewHoldLinkage.model().data(self.listViewHoldLinkage.model().index(x,0))) for x in range(self.listViewHoldLinkage.model().rowCount())]
                holdmscode = [self.listViewHoldmscode.model().data(self.listViewHoldmscode.model().index(x,0)) for x in range(self.listViewHoldmscode.model().rowCount())]        

                #긴급오더, 홀딩오더 데이터프레임화
                df_emgLinkage = pd.DataFrame({'Linkage Number':emgLinkage})
                df_emgmscode = pd.DataFrame({'MS Code':emgmscode})
                df_holdLinkage = pd.DataFrame({'Linkage Number':holdLinkage})
                df_holdmscode = pd.DataFrame({'MS Code':holdmscode})

                #각 Linkage Number 컬럼의 타입을 일치시킴
                df_emgLinkage['Linkage Number'] = df_emgLinkage['Linkage Number'].astype(np.int64)
                df_holdLinkage['Linkage Number'] = df_holdLinkage['Linkage Number'].astype(np.int64)
                
                #긴급오더, 홍딩오더 Join 전 컬럼 추가
                df_emgLinkage['긴급오더'] = '대상'
                df_emgmscode['긴급오더'] = '대상'
                df_holdLinkage['홀딩오더'] = '대상'
                df_holdmscode['홀딩오더'] = '대상'

                #레벨링 리스트 불러오기(멀티프로세싱 적용 후, 분리 예정)
                df_levelingMain = pd.read_excel(list_masterFile[2])
                df_levelingSp = pd.read_excel(list_masterFile[3])
                df_levelingPower = pd.read_excel(list_masterFile[4])

                #미착공 대상만 추출(Main)
                df_levelingMainDropSEQ = df_levelingMain[df_levelingMain['Sequence No'].isnull()]
                df_levelingMainUndepSeq = df_levelingMain[df_levelingMain['Sequence No']=='Undep']
                df_levelingMainUncorSeq = df_levelingMain[df_levelingMain['Sequence No']=='Uncor']
                df_levelingMain = pd.concat([df_levelingMainDropSEQ, df_levelingMainUndepSeq, df_levelingMainUncorSeq])
                df_levelingMain = df_levelingMain.reset_index(drop=True)
                # df_levelingMain['미착공수량'] = df_levelingMain.groupby('Linkage Number')['Linkage Number'].transform('size')

                #미착공 대상만 추출(특수)
                df_levelingSpDropSEQ = df_levelingSp[df_levelingSp['Sequence No'].isnull()]
                df_levelingSpUndepSeq = df_levelingSp[df_levelingSp['Sequence No']=='Undep']
                df_levelingSpUncorSeq = df_levelingSp[df_levelingSp['Sequence No']=='Uncor']
                df_levelingSp = pd.concat([df_levelingSpDropSEQ, df_levelingSpUndepSeq, df_levelingSpUncorSeq])
                df_levelingSp = df_levelingSp.reset_index(drop=True)
                # df_levelingSp['미착공수량'] = df_levelingSp.groupby('Linkage Number')['Linkage Number'].transform('size')

                #미착공 대상만 추출(전원)
                df_levelingPowerDropSEQ = df_levelingPower[df_levelingPower['Sequence No'].isnull()]
                df_levelingPowerUndepSeq = df_levelingPower[df_levelingPower['Sequence No']=='Undep']
                df_levelingPowerUncorSeq = df_levelingPower[df_levelingPower['Sequence No']=='Uncor']
                df_levelingPower = pd.concat([df_levelingPowerDropSEQ, df_levelingPowerUndepSeq, df_levelingPowerUncorSeq])
                df_levelingPower = df_levelingPower.reset_index(drop=True)
                # df_levelingPower['미착공수량'] = df_levelingPower.groupby('Linkage Number')['Linkage Number'].transform('size')

                # if self.isDebug:
                #     df_levelingMain.to_excel('.\\debug\\flow1_main.xlsx')
                #     df_levelingSp.to_excel('.\\debug\\flow1_sp.xlsx')
                #     df_levelingPower.to_excel('.\\debug\\flow1_power.xlsx')

                # 미착공 수주잔 계산
                df_progressFile = pd.read_excel(list_masterFile[1], skiprows=3)
                df_progressFile = df_progressFile.drop(df_progressFile.index[len(df_progressFile.index) - 2:])
                df_progressFile['미착공수주잔'] = df_progressFile['수주\n수량'] - df_progressFile['생산\n지시\n수량']
                df_progressFile['LINKAGE NO'] = df_progressFile['LINKAGE NO'].astype(str).apply(delComma)
                # if self.isDebug:
                #     df_progressFile.to_excel('.\\debug\\flow1.xlsx')

                df_sosFile = pd.read_excel(list_masterFile[0])
                df_sosFile['Linkage Number'] = df_sosFile['Linkage Number'].astype(str)
                # if self.isDebug:
                    # df_sosFile.to_excel('.\\debug\\flow2.xlsx')

                #착공 대상 외 모델 삭제
                df_sosFile = df_sosFile.drop(df_sosFile[df_sosFile['MS Code'].str.contains('ZOTHER')].index)
                df_sosFile = df_sosFile.drop(df_sosFile[df_sosFile['MS Code'].str.contains('YZ')].index)
                df_sosFile = df_sosFile.drop(df_sosFile[df_sosFile['MS Code'].str.contains('SF')].index)
                df_sosFile = df_sosFile.drop(df_sosFile[df_sosFile['MS Code'].str.contains('KM')].index)
                df_sosFile = df_sosFile.drop(df_sosFile[df_sosFile['MS Code'].str.contains('TA80')].index)

                # if self.isDebug:
                    # df_sosFile.to_excel('.\\debug\\flow3.xlsx')

                #워킹데이 캘린더 불러오기
                dfCalendar = pd.read_excel(list_masterFile[5])
                today = datetime.datetime.today().strftime('%Y%m%d')
                if self.isDebug:
                    today = self.debugDate.text()

                #진척 파일 - SOS2파일 Join
                df_sosFileMerge = pd.merge(df_sosFile, df_progressFile, left_on='Linkage Number', right_on='LINKAGE NO', how='left').drop_duplicates(['Linkage Number'])
                #위 파일을 완성지정일 기준 오름차순 정렬 및 인덱스 재설정
                df_sosFileMerge = df_sosFileMerge.sort_values(by=['Planned Prod. Completion date'],
                                                                ascending=[True])
                df_sosFileMerge = df_sosFileMerge.reset_index(drop=True)
                
                #대표모델 Column 생성
                df_sosFileMerge['대표모델'] = df_sosFileMerge['MS Code'].str[:9]
                #남은 워킹데이 Column 생성
                df_sosFileMerge['남은 워킹데이'] = 0
                #긴급오더, 홀딩오더 Linkage Number Column 타입 일치
                df_emgLinkage['Linkage Number'] = df_emgLinkage['Linkage Number'].astype(str)
                df_holdLinkage['Linkage Number'] = df_holdLinkage['Linkage Number'].astype(str)
                #긴급오더, 홀딩오더와 위 Sos파일을 Join
                dfMergeLink = pd.merge(df_sosFileMerge, df_emgLinkage, on='Linkage Number', how='left')
                dfMergemscode = pd.merge(df_sosFileMerge, df_emgmscode, on='MS Code', how='left')
                dfMergeLink = pd.merge(dfMergeLink, df_holdLinkage, on='Linkage Number', how='left')
                dfMergemscode = pd.merge(dfMergemscode, df_holdmscode, on='MS Code', how='left')
                dfMergeLink['긴급오더'] = dfMergeLink['긴급오더'].combine_first(dfMergemscode['긴급오더'])
                dfMergeLink['홀딩오더'] = dfMergeLink['홀딩오더'].combine_first(dfMergemscode['홀딩오더'])

                for i in dfMergeLink.index:
                    dfMergeLink['남은 워킹데이'][i] = checkWorkDay(dfCalendar, today, dfMergeLink['Planned Prod. Completion date'][i])
                    if dfMergeLink['남은 워킹데이'][i] <= 0:
                        dfMergeLink['긴급오더'][i] = '대상'
                # if self.isDebug:
                #     dfMergeLink.to_excel(r'd:\\FAM3_Leveling-1\\input\\flow4.xlsx')

                yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y%m%d')
                if self.isDebug:
                    yesterday = (datetime.datetime.strptime(self.debugDate.text(),'%Y%m%d') - datetime.timedelta(days=1)).strftime('%Y%m%d')

                df_SmtAssyInven = readDB('10.36.15.42',
                                        1521,
                                        'NEURON',
                                        'ymi_user',
                                        'ymi123!',
                                        "SELECT INV_D, PARTS_NO, CURRENT_INV_QTY FROM pdsg0040 where INV_D = TO_DATE("+ str(yesterday) +",'YYYYMMDD')")
                # df_SmtAssyInven.columns = ['INV_D','PARTS_NO','CURRENT_INV_QTY']
                df_SmtAssyInven['현재수량'] = 0
                # print(df_SmtAssyInven)
                # if self.isDebug:
                #     df_SmtAssyInven.to_excel('.\\debug\\flow5.xlsx')

                df_secOrderMainList = pd.read_excel(list_masterFile[7], skiprows=5)
                # print(df_secOrderMainList)
                df_joinSmt = pd.merge(df_secOrderMainList, df_SmtAssyInven, how = 'right', left_on='ASSY NO', right_on='PARTS_NO')
                df_joinSmt['대수'] = df_joinSmt['대수'].fillna(0)
                df_joinSmt['현재수량'] = df_joinSmt['CURRENT_INV_QTY'] - df_joinSmt['대수']
                # df_joinSmt.to_excel(r'd:\\FAM3_Leveling-1\\flow6.xlsx')
                dict_smtCnt = {}
                for i in df_joinSmt.index:
                    dict_smtCnt[df_joinSmt['PARTS_NO'][i]] = df_joinSmt['현재수량'][i]

                df_productTime = readDB('ymzn-bdv19az029-rds.cgbtxsdj6fjy.ap-northeast-1.rds.amazonaws.com',
                                        1521,
                                        'TPROD',
                                        'TEST_SCM',
                                        'test_scm',
                                        'SELECT * FROM FAM3_PRODUCT_TIME_TB')
                df_productTime['TotalTime'] = df_productTime['COMPONENT_SET'].apply(getSec) + df_productTime['MAEDZUKE'].apply(getSec) + df_productTime['MAUNT'].apply(getSec) + df_productTime['LEAD_CUTTING'].apply(getSec) + df_productTime['VISUAL_EXAMINATION'].apply(getSec) + df_productTime['PICKUP'].apply(getSec) + df_productTime['ASSAMBLY'].apply(getSec) + df_productTime['M_FUNCTION_CHECK'].apply(getSec) + df_productTime['A_FUNCTION_CHECK'].apply(getSec) + df_productTime['PERSON_EXAMINE'].apply(getSec)
                df_productTime['대표모델'] = df_productTime['MODEL'].str[:9]
                df_productTime = df_productTime.drop_duplicates(['대표모델'])
                df_productTime = df_productTime.reset_index(drop=True)
                # df_productTime.to_excel(r'd:\\FAM3_Leveling\\flow7.xlsx') #hsj
                # print(df_productTime.columns)

                df_inspectATE = pd.read_excel(list_masterFile[8])
                df_ATEList = df_inspectATE.drop_duplicates(['ATE_NO'])
                df_ATEList = df_ATEList.reset_index(drop=True)
                # df_ATEList.to_excel(r'd:\\FAM3_Leveling\\flow8.xlsx') #hsj
                dict_ate = {}
                max_ateCnt = 0
                for i in df_ATEList.index:
                    if max_ateCnt < len(str(df_ATEList['ATE_NO'][i])):
                        max_ateCnt = len(str(df_ATEList['ATE_NO'][i]))
                    for j in df_ATEList['ATE_NO'][i]:
                        dict_ate[j] = 460 * 60
                # print(dict_ate)
                df_sosAddMainModel = pd.merge(dfMergeLink, df_inspectATE, left_on='대표모델', right_on='MSCODE', how='left')

                df_sosAddMainModel = pd.merge(df_sosAddMainModel, df_productTime[['대표모델','TotalTime','INSPECTION_EQUIPMENT']], on='대표모델', how='left')
                # df_sosAddMainModel.to_excel('.\\debug\\flow9.xlsx') 

                df_mscodeSmtAssy = pd.read_excel(list_masterFile[6])
                df_addSmtAssy = pd.merge(df_sosAddMainModel, df_mscodeSmtAssy, left_on='MS Code', right_on='MS CODE', how='left')
                # for i in range(1,6):
                #     df_addSmtAssy = pd.merge(df_addSmtAssy, df_joinSmt[['PARTS_NO','현재수량']], left_on=f'ROW{str(i)}', right_on='PARTS_NO', how='left')
                #     df_addSmtAssy = df_addSmtAssy.rename(columns = {'현재수량':f'ROW{str(i)}_Cnt'})

                df_addSmtAssy = df_addSmtAssy.drop_duplicates(['Linkage Number'])
                df_addSmtAssy = df_addSmtAssy.reset_index(drop=True)
                # df_addSmtAssy['ATE_NO'] = ''
                # for i in df_addSmtAssy.index:
                #     for j in df_inspectATE.index:
                #         if df_addSmtAssy['대표모델'][i] == df_inspectATE['MSCODE'][j]:
                #             if str(df_addSmtAssy['PRODUCT_TYPE'][i]) == '' or str(df_addSmtAssy['PRODUCT_TYPE'][i]) == 'nan':
                #                 df_addSmtAssy['PRODUCT_TYPE'][i] = df_inspectATE['PRODUCT_TYPE'][j]
                #             if str(df_addSmtAssy['ATE_NO'][i]) == '' or str(df_addSmtAssy['ATE_NO'][i]) == 'nan': 
                #                 df_addSmtAssy['ATE_NO'][i] = df_inspectATE['ATE_NO'][j]
                #             else:
                #                 df_addSmtAssy['ATE_NO'][i] += ',' + df_inspectATE['ATE_NO'][j]
                            
                # df_addSmtAssy.to_excel('.\\debug\\flow9.xlsx') #hsj

                df_addSmtAssy['대표모델별_최소착공필요량_per_일'] = 0
                dict_integCnt = {} #HSJ-대표모델 : 미착공수주잔
                dict_minContCnt = {} #HSJ-대표모델 : 미착공수주잔/WORKDAY

                for i in df_addSmtAssy.index:
                    if df_addSmtAssy['대표모델'][i] in dict_integCnt:
                        dict_integCnt[df_addSmtAssy['대표모델'][i]] += int(df_addSmtAssy['미착공수주잔'][i])
                    else:
                        dict_integCnt[df_addSmtAssy['대표모델'][i]] = int(df_addSmtAssy['미착공수주잔'][i])

                    if df_addSmtAssy['남은 워킹데이'][i] == 0:
                        workDay = 1
                    else:
                        workDay = df_addSmtAssy['남은 워킹데이'][i]
                    
                    if len(dict_minContCnt) > 0:
                        if df_addSmtAssy['대표모델'][i] in dict_minContCnt:
                            if dict_minContCnt[df_addSmtAssy['대표모델'][i]][0] < math.ceil(dict_integCnt[df_addSmtAssy['대표모델'][i]]/workDay):
                                dict_minContCnt[df_addSmtAssy['대표모델'][i]][0] = math.ceil(dict_integCnt[df_addSmtAssy['대표모델'][i]]/workDay)
                                dict_minContCnt[df_addSmtAssy['대표모델'][i]][1] = df_addSmtAssy['Planned Prod. Completion date'][i] #HSJ-완성지정일을 사용하는 이유? - WORKDAY 구분?
                        else:
                            dict_minContCnt[df_addSmtAssy['대표모델'][i]] = [math.ceil(dict_integCnt[df_addSmtAssy['대표모델'][i]]/workDay),
                                                                        df_addSmtAssy['Planned Prod. Completion date'][i]]
                    else:
                        dict_minContCnt[df_addSmtAssy['대표모델'][i]] = [math.ceil(dict_integCnt[df_addSmtAssy['대표모델'][i]]/workDay),
                                                                        df_addSmtAssy['Planned Prod. Completion date'][i]]
                    df_addSmtAssy['대표모델별_최소착공필요량_per_일'][i] = dict_integCnt[df_addSmtAssy['대표모델'][i]]/workDay

                # df_addSmtAssy.to_excel(r'd:\\FAM3_Leveling\\flow9.xlsx')
                
                dict_minContCopy = dict_minContCnt.copy() #HSJ-최소 수량(미착공수주잔/WORKDAY) 복사
                
                df_addSmtAssy['평준화_적용_착공량'] = 0
                for i in df_addSmtAssy.index:
                    if df_addSmtAssy['대표모델'][i] in dict_minContCopy:
                        if dict_minContCopy[df_addSmtAssy['대표모델'][i]][0] >= int(df_addSmtAssy['미착공수주잔'][i]):
                            df_addSmtAssy['평준화_적용_착공량'][i] = int(df_addSmtAssy['미착공수주잔'][i])
                            dict_minContCopy[df_addSmtAssy['대표모델'][i]][0] -= int(df_addSmtAssy['미착공수주잔'][i])
                        else:
                            df_addSmtAssy['평준화_적용_착공량'][i] = dict_minContCopy[df_addSmtAssy['대표모델'][i]][0]
                            dict_minContCopy[df_addSmtAssy['대표모델'][i]][0] = 0

                df_addSmtAssy['잔여_착공량'] = df_addSmtAssy['미착공수주잔'] - df_addSmtAssy['평준화_적용_착공량']


                df_smtCopy = pd.DataFrame(columns=df_addSmtAssy.columns)
                df_addSmtAssy = df_addSmtAssy.sort_values(by=['긴급오더',
                                                                'Planned Prod. Completion date',
                                                                '평준화_적용_착공량'],
                                                                ascending=[False,
                                                                            True,
                                                                            False])
                # df_addSmtAssy.to_excel(r'd:\\FAM3_Leveling\\flow10.xlsx') #hsj

                #알람 부분
                df_SMT_Alarm = pd.DataFrame(columns={'분류','MS CODE','SMT ASSY','수량','검사호기','부족 대수(특수,Power)','부족 시간(Main)','Message'},dtype=str)
                df_SMT_Alarm['수량'] = df_SMT_Alarm['수량'] .astype(int)
                df_SMT_Alarm['부족 시간(Main)'] =df_SMT_Alarm['부족 시간(Main)'].astype(int)
                df_SMT_Alarm['부족 대수(특수,Power)'] =df_SMT_Alarm['부족 대수(특수,Power)'].astype(int)
                df_SMT_Alarm = df_SMT_Alarm[['분류','MS CODE','SMT ASSY','수량','검사호기','부족 대수(특수,Power)','부족 시간(Main)','Message']]
                df_Spcf_Alarm = pd.DataFrame(columns={'분류','L/N','MS CODE','SMT ASSY','수주수량','부족수량','검사호기','대상 검사시간(초)','필요시간(초)','완성예정일'},dtype=str)
                df_Spcf_Alarm['수주수량'] = df_Spcf_Alarm['수주수량'] .astype(int)
                df_Spcf_Alarm['부족수량'] =df_Spcf_Alarm['부족수량'].astype(int)
                df_Spcf_Alarm['대상 검사시간(초)'] =df_Spcf_Alarm['대상 검사시간(초)'].astype(int)
                df_Spcf_Alarm['필요시간(초)'] =df_Spcf_Alarm['필요시간(초)'].astype(int)
                #df_Spcf_Alarm['완성예정일'] =df_Spcf_Alarm['완성예정일'].astype(datetime.datetime)
                df_Spcf_Alarm = df_Spcf_Alarm[['분류','L/N','MS CODE','SMT ASSY','수주수량','부족수량','검사호기','대상 검사시간(초)','필요시간(초)','완성예정일']]
                
                rowCnt = 0
                # df_addSmtAssy = df_addSmtAssy[df_addSmtAssy['PRODUCT_TYPE']=='MAIN'] #HSJ DEL
                df_addSmtAssy = df_addSmtAssy[df_addSmtAssy['PRODUCT_TYPE']=='OTHER']
                df_addSmtAssy = df_addSmtAssy.reset_index(drop=True)
                df_addSmtAssy['SMT반영_착공량'] = 0
                for i in df_addSmtAssy.index:
                    if df_addSmtAssy['PRODUCT_TYPE'][i] == 'OTHER':
                    # if df_addSmtAssy['PRODUCT_TYPE'][i] == 'MAIN': #HSJ-DEL
                        for j in range(1,6):
                            if j == 1:
                                rowCnt = 1
                            if str(df_addSmtAssy[f'ROW{str(j)}'][i]) != '' and str(df_addSmtAssy[f'ROW{str(j)}'][i]) != 'nan':
                                rowCnt = j
                            else:
                                #ksm 알람
                                break
                        smtFlag = False    
                        minCnt = 9999
                        for j in range(1,rowCnt+1):
                                smtAssyName = str(df_addSmtAssy[f'ROW{str(j)}'][i])
                                if smtAssyName != '' and smtAssyName != 'nan':
                                    if df_addSmtAssy['긴급오더'][i] == '대상':
                                        dict_smtCnt[smtAssyName] -= df_addSmtAssy['평준화_적용_착공량'][i]
                                        # df_smtCopy = df_smtCopy.append(df_addSmtAssy.iloc[i])
                                        if dict_smtCnt[smtAssyName] < 0:
                                            logging.warning('「당일착공 대상 : %s」, 「사양 : %s」을 착공하기에는 「SmtAssy : %s」가 「%i 대」부족합니다. SmtAssy 제작을 지시해주세요. 당일착공 대상이므로 착공은 진행합니다.',
                                                            df_addSmtAssy['Linkage Number'][i],
                                                            df_addSmtAssy['MS Code'][i],
                                                            smtAssyName,
                                                            0 - dict_smtCnt[smtAssyName])
                                    else:
                                        if dict_smtCnt[smtAssyName] >= df_addSmtAssy['평준화_적용_착공량'][i]:
                                            if minCnt > df_addSmtAssy['평준화_적용_착공량'][i]:
                                                minCnt = df_addSmtAssy['평준화_적용_착공량'][i]
                                            # dict_smtCnt[smtAssyName] -= df_addSmtAssy['미착공수량'][i]
                                            # df_smtCopy = df_smtCopy.append(df_addSmtAssy.iloc[i])
                                        elif dict_smtCnt[smtAssyName] > 0:
                                            if minCnt > dict_smtCnt[smtAssyName]:
                                                minCnt = dict_smtCnt[smtAssyName]
                                            # df_addSmtAssy['미착공수량'][i] = dict_smtCnt[smtAssyName]
                                            # dict_smtCnt[smtAssyName] -= df_addSmtAssy['미착공수량'][i]
                                            # df_smtCopy = df_smtCopy.append(df_addSmtAssy.iloc[i])
                                        else:
                                            minCnt = 0
                                            logging.warning('「사양 : %s」을 착공하기에는 「SmtAssy : %s」가 부족합니다. SmtAssy 제작을 지시해주세요.',
                                                            df_addSmtAssy['MS Code'][i],
                                                            smtAssyName)
                                else:
                                    logging.warning('「사양 : %s」의 SmtAssy가 %s 파일에 등록되지 않았습니다. 등록 후, 다시 실행해주세요.',
                                                    df_addSmtAssy['MS Code'][i],
                                                    list_masterFile[6])

                        if minCnt != 9999:
                            df_addSmtAssy['SMT반영_착공량'][i] = minCnt
                        else:
                            df_addSmtAssy['SMT반영_착공량'][i] = df_addSmtAssy['평준화_적용_착공량'][i]
                                        
                        if str(df_addSmtAssy['ATE_NO'][i]) !='' and str(df_addSmtAssy['ATE_NO'][i]) !='nan':
                            for j in range(0,len(str(df_addSmtAssy['ATE_NO'][i]))):
                                df_addSmtAssy['ATE_NO'][i][j]

                #HSJ-SMT 재고 수량 반영 잔여 착공량 구하기       
                df_addSmtAssy['SMT반영_잔여_착공량'] = 0
                for i in df_addSmtAssy.index: 
                    if df_addSmtAssy['PRODUCT_TYPE'][i] == 'OTHER':
                        if df_addSmtAssy['잔여_착공량'][i] > 0:
                            for j in range(1,6):
                                if j == 1:
                                    rowCnt = 1
                                if str(df_addSmtAssy[f'ROW{str(j)}'][i]) != '' and str(df_addSmtAssy[f'ROW{str(j)}'][i]) != 'nan':
                                    rowCnt = j
                                else:

                                    break
                            smtFlag = False    
                            minCnt = 9999
                            for j in range(1,rowCnt+1):
                                    smtAssyName = str(df_addSmtAssy[f'ROW{str(j)}'][i])
                                    if smtAssyName != '' and smtAssyName != 'nan':
                                        if df_addSmtAssy['긴급오더'][i] == '대상':
                                            dict_smtCnt[smtAssyName] -= df_addSmtAssy['잔여_착공량'][i]
                                            # df_smtCopy = df_smtCopy.append(df_addSmtAssy.iloc[i])
                                            if dict_smtCnt[smtAssyName] < 0:
                                                logging.warning('「당일착공 대상 : %s」, 「사양 : %s」을 착공하기에는 「SmtAssy : %s」가 「%i 대」부족합니다. SmtAssy 제작을 지시해주세요. 당일착공 대상이므로 착공은 진행합니다.',
                                                                df_addSmtAssy['Linkage Number'][i],
                                                                df_addSmtAssy['MS Code'][i],
                                                                smtAssyName,
                                                                0 - dict_smtCnt[smtAssyName])
                                        else:
                                            if dict_smtCnt[smtAssyName] >= df_addSmtAssy['잔여_착공량'][i]:
                                                if minCnt > df_addSmtAssy['잔여_착공량'][i]:
                                                    minCnt = df_addSmtAssy['잔여_착공량'][i]
                                                # dict_smtCnt[smtAssyName] -= df_addSmtAssy['미착공수량'][i]
                                                # df_smtCopy = df_smtCopy.append(df_addSmtAssy.iloc[i])
                                            elif dict_smtCnt[smtAssyName] > 0:
                                                if minCnt > dict_smtCnt[smtAssyName]:
                                                    minCnt = dict_smtCnt[smtAssyName]
                                                # df_addSmtAssy['미착공수량'][i] = dict_smtCnt[smtAssyName]
                                                # dict_smtCnt[smtAssyName] -= df_addSmtAssy['미착공수량'][i]
                                                # df_smtCopy = df_smtCopy.append(df_addSmtAssy.iloc[i])
                                            else:
                                                minCnt = 0
                                                logging.warning('「사양 : %s」을 착공하기에는 「SmtAssy : %s」가 부족합니다. SmtAssy 제작을 지시해주세요.',
                                                                df_addSmtAssy['MS Code'][i],
                                                                smtAssyName)
                                    else:
                                        logging.warning('「사양 : %s」의 SmtAssy가 %s 파일에 등록되지 않았습니다. 등록 후, 다시 실행해주세요.',
                                                        df_addSmtAssy['MS Code'][i],
                                                        list_masterFile[6])
                            if minCnt != 9999:
                                df_addSmtAssy['SMT반영_잔여_착공량'][i] = minCnt
                            else:
                                df_addSmtAssy['SMT반영_잔여_착공량'][i] = df_addSmtAssy['평준화_적용_착공량'][i]


                # df_addSmtAssy.to_excel(r'd:\\FAM3_Leveling\\flow11.xlsx')
                df_addSmtAssy['임시수량'] = 0
                df_addSmtAssy['설비능력반영_착공량'] = 0
                
                # for i in df_addSmtAssy.index:  
                #     if str(df_addSmtAssy['TotalTime'][i]) != '' and str(df_addSmtAssy['TotalTime'][i]) != 'nan':
                #         if str(df_addSmtAssy['ATE_NO'][i]) != '' and str(df_addSmtAssy['ATE_NO'][i]) != 'nan':
                #             tempTime = 0
                #             ateName = ''
                #             for ate in df_addSmtAssy['ATE_NO'][i]:
                #                 if tempTime < dict_ate[ate]:
                #                     tempTime = dict_ate[ate]
                #                     ateName = ate
                #             if dict_ate[ateName] >= df_addSmtAssy['TotalTime'][i] * df_addSmtAssy['SMT반영_착공량'][i]:
                #                 dict_ate[ateName] -= df_addSmtAssy['TotalTime'][i] * df_addSmtAssy['SMT반영_착공량'][i]
                #                 df_addSmtAssy['설비능력반영_착공량'][i] = df_addSmtAssy['SMT반영_착공량'][i]
                #             elif dict_ate[ateName] >= df_addSmtAssy['TotalTime'][i]:
                #                 tempCnt = int(df_addSmtAssy['SMT반영_착공량'][i])
                #                 for j in range(tempCnt,0,-1):
                #                     # print(dict_ate[ateName])
                #                     # print(int(df_addSmtAssy['TotalTime'][i]) * j)
                #                     if dict_ate[ateName] >= int(df_addSmtAssy['TotalTime'][i]) * j:
                #                         df_addSmtAssy['설비능력반영_착공량'][i] = j
                #                         dict_ate[ateName] -= int(df_addSmtAssy['TotalTime'][i]) * j
                #                         break
                
                #HSJ DEL 
                # for i in df_addSmtAssy.index:  
                #     if str(df_addSmtAssy['TotalTime'][i]) != '' and str(df_addSmtAssy['TotalTime'][i]) != 'nan':
                #         if str(df_addSmtAssy['ATE_NO'][i]) != '' and str(df_addSmtAssy['ATE_NO'][i]) != 'nan':
                #             tempTime = 0
                #             ateName = ''
                #             for ate in df_addSmtAssy['ATE_NO'][i]:
                #                 if tempTime < dict_ate[ate]:
                #                     tempTime = dict_ate[ate]
                #                     ateName = ate
                #                     if ate == df_addSmtAssy['ATE_NO'][i][0]:
                #                         df_addSmtAssy['임시수량'][i] = df_addSmtAssy['SMT반영_착공량'][i]
                #                     if df_addSmtAssy['임시수량'][i] != 0:
                #                         if dict_ate[ateName] >= df_addSmtAssy['TotalTime'][i] * df_addSmtAssy['임시수량'][i]:
                #                             dict_ate[ateName] -= df_addSmtAssy['TotalTime'][i] * df_addSmtAssy['임시수량'][i]
                #                             df_addSmtAssy['설비능력반영_착공량'][i] += df_addSmtAssy['임시수량'][i]
                #                             df_addSmtAssy['임시수량'][i] = 0
                #                             break
                #                         elif dict_ate[ateName] >= df_addSmtAssy['TotalTime'][i]:
                #                             tempCnt = int(df_addSmtAssy['임시수량'][i])
                #                             for j in range(tempCnt,0,-1):
                #                                 # print(dict_ate[ateName])
                #                                 # print(int(df_addSmtAssy['TotalTime'][i]) * j)
                #                                 if dict_ate[ateName] >= int(df_addSmtAssy['TotalTime'][i]) * j:
                #                                     df_addSmtAssy['설비능력반영_착공량'][i] = int(df_addSmtAssy['설비능력반영_착공량'][i]) + j
                #                                     dict_ate[ateName] -= int(df_addSmtAssy['TotalTime'][i]) * j
                #                                     df_addSmtAssy['임시수량'][i] = tempCnt - j
                #                                     break
                #                     else:
                #                         break
                                        
                #             # print(i)
                #             # print(f'설비명 : {ateName}')
                #             # print('남은시간 : ' + str((dict_ate[ateName])))
                #HSJ DEL

                # df_addSmtAssy.to_excel(r'd:\\FAM3_Leveling-1\\flow12.xlsx')

                # df_addSmtAssy['대표모델별_누적착공량'] = ''
                # dict_integAteCnt = {}
                # for i in df_addSmtAssy.index:
                #     if df_addSmtAssy['대표모델'][i] in dict_integAteCnt:
                #         dict_integAteCnt[df_addSmtAssy['대표모델'][i]] += int(df_addSmtAssy['설비능력반영_착공량'][i])
                #     else:
                #         dict_integAteCnt[df_addSmtAssy['대표모델'][i]] = int(df_addSmtAssy['설비능력반영_착공량'][i])
                #     df_addSmtAssy['대표모델별_누적착공량'][i] = dict_integAteCnt[df_addSmtAssy['대표모델'][i]]

                # for key, value in dict_minContCnt.items():
                #     if key in dict_integAteCnt:
                #         if value[0] > dict_integAteCnt[key]:
                #             logging.warning('「%s」 사양이 「완성지정일: %s」 까지 오늘 「착공수량: %i 대」로는 착공량 부족이 예상됩니다. 최소 필요 착공량은 「%i 대」 입니다.', 
                #                 key, 
                #                 str(value[1]),
                #                 dict_integAteCnt[key],
                #                 math.ceil(value[0]))      

                # df_addSmtAssy.to_excel('.\\debug\\flow13.xlsx')
                #HSJ 상기 누적 착공량을 계산해야 하는 이유 = 공수 반영을 해야하기 때문, 하지만 상기 코드를 사용해도 될지는 아직 모름 

                #HSJ 특수 기종 분류표 적용 start
                df_addSmtAssy['착공 확정수량'] = 0
                df_condition= pd.read_excel(list_masterFile[9]) #FAM3 기종 분류표 불러오기
                
                df_condition['No'] = df_condition['No'].fillna(method='ffill')
                df_condition['1차_MAX_그룹'] = df_condition['1차_MAX_그룹'].fillna(method='ffill')
                df_condition['2차_MAX_그룹'] = df_condition['2차_MAX_그룹'].fillna(method='ffill')
                df_condition['1차_MAX'] = df_condition['1차_MAX'].fillna(method='ffill')
                df_condition['2차_MAX'] = df_condition['2차_MAX'].fillna(method='ffill')


                dict_capableCnt = defaultdict(list) #모델 : 공수
                dict_firstMaxCnt = defaultdict(list) #1차_MAX_그룹 : 1차_MAX
                dict_secondMaxCnt = defaultdict(list) #2차_MAX_그룹 : 2차_MAX
                dict_module = defaultdict(list) #모델 : 구분 
                dict_modelFirstGr = defaultdict(list) # 모델 : 1차_MAX_그룹
                dict_modelSecondGr = defaultdict(list) # 모델 : 2차_MAX_그룹
                module_loading = float(self.spOrderinput.text()) #모듈 착공 필요 수량(사용자 기입)
                notmodule_loading =  float(self.spOrderinput.text())  #비모듈 착공 필요 수량(사용자 기입)

                # for i in df_condition.index:
                #     dict_capableCnt[df_condition['MODEL'][i]] = df_condition['공수'][i]
                #     dict_firstMaxCnt[df_condition['1차_MAX_그룹'][i]] = df_condition['1차_MAX'][i]
                #     dict_secondMaxCnt[df_condition['2차_MAX_그룹'][i]] = df_condition['2차_MAX'][i]
                #     dict_module[df_condition['MODEL'][i]] = df_condition['구분'][i] #모듈, 비모듈 구분 용도
                #     dict_modelFirstGr[df_condition['MODEL'][i]] = df_condition['1차_MAX_그룹'][i] #모델 : 1차 그룹 매칭 용도
                #     dict_modelSecondGr[df_addSmtAssy['MODEL'][i]] = df_condition['2차_MAX_그룹'][i] #모델 : 2차 그룹 매칭 용도
                #딕셔너리 설정
                for i in df_condition.index:
                    dict_capableCnt[df_condition['MODEL'][i]] = df_condition['공수'][i]
                    dict_module[df_condition['MODEL'][i]] = df_condition['구분'][i] #모듈, 비모듈 구분 용도
                    if df_condition['2차_MAX_그룹'][i] != '-':
                        dict_firstMaxCnt[df_condition['1차_MAX_그룹'][i]] = df_condition['1차_MAX'][i]
                        dict_modelFirstGr[df_condition['MODEL'][i]] = df_condition['1차_MAX_그룹'][i] #모델 : 1차 그룹 매칭 용도
                        dict_secondMaxCnt[df_condition['2차_MAX_그룹'][i]] = df_condition['2차_MAX'][i]
                        dict_modelSecondGr[df_condition['MODEL'][i]] = df_condition['2차_MAX_그룹'][i] #모델 : 2차 그룹 매칭 용도
                    elif df_condition['1차_MAX_그룹'][i] != '-':
                        dict_firstMaxCnt[df_condition['1차_MAX_그룹'][i]] = df_condition['1차_MAX'][i]
                        dict_modelFirstGr[df_condition['MODEL'][i]] = df_condition['1차_MAX_그룹'][i]
                    else:
                        continue
                
                
                for i in df_addSmtAssy.index:
                    if df_addSmtAssy['긴급오더'][i] == '대상':
                        if df_addSmtAssy['MODEL'][i] in dict_module.keys(): #기종분류표에 model이 있는가
                            if dict_module[df_addSmtAssy['MODEL'][i]] == '모듈': #기종분류표 '구분'이 모듈인가
                                if df_addSmtAssy['MODEL'][i] in dict_modelSecondGr.keys(): #2차 max값 유무
                                    count_emg(df_addSmtAssy['MODEL'][i],df_addSmtAssy['미착공수주잔'][i],df_addSmtAssy['착공 확정수량'][i],2) #긴급 오더-착공량 만큼 뺄거뺌
                                    #if
                                else: 
                                    if df_addSmtAssy['MODEL'][i] in dict_modelSecondGr.keys(): #1차 max값 유무
                                        count_emg(df_addSmtAssy['MODEL'][i],df_addSmtAssy['미착공수주잔'][i],df_addSmtAssy['착공 확정수량'][i],1)
                                    else:
                                        a = count_emg2(df_addSmtAssy['MODEL'][i],df_addSmtAssy['미착공수주잔'][i],df_addSmtAssy['착공 확정수량'][i])
                                        #return 값에 따라 에러 처리
                                        if a == 0:
                                            continue
                                        else:
                                            #알람 처리


                #flow 수정
                for i in df_addSmtAssy.index:
                    if df_addSmtAssy['긴급오더'][i] == '대상':
                        if df_addSmtAssy['MODEL'][i] in dict_module.keys(): #기종분류표에 model이 있는가
                            if df_addSmtAssy['MODEL'][i] in dict_modelSecondGr.keys(): #2차 max값 유무 ver3
                                if df_addSmtAssy['MODEL'][i] in dict_modelFirstGr.keys(): #1차 max값 유무 ver3
                                    if module_loading - df_addSmtAssy['미착공수주잔'][i] * dict_capableCnt[df_condition['MODEL'][i]] > 0: #최대 착공량 -(미착공 수주량*공수) > 0
                                        if dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] - df_addSmtAssy['미착공수주잔'][i] >= 0: #2차 max - 미착공수주량 >= 0
                                            dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] -= df_addSmtAssy['미착공수주잔'][i]
                                            dict_firstMaxCnt[dict_modelFirstGr[df_condition['MODEL'][i]]] -= df_addSmtAssy['미착공수주잔'][i]
                                            if dict_firstMaxCnt[dict_modelFirstGr[df_condition['MODEL'][i]]] > 0: #1차 MAX > 0
                                                module_loading -= df_addSmtAssy['미착공수주잔'][i] #최대 착공량 -= 미착공 수주량[i]
                                                df_addSmtAssy['착공 확정수량'][i] = df_addSmtAssy['미착공수주잔'][i] #확정 수량[i] = 미착공 수주량
                                            else:
                                                module_loading -= df_addSmtAssy['미착공수주잔'][i] #최대 착공량 -= 미착공수주량
                                                df_addSmtAssy['착공 확정수량'] = df_addSmtAssy['미착공수주잔'][i]
                                                #분류2 - 1차 max 알람처리
                                        else:
                                            dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] -= df_addSmtAssy['미착공수주잔'][i]
                                            dict_firstMaxCnt[dict_modelFirstGr[df_condition['MODEL'][i]]] -= df_addSmtAssy['미착공수주잔'][i]
                                            if dict_firstMaxCnt[dict_modelFirstGr[df_condition['MODEL'][i]]] > 0: #1차 MAX > 0
                                                module_loading -= df_addSmtAssy['미착공수주잔'][i] #최대 착공량 -= 미착공수주량
                                                df_addSmtAssy['착공 확정수량'] = df_addSmtAssy['미착공수주잔'][i]
                                                #분류2 - 2차 MAX, 알람처리
                                            else:
                                                module_loading -= df_addSmtAssy['미착공수주잔'][i] #최대 착공량 -= 미착공수주량
                                                df_addSmtAssy['착공 확정수량'] = df_addSmtAssy['미착공수주잔'][i]
                                            
                                                #분류2 - 1차, 2차 MAX, 알람처리
                                    else:
                                        if dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] - df_addSmtAssy['미착공수주잔'][i] >= 0: #2차 max - 미착공수주량 >= 0
                                            dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] -= df_addSmtAssy['미착공수주잔'][i]
                                            dict_firstMaxCnt[dict_modelFirstGr[df_condition['MODEL'][i]]] -= df_addSmtAssy['미착공수주잔'][i]
                                            if dict_firstMaxCnt[dict_modelFirstGr[df_condition['MODEL'][i]]] > 0: #1차 MAX > 0
                                                module_loading -= df_addSmtAssy['미착공수주잔'][i] #최대 착공량 -= 미착공 수주량[i]
                                                df_addSmtAssy['착공 확정수량'] = df_addSmtAssy['미착공수주잔'][i] #확정 수량[i] = 미착공 수주량
                                                #기타2 - 최대 착공량 알람처리
                                            else:
                                                module_loading -= df_addSmtAssy['미착공수주잔'][i] #최대 착공량 -= 미착공수주량
                                                df_addSmtAssy['착공 확정수량'] = df_addSmtAssy['미착공수주잔'][i]
                                                #기타2 최대 착공량, 분류2 - 1차 max 알람 처리
                                        else:
                                            dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] -= df_addSmtAssy['미착공수주잔'][i]
                                            dict_firstMaxCnt[dict_modelFirstGr[df_condition['MODEL'][i]]] -= df_addSmtAssy['미착공수주잔'][i]
                                            if dict_firstMaxCnt[dict_modelFirstGr[df_condition['MODEL'][i]]] > 0: #1차 MAX > 0
                                                module_loading -= df_addSmtAssy['미착공수주잔'][i] #최대 착공량 -= 미착공수주량
                                                df_addSmtAssy['착공 확정수량'] = df_addSmtAssy['미착공수주잔'][i]
                                                #기타2 - 최대착공량, 분류2 - 2차 max 알람처리
                                            else:
                                                module_loading -= df_addSmtAssy['미착공수주잔'][i] #최대 착공량 -= 미착공수주량
                                                df_addSmtAssy['착공 확정수량'] = df_addSmtAssy['미착공수주잔'][i]
                                                #기타2 - 최대착공량, 분류2 - 1차, 2차 max 알람처리
                                else:
                            else:
                        else:

                    else:


                

                df_SMT_Alarm = df_SMT_Alarm.drop_duplicates(subset=['검사호기','분류','Message','MS CODE','SMT ASSY'],keep='last')
                df_Spcf_Alarm = df_Spcf_Alarm.drop_duplicates(subset=['분류','L/N','MS CODE','완성예정일'],keep='last')
                df_addSmtAssy = df_addSmtAssy.reset_index(drop=True)
                df_SMT_Alarm = df_SMT_Alarm.sort_values(by=['분류',
                                                            '수량'],
                                                            ascending=[True,
                                                                        True])
                df_Spcf_Alarm = df_Spcf_Alarm.sort_values(by=['분류',
                                                                '완성예정일',
                                                                'MS CODE',
                                                                'SMT ASSY'],
                                                                ascending=[True,
                                                                            True,
                                                                            True,
                                                                            True])
                #알람 출력
                df_SMT_Alarm = df_SMT_Alarm.reset_index(drop=True)
                df_SMT_Alarm.index = df_SMT_Alarm.index+1
                df_Spcf_Alarm = df_Spcf_Alarm.reset_index(drop=True)
                df_Spcf_Alarm.index = df_Spcf_Alarm.index+1
                df_explain = pd.DataFrame({'분류': ['1','2','기타','폴더','파일명'] ,
                                            '분류별 상황' : ['DB상의 Smt Assy가 부족하여 해당 MS-Code를 착공 내릴 수 없는 경우',
                                                            '당일 착공분(or 긴급착공분)에 대해 검사설비 능력이 부족할 경우',
                                                            'MS-Code와 일치하는 Smt Assy가 마스터 파일에 없는 경우',
                                                            'output ➡ alarm',
                                                            'FAM3_AlarmList_20221028_시분초']})
                Alarmdate = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                Alarm_path = r'.\\input\\AlarmList_Power\\FAM3_AlarmList_' + Alarmdate + r'.xlsx'
                writer = pd.ExcelWriter(Alarm_path,engine='xlsxwriter')
                df_SMT_Alarm.to_excel(writer,sheet_name='정리')
                df_Spcf_Alarm.to_excel(writer,sheet_name='상세')
                df_explain.to_excel(writer,sheet_name='설명')
                writer.save()       







                        


            

            self.runBtn.setEnabled(True)
        except Exception as e:
            logging.exception(e, exc_info=True)                     
            self.runBtn.setEnabled(True)
if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    ui = Ui_MainWindow()
