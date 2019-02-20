# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'travel_fee_calc_form_2.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import os, sys, sip, traceback
from PyQt5.QtWidgets import QMessageBox, QFileDialog, QTableWidgetItem, QDialog, QDesktopWidget
from pandas import read_csv, read_excel, ExcelFile, DataFrame, ExcelWriter, isnull
from datetime import date, timedelta
from interval import Interval
from itertools import permutations
# from numpy import isnan

class Ui_Form(object):
    def setupUi(self, Form, screen_height, screen_width):
        Form.setObjectName("Form")
        Form.resize(750*screen_width/1920, 900*screen_height/1080)
        self.font_size = 12
        self.screen_height, self.screen_width = screen_height, screen_width
        self.datapath = ''
        self.pushButton = QtWidgets.QPushButton(Form)
        left, top, width, height = self.resize_screen(310, 40, 41, 31)
        self.pushButton.setGeometry(QtCore.QRect(left, top, width, height)) # 310, 40, 41, 31
        self.pushButton.setStyleSheet("background-image: url(:/form/Magnifier_32px.png);")
        self.pushButton.setObjectName("pushButton")
        self.textBrowser = QtWidgets.QTextBrowser(Form)
        left, top, width, height = self.resize_screen(370, 530, 361, 321) # 370, 740, 361, 111
        self.textBrowser.setGeometry(QtCore.QRect(left, top, width, height)) # 370, 530, 361, 321
        self.textBrowser.setObjectName("textBrowser")
        self.textBrowser_2 = QtWidgets.QTextBrowser(Form)
        left, top, width, height = self.resize_screen(20, 830, 331, 61)
        self.textBrowser_2.setGeometry(QtCore.QRect(left, top, width, height))
        self.textBrowser_2.setObjectName("textBrowser_2")
        self.lineEdit = QtWidgets.QLineEdit(Form)
        left, top, width, height = self.resize_screen(20, 40, 291, 31)
        self.lineEdit.setGeometry(QtCore.QRect(left, top, width, height))
        self.lineEdit.setContextMenuPolicy(QtCore.Qt.NoContextMenu)
        self.lineEdit.setStyleSheet("border-color: rgb(0, 0, 255);")
        self.lineEdit.setObjectName("lineEdit")

        self.lineEdit_2 = QtWidgets.QLineEdit(Form)
        left, top, width, height = self.resize_screen(20, 110, 291, 31)
        self.lineEdit_2.setGeometry(QtCore.QRect(left, top, width, height))
        self.lineEdit_2.setContextMenuPolicy(QtCore.Qt.NoContextMenu)
        self.lineEdit_2.setStyleSheet("border-color: rgb(0, 0, 255);")
        self.lineEdit_2.setObjectName("lineEdit_2")

        self.pushButton_2 = QtWidgets.QPushButton(Form)
        left, top, width, height = self.resize_screen(590, 860, 141, 31)
        self.pushButton_2.setGeometry(QtCore.QRect(left, top, width, height))
        self.pushButton_2.setObjectName("pushButton_2")

        self.verticalWidget = QtWidgets.QWidget(Form)
        left, top, width, height = self.resize_screen(20, 150, 331, 641)
        self.verticalWidget.setGeometry(QtCore.QRect(left, top, width, height))
        self.verticalWidget.setObjectName("verticalWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")

        self.label = QtWidgets.QLabel(Form)
        left, top, width, height = self.resize_screen(20, 10, 200, 21)
        self.label.setGeometry(QtCore.QRect(left, top, width, height))
        self.font_size = max(6,int(self.font_size * width / 200))
        self.label.setStyleSheet("font: " + str(self.font_size) + "pt \"Adobe Arabic\";")
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Form)
        left, top, width, height = self.resize_screen(20, 80, 200, 21)
        self.label_2.setGeometry(QtCore.QRect(left, top, width, height))
        self.font_size = max(6,int(self.font_size * width / 200))
        self.label_2.setStyleSheet("font: " + str(self.font_size) + "pt \"Adobe Arabic\";")
        self.label_2.setObjectName("label_2")

        self.tableWidget = QtWidgets.QTableWidget(Form)
        left, top, width, height = self.resize_screen(370, 40, 361, 451) # 370, 40, 361, 661
        self.tableWidget.setGeometry(QtCore.QRect(left, top, width, height)) 
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(2)
        self.tableWidget.setRowCount(100)
        self.tableWidget.setColumnWidth(0,210*width/361) #将第一列的单元宽度设置为210
        self.tableWidget.setColumnWidth(1,100*width/361) #将第二列的单元宽度设置为100
        self.tableWidget.setHorizontalHeaderLabels(['变量','数值'])

        self.label_3 = QtWidgets.QLabel(Form)
        left, top, width, height = self.resize_screen(370, 10, 200, 21)
        self.label_3.setGeometry(QtCore.QRect(left, top, width, height))
        self.font_size = max(6,int(self.font_size * width / 200))
        self.label_3.setStyleSheet("font: " + str(self.font_size) + "pt \"Adobe Arabic\";")
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(Form)
        left, top, width, height = self.resize_screen(370, 500, 200, 21) #370, 710, 161, 21
        self.label_4.setGeometry(QtCore.QRect(left, top, width, height)) # 
        self.font_size = max(6,int(self.font_size * width / 161))
        self.label_4.setStyleSheet("font: " + str(self.font_size) + "pt \"Adobe Arabic\";")
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(Form)
        left, top, width, height = self.resize_screen(20, 800, 200, 21)
        self.label_5.setGeometry(QtCore.QRect(left, top, width, height))
        self.font_size = max(6,int(self.font_size * width / 200))
        self.label_5.setStyleSheet("font: " + str(self.font_size) + "pt \"Adobe Arabic\";")
        self.label_5.setObjectName("label_5")

        self.pushButton_3 = QtWidgets.QPushButton(Form)
        left, top, width, height = self.resize_screen(370, 860, 151, 31)
        self.pushButton_3.setGeometry(QtCore.QRect(left, top, width, height))
        self.pushButton_3.setObjectName("pushButton_3")

        self.pushButton_4 = QtWidgets.QPushButton(Form)
        left, top, width, height = self.resize_screen(310, 110, 41, 31)
        self.pushButton_4.setGeometry(QtCore.QRect(left, top, width, height))
        self.pushButton_4.setObjectName("pushButton_4")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

        self.searchFile = ''
        self.table_name_list = []
        self.radioButtonNum = 0
        self.textBrowser2Set = set()
        self.search_table_name_list = []
        self.path_setchecked_dict = {}
        self.tb2_string_output = [] #根据tb2_string_output进行步骤3
        self.history_tb2_string_output = []
        # version 1
        # self.table_item_list = ['入住时间','免房费天数','成人总人数(包括第三成人)','第三成人人数','儿童人数','婴儿人数']
        # version 2
        # self.table_item_list = ['入住时间','优惠折扣','成人总人数(包括第三成人)','是否有第三成人','儿童人数','婴儿人数','免房费天数','免费夜固定价格',\
        #                         '第三成人免费夜固定价格','强制费单价','强制费收取天数','强制餐单价','房价是否包含环境税',\
        #                         '交通是否有优惠','儿童费用是否有优惠']
        # version 3
        # self.table_item_list = ['入住时间','成人总人数(包括第三成人)','儿童人数','婴儿人数',\
        #                         '房价是否包含环境税','交通是否有优惠','儿童费用是否有优惠']
        # version 4
        self.table_item_list = ['入住时间','成人总人数(包括第三成人)','儿童人数','婴儿人数']
        self.table_row_num = 0
        self.span_id = []
        self.client_data = {}
        self.isgetTableWidgetUnitValueError = False

    def resize_screen(self, a, b, c, d):
        return int(a*self.screen_width/1920), int(b*self.screen_height/1080), \
                int(c*self.screen_width/1920), int(d*self.screen_height/1080)

    def retranslateUi(self, Form):
        self._translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(self._translate("Form", "Form"))
        self.pushButton.setText(self._translate("Form", "打开"))
        self.pushButton_2.setText(self._translate("Form", "计算"))
        self.label.setText(self._translate("Form", " 步骤(1): 请输入数据表路径"))
        self.label_2.setText(self._translate("Form", " 步骤(2): 请选择价单"))
        self.label_3.setText(self._translate("Form", " 步骤(3): 请输入客户信息数据"))
        self.label_4.setText(self._translate("Form", " 步骤(4): 输出"))
        self.label_5.setText(self._translate("Form", " 已选择的价单"))
        self.pushButton_3.setText(self._translate("Form", "重置"))
        self.pushButton_4.setText(self._translate("Form", "创建"))


class inputDialog(QDialog):
    def __init__(self, parent=None, rbStringList=''):
        QDialog.__init__(self, parent)
        self.font_size = 12
        self.screen_height, self.screen_width = self.center()
        self.resize(500*self.screen_width/1920, 100*self.screen_height/1080)
        self.setWindowTitle('The Dialog from Damon')
        self.chosedId = None
        self.rbStringList = rbStringList

        grid = QtWidgets.QGridLayout()
        self.label = QtWidgets.QLabel('出现多种选择方案，请做进一步选择（注意：编号顺序表示选择房型顺序）：',parent=self)
        left, top, width, height = self.resize_screen(20, 10, 200, 21)
        self.label.setGeometry(QtCore.QRect(left, top, width, height))
        self.font_size = max(6,int(self.font_size * width / 200))
        self.label.setStyleSheet("font: " + str(self.font_size) + "pt \"Adobe Arabic\";")
        self.label.setObjectName("label")
        grid.addWidget(self.label, 0, 0, 1, 1)
        
        for i in range(len(rbStringList)):
            exec("self.radioButton_" + str(i+1) + " = QtWidgets.QRadioButton('', parent=self)")
            exec("self.radioButton_" + str(i+1) + ".setAutoExclusive(True)")
            exec("self.radioButton_" + str(i+1) + ".setObjectName('radioButton_" + str(i+1) + "')")
            exec("grid.addWidget(self.radioButton_" + str(i+1) + ", " + str(i+1) + ", 0, 1, 1)")
            exec("self.radioButton_" + str(i+1) + ".setText('" + rbStringList[i] + "')")

        for Id in range(len(rbStringList)):
            exec("self.radioButton_" + str(Id+1) + ".toggled.connect(self.radioButtonAct)")

        buttonBox = QtWidgets.QDialogButtonBox(parent=self)
        buttonBox.setOrientation(QtCore.Qt.Horizontal)  # 设置为水平方向
        buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Ok)  # 确定和取消两个按钮

        buttonBox.accepted.connect(self.accept)  # 确定
        buttonBox.rejected.connect(self.reject)  # 取消

        layout = QtWidgets.QVBoxLayout()
        layout.addLayout(grid)

        spacerItem = QtWidgets.QSpacerItem(20, 48, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        layout.addItem(spacerItem)
        layout.addWidget(buttonBox)
        self.setLayout(layout)

    def center(self):
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        self.move((screen.width() - size.width()) / 2,  
        (screen.height() - size.height()) / 2)
        return int(screen.height()), int(screen.width())

    def resize_screen(self, a, b, c, d):
        return a*self.screen_width/1920, b*self.screen_height/1080, c*self.screen_width/1920, d*self.screen_height/1080

    def radioButtonAct(self):
        for Id in range(len(self.rbStringList)):
            exec("if self.radioButton_" + str(Id+1) + ".isChecked(): \
                    self.chosedId = Id")            
            

class inputDialog2(QDialog):
    def __init__(self, parent=None):
        QDialog.__init__(self, parent)
        self.font_size = 12
        self.screen_height, self.screen_width = self.center()
        self.resize(100*self.screen_width/1920, 80*self.screen_height/1080)
        self.setWindowTitle('The Dialog from Damon')

        grid = QtWidgets.QGridLayout()
        self.label1 = QtWidgets.QLabel('第一个儿童房费：',parent=self)
        left, top, width, height = self.resize_screen(20, 10, 200, 21)
        self.label1.setGeometry(QtCore.QRect(left, top, width, height))
        self.font_size = max(6,int(self.font_size * width / 200))
        self.label1.setStyleSheet("font: " + str(self.font_size) + "pt \"Adobe Arabic\";")
        self.label1.setObjectName("label1")
        grid.addWidget(self.label1, 0, 0, 1, 1)
        self.input_1 = QtWidgets.QLineEdit(parent=self)
        self.input_1.setText("0")
        grid.addWidget(self.input_1, 0, 1, 1, 1)

        self.label2 = QtWidgets.QLabel('第二个儿童房费：',parent=self)
        left, top, width, height = self.resize_screen(20, 10, 200, 21)
        self.label2.setGeometry(QtCore.QRect(left, top, width, height))
        self.font_size = max(6,int(self.font_size * width / 200))
        self.label2.setStyleSheet("font: " + str(self.font_size) + "pt \"Adobe Arabic\";")
        self.label2.setObjectName("label2")
        grid.addWidget(self.label2, 1, 0, 1, 1)
        self.input_2 = QtWidgets.QLineEdit(parent=self)
        self.input_2.setText("0")
        grid.addWidget(self.input_2, 1, 1, 1, 1)

        buttonBox = QtWidgets.QDialogButtonBox(parent=self)
        buttonBox.setOrientation(QtCore.Qt.Horizontal)  # 设置为水平方向
        buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Ok)  # 确定和取消两个按钮

        buttonBox.accepted.connect(self.accept)  # 确定
        buttonBox.rejected.connect(self.reject)  # 取消

        layout = QtWidgets.QVBoxLayout()
        layout.addLayout(grid)

        spacerItem = QtWidgets.QSpacerItem(20, 48, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        layout.addItem(spacerItem)
        layout.addWidget(buttonBox)
        self.setLayout(layout)

    def center(self):
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        self.move((screen.width() - size.width()) / 2,  
        (screen.height() - size.height()) / 2)
        return int(screen.height()), int(screen.width())

    def resize_screen(self, a, b, c, d):
        return a*self.screen_width/1920, b*self.screen_height/1080, c*self.screen_width/1920, d*self.screen_height/1080


class inputDialog3(QDialog):
    def __init__(self, parent=None):
        QDialog.__init__(self, parent)
        self.font_size = 12
        self.screen_height, self.screen_width = self.center()
        self.resize(1000*self.screen_width/1920, 600*self.screen_height/1080)
        self.setWindowTitle('请输入简要信息创建数据表')
        self.setColNum = 15 # 设置多少列的表格
        self.setRowNum = 30 # 设置多少行的表格
        self.HeaderLabelsList = ['起始时间','终止时间','房型全称','最大容纳人数(无附加费)','最大容纳人数(有附加费)',\
                                '打包房型全称','打包适用入住时间段','打包所含费用','交通组合方式','餐费类型',\
                                '附加费(强制费)收取起始时间','附加费(强制费)收取终止时间','提前预定天数以及相应的折扣',\
                                '入住天数以及对应折扣','住付优惠']

        grid = QtWidgets.QGridLayout()
        self.label1 = QtWidgets.QLabel('表格名称：',parent=self)
        self.label1.setStyleSheet("font: 8pt \"Adobe Arabic\";")
        self.label1.setObjectName("label1")
        grid.addWidget(self.label1, 0, 0, 1, 1)
        self.input = QtWidgets.QLineEdit(parent=self)
        grid.addWidget(self.input,1, 0, 1, 1)

        self.button = QtWidgets.QPushButton('Open CSV/XLSX', parent=self)
        self.button.clicked.connect(self.handleButton)
        grid.addWidget(self.button,1, 1, 1, 1)

        self.tableWidget1 = QtWidgets.QTableWidget(parent=self)
        left, top, width, height = self.resize_screen(20, 10, 200, 21)
        self.tableWidget1.setGeometry(QtCore.QRect(left, top, width, height)) #370, 40, 361, 661
        self.tableWidget1.setColumnCount(self.setColNum)
        self.tableWidget1.setRowCount(self.setRowNum)
        self.tableWidget1.setColumnWidth(0,70*width/200) #将第一列的单元宽度设置为210
        self.tableWidget1.setColumnWidth(1,70*width/200)
        self.tableWidget1.setColumnWidth(2,200*width/200) 
        self.tableWidget1.setColumnWidth(3,150*width/200) 
        self.tableWidget1.setColumnWidth(4,150*width/200) 
        self.tableWidget1.setColumnWidth(5,200*width/200) 
        self.tableWidget1.setColumnWidth(6,130*width/200) 
        self.tableWidget1.setColumnWidth(7,100*width/200) 
        self.tableWidget1.setColumnWidth(8,150*width/200) 
        self.tableWidget1.setColumnWidth(9,60*width/200) 
        self.tableWidget1.setColumnWidth(10,170*width/200) 
        self.tableWidget1.setColumnWidth(11,170*width/200) 
        self.tableWidget1.setColumnWidth(12,170*width/200) 
        self.tableWidget1.setColumnWidth(13,160*width/200) 
        self.tableWidget1.setColumnWidth(14,160*width/200) 
        self.tableWidget1.setHorizontalHeaderLabels(self.HeaderLabelsList)
        grid.addWidget(self.tableWidget1, 2, 0, 1, 1)

        # 填入基本的6种交通组合方式
        traffic_combination = ['单程内飞快艇和单程水飞','单程快艇和单程水飞','往返快艇','往返内飞','往返水飞','往返内飞快艇']
        for i,tc in enumerate(traffic_combination):
            newItem = QTableWidgetItem(tc)
            self.tableWidget1.setItem(i, 8, newItem)

        # 填入基本的5种餐费类型
        fee_type = ['BB','HB','FB','AI','PAI']
        for i,ft in enumerate(fee_type):
            newItem = QTableWidgetItem(ft)
            self.tableWidget1.setItem(i, 9, newItem)

        buttonBox = QtWidgets.QDialogButtonBox(parent=self)
        buttonBox.setOrientation(QtCore.Qt.Horizontal)  # 设置为水平方向
        buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Ok)  # 确定和取消两个按钮

        buttonBox.accepted.connect(self.accept)  # 确定
        buttonBox.rejected.connect(self.reject)  # 取消

        layout = QtWidgets.QVBoxLayout()
        layout.addLayout(grid)

        layout.addWidget(buttonBox)
        self.setLayout(layout)

    def handleButton(self):
        filters = (
            'Excel Files (*.xls *.xml *.xlsx *.xlsm)',
            'CSV files (*.csv *.txt)',
            )
        path, filter = QtWidgets.QFileDialog.getOpenFileName(
            self, 'Open File', '', ';;'.join(filters))
        HeaderLabelsDict = dict(zip(range(self.setColNum), self.HeaderLabelsList))
        if path:
            csv = filter.startswith('CSV')
            if csv:
                dataframe = read_csv(path)
            else:
                dataframe = read_excel(path)
            for Id_1 in HeaderLabelsDict:
                for Id_2,content in enumerate(dataframe[HeaderLabelsDict[Id_1]]):
                    if not isnull(content):
                        try:
                            content = int(content)
                        except:
                            pass
                        newItem = QTableWidgetItem(str(content))
                        self.tableWidget1.setItem(Id_2, Id_1, newItem)
                    else:
                        break

    def center(self):
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        # self.move((screen.width() - size.width()) / 2,  
        # (screen.height() - size.height()) / 2)
        return int(screen.height()), int(screen.width())

    def resize_screen(self, a, b, c, d):
        return a*self.screen_width/1920, b*self.screen_height/1080, c*self.screen_width/1920, d*self.screen_height/1080


class theWindowOfDamon(QtWidgets.QWidget,Ui_Form):
    def __init__(self):
        super(theWindowOfDamon,self).__init__() 
        screen_height, screen_width = self.center()
        self.setupUi(self, screen_height, screen_width)
        self.setWindowTitle("Damon")
        
        self.pushButton.clicked.connect(self.openDataFile)
        self.pushButton_4.clicked.connect(self.createTableForIsland)
        self.lineEdit.textChanged.connect(self.leAct) 
        self.lineEdit_2.textChanged.connect(self.le2Act) 
        self.pushButton_3.clicked.connect(self.reset)
        self.pushButton_2.clicked.connect(self.travel_fee_calc_func)

    def center(self):
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        # self.move((screen.width() - size.width()) / 2,  
        # (screen.height() - size.height()) / 2)
        return int(screen.height()), int(screen.width())

    def createTableForIsland(self):
        if self.datapath == '':
            self.msg("错误","请输入数据表文件夹路径！")
        else:
            dialog3 = inputDialog3()
            if dialog3.exec_():
                df = DataFrame(columns=dialog3.HeaderLabelsList)
                
                for row_Id in range(dialog3.setRowNum): #setRowNum=30
                    for col_Id, col_name in enumerate(dialog3.HeaderLabelsList):
                        if dialog3.tableWidget1.item(row_Id, col_Id) is not None:
                            df.loc[row_Id,col_name] = dialog3.tableWidget1.item(row_Id, col_Id).text()

                df_date = DataFrame({'起始日期':df['起始时间'], '终止日期':df['终止时间']})
                df_date = df_date.dropna(axis=0, how='all')
                df_house = DataFrame({'房型全称':df['房型全称']})
                df_house = df_house.dropna(axis=0, how='all')
                df_max_amount_without_fee = DataFrame({'最大容纳人数(无附加费)':df['最大容纳人数(无附加费)']})
                df_max_amount_without_fee = df_max_amount_without_fee.dropna(axis=0, how='all') # 2
                df_max_amount_with_fee = DataFrame({'最大容纳人数(有附加费)':df['最大容纳人数(有附加费)']})
                df_max_amount_with_fee = df_max_amount_with_fee.dropna(axis=0, how='all') # (2+2)/3
                df_house_db = DataFrame({'打包房型全称':df['打包房型全称']})
                df_house_db = df_house_db.dropna(axis=0, how='all')
                df_house_db_period = DataFrame({'打包适用入住时间段':df['打包适用入住时间段']})
                df_house_db_period = df_house_db_period.dropna(axis=0, how='all')
                df_house_db_fee = DataFrame({'打包所含费用':df['打包所含费用']})
                df_house_db_fee = df_house_db_fee.dropna(axis=0, how='all')
                df_traffic_combination = DataFrame({'交通组合方式':df['交通组合方式']})
                df_traffic_combination = df_traffic_combination.dropna(axis=0, how='all')
                df_fee_type = DataFrame({'餐费类型':df['餐费类型']})
                df_fee_type = df_fee_type.dropna(axis=0, how='all')
                df_date_fjf = DataFrame({'附加费(强制费)收取起始时间':df['附加费(强制费)收取起始时间'],\
                                         '附加费(强制费)收取终止时间':df['附加费(强制费)收取终止时间']})
                df_date_fjf = df_date_fjf.dropna(axis=0, how='all')
                df_discount_by_early_bird = DataFrame({'提前预定天数以及相应的折扣':df['提前预定天数以及相应的折扣']})
                df_discount_by_early_bird = df_discount_by_early_bird.dropna(axis=0, how='all')
                df_discount_by_stay_days = DataFrame({'入住天数以及对应折扣':df['入住天数以及对应折扣']})
                df_discount_by_stay_days = df_discount_by_stay_days.dropna(axis=0, how='all')
                df_discount = DataFrame({'住付优惠':df['住付优惠']})
                df_discount = df_discount.dropna(axis=0, how='all')

                df_output_all = {}
                # 房型价格表
                df_output = DataFrame(columns=['房型','起始日期','终止日期','最大容纳人数(无附加费)', \
                                                '最大容纳人数(有附加费)','单价'])
                for k,house_type in enumerate(df_house['房型全称']):
                    sheet_1 = DataFrame(columns=['房型','起始日期','终止日期','最大容纳人数(无附加费)', \
                                                '最大容纳人数(有附加费)','单价'])
                    min_num = 1
                    max_num = int(df_max_amount_without_fee['最大容纳人数(无附加费)'].iloc[k])
                    for i in range(df_date.shape[0]):
                        for j, num in enumerate(range(min_num,max_num+1)):
                            sheet_1.loc[i*len(range(min_num,max_num+1))+j,'最大容纳人数(无附加费)'] = num
                            sheet_1.loc[i*len(range(min_num,max_num+1))+j,'起始日期'] = df['起始时间'].iloc[i]
                            sheet_1.loc[i*len(range(min_num,max_num+1))+j,'终止日期'] = df['终止时间'].iloc[i]
                            sheet_1.loc[i*len(range(min_num,max_num+1))+j,'房型'] = house_type
                            sheet_1.loc[i*len(range(min_num,max_num+1))+j,'最大容纳人数(有附加费)'] = df_max_amount_with_fee['最大容纳人数(有附加费)'].iloc[k]
                            sheet_1.loc[i*len(range(min_num,max_num+1))+j,'单价'] = 0
                    df_output = df_output.append(sheet_1)
                df_output.set_index(list(df_output.keys())[0], inplace=True)
                df_output_all.update({'房型价格表':df_output})

                # 房型简称表
                df_output = DataFrame(columns=['房型全称','房型简称'])
                for i,house_type in enumerate(df_house['房型全称']):
                    df_output.loc[i,'房型全称'] = house_type
                    df_output.loc[i,'房型简称'] = '房型' + str(i+1)
                df_output.set_index(list(df_output.keys())[0], inplace=True)
                df_output_all.update({'房型简称':df_output})

                # 房型打包价格表
                col_name = ['打包房型全称']
                col_name.extend(list(df_fee_type['餐费类型']))
                df_output = DataFrame(columns=col_name)
                try:
                    for i,house_type in enumerate(df_house_db['打包房型全称']):
                        df_output.loc[i,'打包房型全称'] = house_type
                        for ft in list(df_fee_type['餐费类型']):
                            df_output.loc[i,ft] = 0
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'房型打包价格表':df_output})
                except:
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'房型打包价格表':df_output})

                # 打包价格下，第三成人费用表
                col_name = ['类型','打包房型全称','打包适用入住时间段','打包所含费用']; col_name_1 = []
                col_name.extend(list(df_fee_type['餐费类型'])); col_name_1.extend(list(df_fee_type['餐费类型']))
                col_name.extend(list(df_traffic_combination['交通组合方式'])); col_name_1.extend(list(df_traffic_combination['交通组合方式']))
                df_output = DataFrame(columns=col_name)
                try:
                    for i,t in enumerate(['成人','儿童','婴儿']):
                        sheet_2 = DataFrame(columns=col_name)
                        for j in range(df_house_db.shape[0]):
                            sheet_2.loc[i*df_date_fjf.shape[0]+j,'类型'] = t
                            sheet_2.loc[i*df_date_fjf.shape[0]+j,'打包房型全称'] = df_house_db['打包房型全称'].iloc[j]
                            sheet_2.loc[i*df_date_fjf.shape[0]+j,'打包适用入住时间段'] = df_house_db_period['打包适用入住时间段'].iloc[j]
                            sheet_2.loc[i*df_date_fjf.shape[0]+j,'打包所含费用'] = df_house_db_fee['打包所含费用'].iloc[j]
                            for name in col_name_1:
                                sheet_2.loc[i*df_date_fjf.shape[0]+j,name] = 0
                        df_output = df_output.append(sheet_2)
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'第三人费用表(打包情况)':df_output})
                except:
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'第三人费用表(打包情况)':df_output})

                # 餐费价格表, 在合同价格计算时，用于计算升级餐型
                col_name = ['分类']
                col_name.extend(list(df_fee_type['餐费类型']))
                df_output = DataFrame(columns=col_name)
                try:
                    for i,t in enumerate(['成人','儿童','婴儿']):
                        df_output.loc[i,'分类'] = t
                        for ft in list(df_fee_type['餐费类型']):
                            df_output.loc[i,ft] = 0
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'餐费价格表':df_output})
                except:
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'餐费价格表':df_output})

                # 交通组合价格表
                col_name = ['分类']
                col_name.extend(list(df_traffic_combination['交通组合方式']))
                df_output = DataFrame(columns=col_name)
                try:
                    for i,t in enumerate(['成人','儿童','婴儿']):
                        df_output.loc[i,'分类'] = t
                        for ft in list(df_traffic_combination['交通组合方式']):
                            df_output.loc[i,ft] = 0
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'交通组合价格表':df_output})
                except:
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'交通组合价格表':df_output})

                # 节假日附加费用表和强制费用表
                col_name = ['分类','起始日期','终止日期']
                col_name.extend(list(df_fee_type['餐费类型']))
                col_name.extend(['附加费'])
                df_output = DataFrame(columns=col_name)
                try:
                    for i,t in enumerate(['成人','儿童','婴儿']): #婴儿没有附加费或者强制费
                        sheet_3 = DataFrame(columns=col_name)
                        for j in range(df_date_fjf.shape[0]):
                            sheet_3.loc[i*df_date_fjf.shape[0]+j,'起始日期'] = df_date_fjf['附加费(强制费)收取起始时间'].iloc[j]
                            sheet_3.loc[i*df_date_fjf.shape[0]+j,'终止日期'] = df_date_fjf['附加费(强制费)收取终止时间'].iloc[j]
                            sheet_3.loc[i*df_date_fjf.shape[0]+j,'BB'] = 0
                            sheet_3.loc[i*df_date_fjf.shape[0]+j,'HB'] = 0
                            sheet_3.loc[i*df_date_fjf.shape[0]+j,'FB'] = 0
                            sheet_3.loc[i*df_date_fjf.shape[0]+j,'AI'] = 0
                            sheet_3.loc[i*df_date_fjf.shape[0]+j,'PAI'] = 0
                            sheet_3.loc[i*df_date_fjf.shape[0]+j,'附加费'] = 0
                            sheet_3.loc[i*df_date_fjf.shape[0]+j,'分类'] = t
                        df_output = df_output.append(sheet_3)
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'节假日附加费和强制费用表':df_output})
                except:
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'节假日附加费和强制费用表':df_output})

                # 住付优惠表，住付优惠不分房型
                col_name = ['适用时间段','住付优惠']
                df_output = DataFrame(columns=col_name)
                try:
                    for i,row in enumerate(df_discount['住付优惠']):
                        sheet_4 = DataFrame(columns=col_name)
                        period, discount = row.split("#")
                        discount_list = discount.split("/")
                        for j, d in enumerate(discount_list):
                            sheet_4.loc[i*len(discount_list)+j,'适用时间段'] = period
                            sheet_4.loc[i*len(discount_list)+j,'住付优惠'] = d
                        df_output = df_output.append(sheet_4)
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'住付优惠表':df_output})
                except:
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'住付优惠表':df_output})

                # 提前预定天数以及相应的折扣表Early Bird Discount
                col_name = ['适用时间段','提前预定天数','折扣']
                df_output = DataFrame(columns=col_name)
                try:
                    for i, row in enumerate(df_discount_by_early_bird['提前预定天数以及相应的折扣']):
                        row_list = row.split("#")
                        for j, name in enumerate(col_name):
                            df_output.loc[i,name] = row_list[j]
                    df_output['能混用的优惠方式'] = ''
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'EBD折扣表':df_output})
                except:
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'EBD折扣表':df_output})

                # 入住天数以及对应折扣表Long Stay Discount
                col_name = ['适用时间段','入住天数','折扣']
                df_output = DataFrame(columns=col_name)
                try:
                    for i, row in enumerate(df_discount_by_stay_days['入住天数以及对应折扣']):
                        row_list = row.split("#")
                        for j, name in enumerate(col_name):
                            df_output.loc[i,name] = row_list[j]
                    df_output['能混用的优惠方式'] = ''
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'LSD折扣表':df_output})
                except:
                    df_output.set_index(list(df_output.keys())[0], inplace=True)
                    df_output_all.update({'LSD折扣表':df_output})

                # 其他单价表
                df_output = DataFrame(columns=['名称','数值'])
                df_output.loc[0,'名称'] = '混住费'; df_output.loc[0,'数值'] = '0'
                df_output.loc[1,'名称'] = 'EBD和LSD不可混用的日期'; df_output.loc[1,'数值'] = ''
                df_output.loc[2,'名称'] = 'LSD和EBD是否适用于第三人和餐型'; df_output.loc[2,'数值'] = '否'
                df_output.loc[3,'名称'] = '房价是否包含环境税'; df_output.loc[3,'数值'] = '否'
                df_output.loc[4,'名称'] = '交通是否有优惠'; df_output.loc[4,'数值'] = '否'
                df_output.loc[5,'名称'] = '儿童费用是否有优惠'; df_output.loc[5,'数值'] = '否'
                df_output.loc[6,'名称'] = '免费夜强制费'; df_output.loc[6,'数值'] = '0'
                df_output.loc[7,'名称'] = '免费夜固定房费价格'; df_output.loc[7,'数值'] = '0'
                df_output.loc[8,'名称'] = '第三成人免费夜固定房费价格'; df_output.loc[8,'数值'] = '0'
                df_output.loc[9,'名称'] = '成人环境税'; df_output.loc[9,'数值'] = '6'
                df_output.loc[10,'名称'] = '儿童环境税'; df_output.loc[10,'数值'] = '6'
                df_output.loc[11,'名称'] = '婴儿环境税'; df_output.loc[11,'数值'] = '6'
                for i in range(12,12+df_house.shape[0]):
                    df_output.loc[i,'名称'] = df_house['房型全称'].iloc[i-12] + '第三成人费用(合同)'
                    df_output.loc[i,'数值'] = 0
                df_output.set_index(list(df_output.keys())[0], inplace=True)
                df_output_all.update({'其他单价表':df_output})

                if (dialog3.input.text() is None) or (dialog3.input.text() == ''):
                    self.msg("错误","请输入数据表名称！")
                else:
                    writer = ExcelWriter(os.path.join(self.datapath,dialog3.input.text()+'.xlsx'))
                    for Id, item in enumerate(df_output_all):
                        df_output_all[item].to_excel(writer, sheet_name=item)
                    writer.save()
                    writer.close()
                    reply = QMessageBox.information(self,'提示','已创建'+str(dialog3.input.text())+'.xlsx数据表！')

    def crtRBFoSrchFile(self):
        # create radio-button for searched file
        if self.datapath == '':
            self.msg("错误","请输入数据表文件夹路径！")
        else:
            self.path_setchecked_dict.update({self.datapath:set()})
            if self.searchFile != '':
                if self.verticalLayout.count() != 0:
                    for i in range(self.verticalLayout.count()-1):
                        exec("sip.delete(self.radioButton_" + str(i+1) + ")")
                    self.verticalLayout.removeItem(self.spacerItem1)
                Id = 0
                self.search_table_name_list = []
                for table_name in self.table_name_list:
                    if self.searchFile in table_name:
                        exec("self.radioButton_" + str(Id+1) + "= QtWidgets.QRadioButton(self.verticalWidget)")
                        exec("self.radioButton_" + str(Id+1) + ".setAutoExclusive(False)")
                        exec("self.radioButton_" + str(Id+1) + ".setObjectName('radioButton_" + str(Id+1) + "')")
                        exec("self.verticalLayout.addWidget(self.radioButton_" + str(Id+1) + ")")
                        exec("self.radioButton_" + str(Id+1) + ".setText(self._translate('Form', '" + table_name + "'))")
                        Id += 1
                        self.search_table_name_list.append(table_name)
                self.spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, \
                                   QtWidgets.QSizePolicy.Expanding)
                self.verticalLayout.addItem(self.spacerItem1)
                self.radioButtonNum = Id 

    def crtRB(self):
        # create radio-button
        if self.verticalLayout.count() != 0:
            for i in range(self.verticalLayout.count()-1):
                exec("sip.delete(self.radioButton_" + str(i+1) + ")")
            self.verticalLayout.removeItem(self.spacerItem1)
        try:
            if self.datapath != '':
                self.path_setchecked_dict.update({self.datapath:set()})
                self.table_name_list = [name.split('.')[0] for name in os.listdir(self.datapath) if '.xlsx' in name]
                for Id, name in enumerate(self.table_name_list):
                    exec("self.radioButton_" + str(Id+1) + "= QtWidgets.QRadioButton(self.verticalWidget)")
                    exec("self.radioButton_" + str(Id+1) + ".setAutoExclusive(False)")
                    exec("self.radioButton_" + str(Id+1) + ".setObjectName('radioButton_" + str(Id+1) + "')")
                    exec("self.verticalLayout.addWidget(self.radioButton_" + str(Id+1) + ")")
                    exec("self.radioButton_" + str(Id+1) + ".setText(self._translate('Form', '" + name + "'))")
                self.spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, \
                                   QtWidgets.QSizePolicy.Expanding)
                self.verticalLayout.addItem(self.spacerItem1)
                self.radioButtonNum = Id + 1
                self.search_table_name_list = self.table_name_list
        except Exception as e:
                self.msg("错误",str(e))

    def reset(self):
        # reset radio-button, table widget and output
        self.history_tb2_string_output = []
        self.path_setchecked_dict = {}
        self.textBrowser_2.clear()
        for Id in range(self.radioButtonNum):
            exec("self.radioButton_" + str(Id+1) + ".setChecked(False)")

    def leAct(self):
        try:
            if self.lineEdit.text() != '':
                self.datapath = self.lineEdit.text()
                self.tableList = []
                self.crtRB()
                for Id in range(self.radioButtonNum):
                    exec("self.radioButton_" + str(Id+1) + ".toggled.connect(self.radioButtonAct)")
        except:
            pass

    def le2Act(self):
        self.history_tb2_string_output.extend(self.tb2_string_output)
        self.history_tb2_string_output = list(set(self.history_tb2_string_output))
        if self.lineEdit_2.text() != '':
            self.searchFile = self.lineEdit_2.text()
            self.crtRBFoSrchFile()
            for Id in range(self.radioButtonNum):
                exec("if self.radioButton_" + str(Id+1) + ".text() in self.history_tb2_string_output:\
                         self.radioButton_" + str(Id+1) + ".setChecked(True)")
                exec("self.radioButton_" + str(Id+1) + ".toggled.connect(self.radioButtonAct)")
            self.textBrowser_2.clear()
            self.textBrowser_2.append(', '.join(self.history_tb2_string_output))
        else:
            self.crtRB()
            for Id in range(self.radioButtonNum):
                exec("if self.radioButton_" + str(Id+1) + ".text() in self.history_tb2_string_output:\
                         self.radioButton_" + str(Id+1) + ".setChecked(True)")
                exec("self.radioButton_" + str(Id+1) + ".toggled.connect(self.radioButtonAct)")
            self.textBrowser_2.clear()
            self.textBrowser_2.append(', '.join(self.history_tb2_string_output))

    def radioButtonAct(self):
        self.textBrowser_2.clear()
        self.path_setchecked_dict.update({self.datapath:set()})
        notsetCheckedSet = set()
        for Id in range(self.radioButtonNum):
            exec("if self.radioButton_" + str(Id+1) + ".isChecked(): \
                    self.path_setchecked_dict[self.datapath].add(self.search_table_name_list[Id])")
            exec("if not self.radioButton_" + str(Id+1) + ".isChecked(): \
                    notsetCheckedSet.add(self.search_table_name_list[Id])")
        self.tb2_string_output = [] 
        for idx, path in enumerate(self.path_setchecked_dict):
            self.tb2_string_output.extend(list(self.path_setchecked_dict[path])) 
        self.history_tb2_string_output.extend(self.tb2_string_output)
        self.history_tb2_string_output = list(set(self.history_tb2_string_output) - notsetCheckedSet)
        self.textBrowser_2.append(', '.join(self.history_tb2_string_output))
        # print(', '.join(self.history_tb2_string_output))

        tmp_table_list = []
        tmp_table_set = {}
        if self.history_tb2_string_output:
            for table_name in self.history_tb2_string_output:
                try:
                    df = read_excel(os.path.join(self.datapath,table_name+'.xlsx'), sheet_name='房型简称')
                    tmp_table_list.extend(list(df['房型全称']))
                    tmp_table_set.update({table_name:list(df['房型全称'])})
                except Exception as e:
                    self.msg("错误","该路径不存在" + table_name + ".xlsx文件！")

        self.tableWidget.clearContents()
        self.tableWidget.clearSpans()
        self.span_id = []
        self.client_data = {}
        if tmp_table_set:
            for Id_1, item in enumerate(self.table_item_list):
                newItem = QTableWidgetItem(item)
                self.tableWidget.setItem(Id_1, 0, newItem)
                if item == '是否有第三成人':
                    self.comBox1 = QtWidgets.QComboBox()
                    self.comBox1.addItems(['否','是'])
                    self.comBox1.setStyleSheet('QComboBox{margin:3px}')
                    self.tableWidget.setCellWidget(Id_1,1,self.comBox1)
                elif item == '房价是否包含环境税':
                    self.comBox2 = QtWidgets.QComboBox()
                    self.comBox2.addItems(['否','是'])
                    self.comBox2.setStyleSheet('QComboBox{margin:3px}')
                    self.tableWidget.setCellWidget(Id_1,1,self.comBox2)
                elif item == '交通是否有优惠':
                    self.comBox3 = QtWidgets.QComboBox()
                    self.comBox3.addItems(['否','是'])
                    self.comBox3.setStyleSheet('QComboBox{margin:3px}')
                    self.tableWidget.setCellWidget(Id_1,1,self.comBox3)
                elif item == '儿童费用是否有优惠':
                    self.comBox4 = QtWidgets.QComboBox()
                    self.comBox4.addItems(['否','是'])
                    self.comBox4.setStyleSheet('QComboBox{margin:3px}')
                    self.tableWidget.setCellWidget(Id_1,1,self.comBox4)

            start_id = Id_1+1
            for table_name in tmp_table_set:
                self.tableWidget.setSpan(start_id,0,1,2)
                newItem = QTableWidgetItem("价单："+table_name)
                newItem.setForeground(QtGui.QBrush(QtGui.QColor(255, 0, 0)))
                newItem.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                self.tableWidget.setItem(start_id,0,newItem)
                self.span_id.append(start_id)

                tmp = list(set(tmp_table_set[table_name]))
                for Id_2, item in enumerate(tmp):
                    newItem = QTableWidgetItem(item)
                    self.tableWidget.setItem(Id_2+start_id+1, 0, newItem)
                start_id += Id_2+2

            newItem = QTableWidgetItem("其他附加费：")
            newItem.setForeground(QtGui.QBrush(QtGui.QColor(255, 0, 0)))
            self.tableWidget.setItem(start_id, 0, newItem)
            self.table_row_num = start_id + 1
        else:
            self.table_row_num = 0

    def getTableWidgetUnitValue(self):
        self.isgetTableWidgetUnitValueError = False
        for Id in range(self.table_row_num):
            if Id not in self.span_id:
                if self.tableWidget.item(Id, 1) is None:
                    if self.tableWidget.item(Id, 0).text() == '是否有第三成人':
                        self.client_data.update({self.tableWidget.item(Id, 0).text():self.comBox1.currentText()})
                    elif self.tableWidget.item(Id, 0).text() == '房价是否包含环境税':
                        self.client_data.update({self.tableWidget.item(Id, 0).text():self.comBox2.currentText()})
                    elif self.tableWidget.item(Id, 0).text() == '交通是否有优惠':
                        self.client_data.update({self.tableWidget.item(Id, 0).text():self.comBox3.currentText()})
                    elif self.tableWidget.item(Id, 0).text() == '儿童费用是否有优惠':
                        self.client_data.update({self.tableWidget.item(Id, 0).text():self.comBox4.currentText()})
                    elif self.tableWidget.item(Id, 0).text() == '优惠折扣':
                        self.client_data.update({self.tableWidget.item(Id, 0).text():1})
                    else:
                        self.client_data.update({self.tableWidget.item(Id, 0).text():0})
                else:
                    if self.tableWidget.item(Id, 1).text() == '':
                        if self.tableWidget.item(Id, 0).text() == '优惠折扣':
                            self.client_data.update({self.tableWidget.item(Id, 0).text():1})
                        else:
                            self.client_data.update({self.tableWidget.item(Id, 0).text():0})
                    else:
                        try:
                            if self.tableWidget.item(Id, 0).text() == '优惠折扣':
                                self.client_data.update({self.tableWidget.item(Id, 0).text():\
                                                        float(self.tableWidget.item(Id, 1).text())})
                            else:
                                self.client_data.update({self.tableWidget.item(Id, 0).text():\
                                                        int(self.tableWidget.item(Id, 1).text())})
                        except Exception as e:
                            self.isgetTableWidgetUnitValueError = True
                            self.msg("错误",self.tableWidget.item(Id, 0).text()+"的输入格式不对！")

        if (self.client_data['优惠折扣'] > 1) or (self.client_data['优惠折扣'] < 0):
            self.isgetTableWidgetUnitValueError = True
            self.msg("错误","优惠折扣不能大于1或小于0！")
        # print(self.client_data)

    def openDataFile(self):
        filename = QFileDialog.getExistingDirectory(self, "选取文件夹", os.getcwd())  
        if filename:
            if self.lineEdit.text() == '':
                self.lineEdit.setText(filename)
            else:
                self.leAct()

    def msg(self,string1,string2):
        reply = QMessageBox.critical(self,string1,string2)

    def get_specific_dates_list(self,start_date,days=0):
        output = []
        start_date = str(start_date)
        start_date = date(int(start_date[0:4]), int(start_date[4:6]), int(start_date[6:]))
        for day_plus in range(days):
            end_date = start_date + timedelta(days=day_plus)
            end_date = end_date.year * 10000 + end_date.month * 100 + end_date.day
            output.append(end_date)
        return output

    def is_number(self,uchar):
        if uchar >= u'\u0030' and uchar <= u'\u0039':
            return True
        else:
            return False

    def is_date_valid(self,Date):
        try:
            date(int(Date[0:4]), int(Date[4:6]), int(Date[6:]))
            return True
        except Exception as e:
            return False

    def closeEvent(self, event):
        reply = QMessageBox.question(self, "信息", "确认要退出?", QMessageBox.Yes |
                                     QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def single_island_calc(self):
        # 导入所有表格
        sheet_names = ExcelFile(os.path.join(self.datapath,self.history_tb2_string_output[0] \
                                +'.xlsx')).sheet_names  # see all sheet names
        table_set = {}
        for sn in sheet_names:
            table_set.update({sn:read_excel(os.path.join(self.datapath,self.history_tb2_string_output[0] \
                            +'.xlsx'), sheet_name=sn)})

        # 获取单价表数据
        cal_unit = table_set['其他单价表']['单位'].iloc[0]
        adult_traffic_price = table_set['其他单价表'][table_set['其他单价表']['名称'] == '成人交通费']['价格'].values[0]
        child_traffic_price = table_set['其他单价表'][table_set['其他单价表']['名称'] == '儿童交通费']['价格'].values[0]
        adult_env_tax_price = table_set['其他单价表'][table_set['其他单价表']['名称'] == '成人环境税']['价格'].values[0]
        child_env_tax_price = table_set['其他单价表'][table_set['其他单价表']['名称'] == '儿童环境税']['价格'].values[0]
        baby_env_tax_price = table_set['其他单价表'][table_set['其他单价表']['名称'] == '婴儿环境税']['价格'].values[0]
        third_person_set = {}
        for row in table_set['其他单价表'][table_set['其他单价表']['名称'].str.contains('第三人费用')==True].values:
            third_person_set.update({row[0].split('第三人费用')[0]:row[2]})

        adult_amount = self.client_data['成人总人数(包括第三成人)']
        if self.client_data['是否有第三成人'] == '是':
            adult_amount -= 1 
        child_amount = self.client_data['儿童人数']
        baby_amount = self.client_data['婴儿人数']

        # 统计入住天数
        check_in_nights = 0; house_whole_name_list = []; days_set = {}; isWrong = False; third_person_cost_set = {}
        for i in range(self.span_id[0]+1,self.table_row_num-1):
            check_in_nights += self.client_data[self.tableWidget.item(i, 0).text()]
            if self.client_data[self.tableWidget.item(i, 0).text()] != 0:
                selectby_house_type = table_set['房型价格表'][table_set['房型价格表']['房型'] == self.tableWidget.item(i, 0).text()]
                tmp_list = [int(t) for t in set(selectby_house_type['人数'])]
                if adult_amount > max(tmp_list):
                    self.msg("错误","成人人数超过了 "+self.tableWidget.item(i, 0).text()+" 房型的最大容纳人数，房型价格表中无该条数据！")
                    isWrong = True
                    break
                elif adult_amount < min(tmp_list):
                    self.msg("错误","房型 "+self.tableWidget.item(i, 0).text()+" 无"+str(self.client_data['成人总人数(包括第三成人)'])+"个人的数据记录 ！")
                    isWrong = True
                    break
                else:
                    house_whole_name_list.append(self.tableWidget.item(i, 0).text())
                    days_set.update({self.tableWidget.item(i, 0).text():self.client_data[self.tableWidget.item(i, 0).text()]})
                    third_person_cost_set.update({self.tableWidget.item(i, 0).text():third_person_set[self.tableWidget.item(i, 0).text()]})

        if not isWrong:
            if check_in_nights == 0:
                self.msg("错误","请输入特定房型的入住天数！")
            else:
                start_date = str(self.client_data['入住时间'])
                start_date = date(int(start_date[0:4]), int(start_date[4:6]), int(start_date[6:]))
                end_date = start_date + timedelta(days=check_in_nights)
                end_date = end_date.year * 10000 + end_date.month * 100 + end_date.day
                start_date = start_date.year * 10000 + start_date.month * 100 + start_date.day

                date_set = set()
                for Id, _date in enumerate(table_set['房型价格表']['起始日期']):
                    date_set.add(str(_date)+' '+str(table_set['房型价格表']['终止日期'].iloc[Id]))

                K_start_date = 0; K_end_date = 0
                for date_interval in date_set:
                    a, b = date_interval.split(' ')
                    if start_date in Interval(int(a), int(b)):
                        K_start_date += 1
                    if end_date in Interval(int(a), int(b)):
                        K_end_date += 1
                if K_start_date == 0 and K_end_date != 0:
                    self.msg("错误","起始日期不在数据表设置的日期范围内！")
                elif K_start_date != 0 and K_end_date == 0:
                    self.msg("错误","结束日期超出数据表设置的日期范围！")
                elif K_start_date == 0 and K_end_date == 0:
                    self.msg("错误","起始日期和结束日期均不在数据表设置的日期范围！")
                else:
                    try:
                        price_set = {}
                        for house_type in house_whole_name_list:
                            selectby_house_type = table_set['房型价格表'][table_set['房型价格表']['房型'] == house_type]
                            selectby_house_type = selectby_house_type.reset_index(drop=True)
                            record_id_list = []
                            for i in range(len(selectby_house_type)):
                                if start_date in Interval(selectby_house_type['起始日期'][i], selectby_house_type['终止日期'][i]) and \
                                    end_date in Interval(selectby_house_type['起始日期'][i], selectby_house_type['终止日期'][i]):
                                    record_id_list.append(i)
                            selectby_house_type = selectby_house_type.iloc[record_id_list]
                            if len(selectby_house_type) != 1:
                                price_set.update({house_type:selectby_house_type[selectby_house_type['人数'] == adult_amount]['单价'].values[0]})
                            else:
                                price_set.update({house_type:selectby_house_type['单价'].values[0]})

                        whole_cost = 0

                        # 计算正常成人房费
                        house_cost = 0
                        for Id, (house_type, house_price) in enumerate(price_set.items()):
                            house_cost += house_price * days_set[house_type]

                        # 计算第三成人房费
                        for house_type in third_person_cost_set:
                            house_cost += third_person_cost_set[house_type] * days_set[house_type]

                        # 减去正常成人免费夜费用
                        if self.client_data['免房费天数'] != 0:
                            if self.client_data['免费夜固定价格'] == 0:
                                house_cost -= self.client_data['免房费天数'] * price_set[min(price_set,key=price_set.get)]
                            else:
                                house_cost -= self.client_data['免房费天数'] * self.client_data['免费夜固定价格'] 

                        # 计算总的环境税（包括成人/儿童/婴儿）
                        env_tax_cost = (adult_env_tax_price * (adult_amount+1) + child_env_tax_price * child_amount + baby_env_tax_price * baby_amount) * check_in_nights

                        # 减去第三成人免费夜费用
                        if self.client_data['免房费天数'] != 0:
                            if self.client_data['第三成人免费夜固定价格'] == 0:
                                house_cost -= self.client_data['免房费天数'] * price_set[min(price_set,key=price_set.get)]
                            else:
                                house_cost -= self.client_data['免房费天数'] * self.client_data['第三成人免费夜固定价格'] 

                        if self.client_data['是否有第三成人'] == '是':
                            if self.client_data['房价是否包含环境税'] == '是':
                                house_cost -= adult_env_tax_price * check_in_nights * (adult_amount+1) # 成人环境税
                            traffic_cost = adult_traffic_price * (adult_amount+1) + child_traffic_price * child_amount # 计算总交通费
                            qiangzhi_dinner = self.client_data['强制餐单价'] * (adult_amount+1+child_amount) # 强制餐，包括成人和儿童
                        else:
                            if self.client_data['房价是否包含环境税'] == '是':
                                house_cost -= adult_env_tax_price * check_in_nights * adult_amount # 成人环境税
                            traffic_cost = adult_traffic_price * adult_amount + child_traffic_price * child_amount # 计算总交通费
                            qiangzhi_dinner = self.client_data['强制餐单价'] * (adult_amount+child_amount) # 强制餐，包括成人和儿童
                        

                        self.child_cost = self.child_cost * check_in_nights
                        # 计算与折扣相乘的值
                        if self.client_data['交通是否有优惠'] == '是':
                            if self.client_data['儿童费用是否有优惠'] == '是':
                                whole_cost = (house_cost + traffic_cost + self.child_cost) * self.client_data['优惠折扣'] + env_tax_cost + self.client_data['强制费单价'] \
                                            * self.client_data['强制费收取天数'] + qiangzhi_dinner + self.client_data['其他附加费：']  
                                if  self.client_data['房价是否包含环境税'] == '是':
                                    self.string_outprint += " * 实际费用 = ( 房费(已扣环境税,剩" + str(house_cost) + ") + 交通费(" + str(traffic_cost) + ") + 儿童费用(" + str(self.child_cost) + \
                                                            ") ) * 优惠折扣 + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                            ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                else:
                                    self.string_outprint += " * 实际费用 = ( 房费(" + str(house_cost) + ") + 交通费(" + str(traffic_cost) + ") + 儿童费用(" + str(self.child_cost) + \
                                                            ") ) * 优惠折扣 + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                            ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                            else:
                                whole_cost = (house_cost + traffic_cost) * self.client_data['优惠折扣'] + self.child_cost + env_tax_cost + self.client_data['强制费单价'] \
                                            * self.client_data['强制费收取天数'] + qiangzhi_dinner + self.client_data['其他附加费：']   
                                if  self.client_data['房价是否包含环境税'] == '是':
                                    self.string_outprint += " * 实际费用 = ( 房费(已扣环境税,剩" + str(house_cost) + ") + 交通费(" + str(traffic_cost) + ") ) * 优惠折扣 + 儿童费用(" + str(self.child_cost) + \
                                                            ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                            ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                else:
                                    self.string_outprint += " * 实际费用 = ( 房费(" + str(house_cost) + ") + 交通费(" + str(traffic_cost) + ") ) * 优惠折扣 + 儿童费用(" + str(self.child_cost) + \
                                                            ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                            ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                        else:
                            if self.client_data['儿童费用是否有优惠'] == '是':
                                whole_cost = (house_cost + self.child_cost) * self.client_data['优惠折扣'] + traffic_cost + env_tax_cost + self.client_data['强制费单价'] \
                                            * self.client_data['强制费收取天数'] + qiangzhi_dinner + self.client_data['其他附加费：']   
                                if  self.client_data['房价是否包含环境税'] == '是':
                                    self.string_outprint += " * 实际费用 = ( 房费(已扣环境税,剩" + str(house_cost) + ") + 儿童费用(" + str(self.child_cost) + ") ) * 优惠折扣 + 交通费(" + \
                                                            str(traffic_cost) + ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                            ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                else:
                                    self.string_outprint += " * 实际费用 = ( 房费(" + str(house_cost) + ") + 儿童费用(" + str(self.child_cost) + ") ) * 优惠折扣 + 交通费(" + \
                                                            str(traffic_cost) + ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                            ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit          
                            else:
                                whole_cost = house_cost * self.client_data['优惠折扣'] + self.child_cost + traffic_cost + env_tax_cost + self.client_data['强制费单价'] \
                                            * self.client_data['强制费收取天数'] + qiangzhi_dinner + self.client_data['其他附加费：']        
                                if  self.client_data['房价是否包含环境税'] == '是':
                                    self.string_outprint += " * 实际费用 = 房费(已扣环境税,剩" + str(house_cost) + ") * 优惠折扣 + 儿童费用(" + str(self.child_cost) + ") + 交通费(" + \
                                                            str(traffic_cost) + ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                            ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                else:
                                    self.string_outprint += " * 实际费用 = 房费(" + str(house_cost) + ") * 优惠折扣 + 儿童费用(" + str(self.child_cost) + ") + 交通费(" + \
                                                            str(traffic_cost) + ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                            ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit

                    except Exception as e:
                        # 判断是否有多个房型
                        if len(house_whole_name_list) == 1:
                            price_list = []
                            start_date_1 = self.client_data['入住时间']
                            house_type = house_whole_name_list[0]
                            selectby_house_type = table_set['房型价格表'][(table_set['房型价格表']['房型'] == house_type) & \
                                                                                    (table_set['房型价格表']['人数'] == self.client_data['成人总人数(包括第三成人)'])]
                            selectby_house_type = selectby_house_type.reset_index(drop=True)
                            dates_list = self.get_specific_dates_list(start_date_1,days_set[house_type])
                            sum_price = 0
                            for d in dates_list:
                                price = selectby_house_type[(d >= selectby_house_type['起始日期']) & \
                                                          (d <= selectby_house_type['终止日期'])]['单价'].values[0]
                                sum_price += price
                                price_list.append(price)
                            start_date_1 = self.get_specific_dates_list(dates_list[-1],2)[1]
                            price_set.update({house_type:sum_price})

                            whole_cost = 0

                            # 计算正常成人房费
                            house_cost = 0
                            for Id, (house_type, house_price) in enumerate(price_set.items()):
                                house_cost += house_price * days_set[house_type]

                            # 计算第三成人房费
                            for house_type in third_person_cost_set:
                                house_cost += third_person_cost_set[house_type] * days_set[house_type]

                            # 减去正常成人免费夜费用
                            if self.client_data['免房费天数'] != 0:
                                if self.client_data['免费夜固定价格'] == 0:
                                    house_cost -= self.client_data['免房费天数'] * price_set[min(price_set,key=price_set.get)]
                                else:
                                    house_cost -= self.client_data['免房费天数'] * self.client_data['免费夜固定价格'] 

                            # 计算总的环境税（包括成人/儿童/婴儿）
                            env_tax_cost = (adult_env_tax_price * (adult_amount+1) + child_env_tax_price * child_amount + baby_env_tax_price * baby_amount) * check_in_nights

                            # 减去第三成人免费夜费用
                            if self.client_data['免房费天数'] != 0:
                                if self.client_data['第三成人免费夜固定价格'] == 0:
                                    house_cost -= self.client_data['免房费天数'] * price_set[min(price_set,key=price_set.get)]
                                else:
                                    house_cost -= self.client_data['免房费天数'] * self.client_data['第三成人免费夜固定价格'] 

                            if self.client_data['是否有第三成人'] == '是':
                                if self.client_data['房价是否包含环境税'] == '是':
                                    house_cost -= adult_env_tax_price * check_in_nights * (adult_amount+1) # 成人环境税
                                traffic_cost = adult_traffic_price * (adult_amount+1) + child_traffic_price * child_amount # 计算总交通费
                                qiangzhi_dinner = self.client_data['强制餐单价'] * (adult_amount+1+child_amount) # 强制餐，包括成人和儿童
                            else:
                                if self.client_data['房价是否包含环境税'] == '是':
                                    house_cost -= adult_env_tax_price * check_in_nights * adult_amount # 成人环境税
                                traffic_cost = adult_traffic_price * adult_amount + child_traffic_price * child_amount # 计算总交通费
                                qiangzhi_dinner = self.client_data['强制餐单价'] * (adult_amount+child_amount) # 强制餐，包括成人和儿童


                            self.child_cost = self.child_cost * check_in_nights
                            # 计算与折扣相乘的值
                            if self.client_data['交通是否有优惠'] == '是':
                                if self.client_data['儿童费用是否有优惠'] == '是':
                                    whole_cost = (house_cost + traffic_cost + self.child_cost) * self.client_data['优惠折扣'] + env_tax_cost + self.client_data['强制费单价'] \
                                                * self.client_data['强制费收取天数'] + qiangzhi_dinner + self.client_data['其他附加费：']  
                                    if  self.client_data['房价是否包含环境税'] == '是':
                                        self.string_outprint += " * 实际费用 = ( 房费(已扣环境税,剩" + str(house_cost) + ") + 交通费(" + str(traffic_cost) + ") + 儿童费用(" + str(self.child_cost) + \
                                                                ") ) * 优惠折扣 + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                    else:
                                        self.string_outprint += " * 实际费用 = ( 房费(" + str(house_cost) + ") + 交通费(" + str(traffic_cost) + ") + 儿童费用(" + str(self.child_cost) + \
                                                                ") ) * 优惠折扣 + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                else:
                                    whole_cost = (house_cost + traffic_cost) * self.client_data['优惠折扣'] + self.child_cost + env_tax_cost + self.client_data['强制费单价'] \
                                                * self.client_data['强制费收取天数'] + qiangzhi_dinner + self.client_data['其他附加费：']   
                                    if  self.client_data['房价是否包含环境税'] == '是':
                                        self.string_outprint += " * 实际费用 = ( 房费(已扣环境税,剩" + str(house_cost) + ") + 交通费(" + str(traffic_cost) + ") ) * 优惠折扣 + 儿童费用(" + str(self.child_cost) + \
                                                                ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                    else:
                                        self.string_outprint += " * 实际费用 = ( 房费(" + str(house_cost) + ") + 交通费(" + str(traffic_cost) + ") ) * 优惠折扣 + 儿童费用(" + str(self.child_cost) + \
                                                                ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                            else:
                                if self.client_data['儿童费用是否有优惠'] == '是':
                                    whole_cost = (house_cost + self.child_cost) * self.client_data['优惠折扣'] + traffic_cost + env_tax_cost + self.client_data['强制费单价'] \
                                                * self.client_data['强制费收取天数'] + qiangzhi_dinner + self.client_data['其他附加费：']   
                                    if  self.client_data['房价是否包含环境税'] == '是':
                                        self.string_outprint += " * 实际费用 = ( 房费(已扣环境税,剩" + str(house_cost) + ") + 儿童费用(" + str(self.child_cost) + ") ) * 优惠折扣 + 交通费(" + \
                                                                str(traffic_cost) + ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                    else:
                                        self.string_outprint += " * 实际费用 = ( 房费(" + str(house_cost) + ") + 儿童费用(" + str(self.child_cost) + ") ) * 优惠折扣 + 交通费(" + \
                                                                str(traffic_cost) + ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit          
                                else:
                                    whole_cost = house_cost * self.client_data['优惠折扣'] + self.child_cost + traffic_cost + env_tax_cost + self.client_data['强制费单价'] \
                                                * self.client_data['强制费收取天数'] + qiangzhi_dinner + self.client_data['其他附加费：']        
                                    if  self.client_data['房价是否包含环境税'] == '是':
                                        self.string_outprint += " * 实际费用 = 房费(已扣环境税,剩" + str(house_cost) + ") * 优惠折扣 + 儿童费用(" + str(self.child_cost) + ") + 交通费(" + \
                                                                str(traffic_cost) + ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                    else:
                                        self.string_outprint += " * 实际费用 = 房费(" + str(house_cost) + ") * 优惠折扣 + 儿童费用(" + str(self.child_cost) + ") + 交通费(" + \
                                                                str(traffic_cost) + ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                        else:
                            all_per_list = list(permutations(house_whole_name_list,len(house_whole_name_list)))
                            rb_string_list = []
                            for per in all_per_list:
                                string = ''
                                for Id, s in enumerate(per):
                                    string += str(Id+1) + '. ' + s + '; '
                                rb_string_list.append(string)
                            dialog = inputDialog(rbStringList=rb_string_list)
                            if dialog.exec_():
                                if dialog.chosedId is None:
                                    self.msg("注意","请选择方案！")
                                else:
                                    # print(all_per_list[int(dialog.chosedId)])
                                    # print(days_set)
                                    price_set = {}; price_list = []
                                    start_date_1 = self.client_data['入住时间']
                                    for house_type in all_per_list[int(dialog.chosedId)]:
                                        selectby_house_type = table_set['房型价格表'][(table_set['房型价格表']['房型'] == house_type) & \
                                                                                    (table_set['房型价格表']['人数'] == self.client_data['成人总人数(包括第三成人)'])]
                                        selectby_house_type = selectby_house_type.reset_index(drop=True)
                                        dates_list = self.get_specific_dates_list(start_date_1,days_set[house_type])
                                        sum_price = 0
                                        if len(selectby_house_type) != 0: # 有人数数据
                                            for d in dates_list:
                                                price = selectby_house_type[(d >= selectby_house_type['起始日期']) & \
                                                                          (d <= selectby_house_type['终止日期'])]['单价'].values[0]
                                                sum_price += price
                                                price_list.append(price)
                                        else: # 无人数数据
                                            selectby_house_type = table_set['房型价格表'][(table_set['房型价格表']['房型'] == house_type)]
                                            selectby_house_type = selectby_house_type.reset_index(drop=True)
                                            for d in dates_list:
                                                price = selectby_house_type[(d >= selectby_house_type['起始日期']) & \
                                                                          (d <= selectby_house_type['终止日期'])]['单价'].values[0]
                                                sum_price += price
                                                price_list.append(price)
                                        start_date_1 = self.get_specific_dates_list(dates_list[-1],2)[1]
                                        price_set.update({house_type:sum_price})

                                    #
                                    whole_cost = 0

                                    # 计算正常成人房费
                                    house_cost = 0
                                    for Id, (house_type, house_price) in enumerate(price_set.items()):
                                        house_cost += house_price * days_set[house_type]

                                    # 计算第三成人房费
                                    for house_type in third_person_cost_set:
                                        house_cost += third_person_cost_set[house_type] * days_set[house_type]

                                    # 减去正常成人免费夜费用
                                    if self.client_data['免房费天数'] != 0:
                                        if self.client_data['免费夜固定价格'] == 0:
                                            house_cost -= self.client_data['免房费天数'] * price_set[min(price_set,key=price_set.get)]
                                        else:
                                            house_cost -= self.client_data['免房费天数'] * self.client_data['免费夜固定价格'] 

                                    # 计算总的环境税（包括成人/儿童/婴儿）
                                    env_tax_cost = (adult_env_tax_price * (adult_amount+1) + child_env_tax_price * child_amount + baby_env_tax_price * baby_amount) * check_in_nights

                                    # 减去第三成人免费夜费用
                                    if self.client_data['免房费天数'] != 0:
                                        if self.client_data['第三成人免费夜固定价格'] == 0:
                                            house_cost -= self.client_data['免房费天数'] * price_set[min(price_set,key=price_set.get)]
                                        else:
                                            house_cost -= self.client_data['免房费天数'] * self.client_data['第三成人免费夜固定价格'] 

                                    if self.client_data['是否有第三成人'] == '是':
                                        if self.client_data['房价是否包含环境税'] == '是':
                                            house_cost -= adult_env_tax_price * check_in_nights * (adult_amount+1) # 成人环境税
                                        traffic_cost = adult_traffic_price * (adult_amount+1) + child_traffic_price * child_amount # 计算总交通费
                                        qiangzhi_dinner = self.client_data['强制餐单价'] * (adult_amount+1+child_amount) # 强制餐，包括成人和儿童
                                    else:
                                        if self.client_data['房价是否包含环境税'] == '是':
                                            house_cost -= adult_env_tax_price * check_in_nights * adult_amount # 成人环境税
                                        traffic_cost = adult_traffic_price * adult_amount + child_traffic_price * child_amount # 计算总交通费
                                        qiangzhi_dinner = self.client_data['强制餐单价'] * (adult_amount+child_amount) # 强制餐，包括成人和儿童


                                    self.child_cost = self.child_cost * check_in_nights
                                    # 计算与折扣相乘的值
                                    if self.client_data['交通是否有优惠'] == '是':
                                        if self.client_data['儿童费用是否有优惠'] == '是':
                                            whole_cost = (house_cost + traffic_cost + self.child_cost) * self.client_data['优惠折扣'] + env_tax_cost + self.client_data['强制费单价'] \
                                                        * self.client_data['强制费收取天数'] + qiangzhi_dinner + self.client_data['其他附加费：']  
                                            if  self.client_data['房价是否包含环境税'] == '是':
                                                self.string_outprint += " * 实际费用 = ( 房费(已扣环境税,剩" + str(house_cost) + ") + 交通费(" + str(traffic_cost) + ") + 儿童费用(" + str(self.child_cost) + \
                                                                        ") ) * 优惠折扣 + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                        ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                            else:
                                                self.string_outprint += " * 实际费用 = ( 房费(" + str(house_cost) + ") + 交通费(" + str(traffic_cost) + ") + 儿童费用(" + str(self.child_cost) + \
                                                                        ") ) * 优惠折扣 + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                        ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                        else:
                                            whole_cost = (house_cost + traffic_cost) * self.client_data['优惠折扣'] + self.child_cost + env_tax_cost + self.client_data['强制费单价'] \
                                                        * self.client_data['强制费收取天数'] + qiangzhi_dinner + self.client_data['其他附加费：']   
                                            if  self.client_data['房价是否包含环境税'] == '是':
                                                self.string_outprint += " * 实际费用 = ( 房费(已扣环境税,剩" + str(house_cost) + ") + 交通费(" + str(traffic_cost) + ") ) * 优惠折扣 + 儿童费用(" + str(self.child_cost) + \
                                                                        ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                        ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                            else:
                                                self.string_outprint += " * 实际费用 = ( 房费(" + str(house_cost) + ") + 交通费(" + str(traffic_cost) + ") ) * 优惠折扣 + 儿童费用(" + str(self.child_cost) + \
                                                                        ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                        ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                    else:
                                        if self.client_data['儿童费用是否有优惠'] == '是':
                                            whole_cost = (house_cost + self.child_cost) * self.client_data['优惠折扣'] + traffic_cost + env_tax_cost + self.client_data['强制费单价'] \
                                                        * self.client_data['强制费收取天数'] + qiangzhi_dinner + self.client_data['其他附加费：']   
                                            if  self.client_data['房价是否包含环境税'] == '是':
                                                self.string_outprint += " * 实际费用 = ( 房费(已扣环境税,剩" + str(house_cost) + ") + 儿童费用(" + str(self.child_cost) + ") ) * 优惠折扣 + 交通费(" + \
                                                                        str(traffic_cost) + ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                        ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                            else:
                                                self.string_outprint += " * 实际费用 = ( 房费(" + str(house_cost) + ") + 儿童费用(" + str(self.child_cost) + ") ) * 优惠折扣 + 交通费(" + \
                                                                        str(traffic_cost) + ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                        ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit          
                                        else:
                                            whole_cost = house_cost * self.client_data['优惠折扣'] + self.child_cost + traffic_cost + env_tax_cost + self.client_data['强制费单价'] \
                                                        * self.client_data['强制费收取天数'] + qiangzhi_dinner + self.client_data['其他附加费：']        
                                            if  self.client_data['房价是否包含环境税'] == '是':
                                                self.string_outprint += " * 实际费用 = 房费(已扣环境税,剩" + str(house_cost) + ") * 优惠折扣 + 儿童费用(" + str(self.child_cost) + ") + 交通费(" + \
                                                                        str(traffic_cost) + ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                        ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit
                                            else:
                                                self.string_outprint += " * 实际费用 = 房费(" + str(house_cost) + ") * 优惠折扣 + 儿童费用(" + str(self.child_cost) + ") + 交通费(" + \
                                                                        str(traffic_cost) + ") + 总环境税费(" + str(env_tax_cost) + ") + 强制费(" + str(self.client_data['强制费单价'] * self.client_data['强制费收取天数']) + \
                                                                        ") + 强制餐费(" + str(qiangzhi_dinner) + ") + 其他附加费(" + str(self.client_data['其他附加费：']) + ") = " + str(whole_cost) + ' ' + cal_unit

    def multi_island_calc(self):
        pass

    def travel_fee_calc_func(self):
        self.textBrowser.clear()
        if self.history_tb2_string_output:
            self.client_data = {}
            self.getTableWidgetUnitValue()
            if not self.isgetTableWidgetUnitValueError:
                if self.client_data['入住时间'] == 0:
                    self.msg("错误","入住日期不能为空！")
                elif (len(str(self.client_data['入住时间'])) != 8): 
                    self.msg("错误","入住日期格式不对！")
                elif not self.is_date_valid(str(self.client_data['入住时间'])):
                    self.msg("错误","入住日期不存在！")
                elif self.client_data['成人总人数(包括第三成人)'] == 0:
                    self.msg("错误","成人人数不能为0！")
                else:
                    if len(self.history_tb2_string_output) == 1:
                        self.string_outprint = ' > 单价单计算: ' + '\n'
                        self.child_cost = 0
                        if self.client_data['儿童人数'] != 0: # 规定这里是两个小孩
                            dialog2 = inputDialog2()
                            if dialog2.exec_():
                                if (dialog2.input_1.text() is not None) or (dialog2.input_1.text() != ''):
                                    if self.is_number(dialog2.input_1.text()):
                                        self.child_cost += int(dialog2.input_1.text())
                                        if (dialog2.input_2.text() is not None) or (dialog2.input_2.text() != ''):
                                            if self.is_number(dialog2.input_2.text()):
                                                self.child_cost += int(dialog2.input_2.text())
                                                self.single_island_calc()
                                            else:
                                                self.msg("错误","输入的不是数字！")
                                    else:
                                        self.msg("错误","输入的不是数字！")
                                else:
                                    self.msg("错误","请输入值！")
                        else:
                            self.single_island_calc()
                    else:
                        self.string_outprint = '-' * 23 + ' 多价单计算 ' + '-' * 23 + '\n' 
                        self.multi_island_calc()
                    self.textBrowser.append(self.string_outprint)
        else:
            self.msg("注意","请选择计算价单！")

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    demon = theWindowOfDamon()
    demon.show()
    sys.exit(app.exec_())