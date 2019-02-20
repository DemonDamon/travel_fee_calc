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
from functools import partial 
import pickle
from numpy import ceil

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

        self.buttonBox = QtWidgets.QDialogself.buttonBox(parent=self)
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)  # 设置为水平方向
        self.buttonBox.setStandardButtons(QtWidgets.QDialogself.buttonBox.Cancel | QtWidgets.QDialogself.buttonBox.Ok)  # 确定和取消两个按钮

        self.buttonBox.accepted.connect(self.accept)  # 确定
        self.buttonBox.rejected.connect(self.reject)  # 取消

        layout = QtWidgets.QVBoxLayout()
        layout.addLayout(grid)

        spacerItem = QtWidgets.QSpacerItem(20, 48, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        layout.addItem(spacerItem)
        layout.addWidget(self.buttonBox)
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

        self.buttonBox = QtWidgets.QDialogself.buttonBox(parent=self)
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)  # 设置为水平方向
        self.buttonBox.setStandardButtons(QtWidgets.QDialogself.buttonBox.Cancel | QtWidgets.QDialogself.buttonBox.Ok)  # 确定和取消两个按钮

        self.buttonBox.accepted.connect(self.accept)  # 确定
        self.buttonBox.rejected.connect(self.reject)  # 取消

        layout = QtWidgets.QVBoxLayout()
        layout.addLayout(grid)

        spacerItem = QtWidgets.QSpacerItem(20, 48, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        layout.addItem(spacerItem)
        layout.addWidget(self.buttonBox)
        self.setLayout(layout)

    def center(self):
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        self.move((screen.width() - size.width()) / 2,  
        (screen.height() - size.height()) / 2)
        return int(screen.height()), int(screen.width())

    def resize_screen(self, a, b, c, d):
        return a*self.screen_width/1920, b*self.screen_height/1080, c*self.screen_width/1920, d*self.screen_height/1080


class createTableFirstDiaglog(QDialog):
    def __init__(self,parent=None):
        QDialog.__init__(self, parent)
        # 设置缓存文件夹
        if not os.path.exists('./cache'):
            self.cache_root_dir = os.path.join(os.getcwd(),'cache')
            os.makedirs(self.cache_root_dir)

        self.cache_data_set = {}

        # 设置窗口大小和屏幕位置
        self.font_size = 12
        self.screen_height, self.screen_width = self.center()
        self.resize(1200*self.screen_width/1920, 600*self.screen_height/1080)
        self.setWindowTitle('请输入简要信息创建数据表')

        self.setColNum = 3 # 设置多少列的表格
        self.setRowNum = 30 # 设置多少行的表格
        self.HeaderLabelsList = ['房型名称','最大容纳人数(无附加费)','最大容纳人数(有附加费)']

        self.spacerItem = QtWidgets.QSpacerItem(0, 0, QtWidgets.QSizePolicy.Minimum, \
                            QtWidgets.QSizePolicy.Expanding)

        grid = QtWidgets.QGridLayout()

        # label 1
        self.label1 = QtWidgets.QLabel('合同情况',parent=self)
        self.label1.setStyleSheet("font: 15pt \"Adobe Arabic\";")
        self.label1.setObjectName("label1")
        grid.addWidget(self.label1, 0, 0, 1, 1)

        # label 2
        self.label2 = QtWidgets.QLabel('*有效日期范围',parent=self)
        self.label2.setStyleSheet("font: 10pt \"Adobe Arabic\";")
        self.label2.setObjectName("label1")
        grid.addWidget(self.label2, 1, 0, 1, 1)

        # table
        self.tableWidget1 = QtWidgets.QTableWidget(parent=self)
        left, top, width, height = self.resize_screen(20, 10, 200, 21)
        self.tableWidget1.setGeometry(QtCore.QRect(left, top, width, height)) #370, 40, 361, 661
        self.tableWidget1.setColumnCount(self.setColNum)
        self.tableWidget1.setRowCount(self.setRowNum)
        self.tableWidget1.setColumnWidth(0,200*width/200) #将第一列的单元宽度设置为210
        self.tableWidget1.setColumnWidth(1,160*width/200)
        self.tableWidget1.setColumnWidth(2,160*width/200) 
        self.tableWidget1.setHorizontalHeaderLabels(self.HeaderLabelsList)
        grid.addWidget(self.tableWidget1, 2, 1, 1, 1)

        # inserted layout (sub-layout) 
        self.createSubLayout()
        grid.addLayout(self.sub_layout, 2, 0, 1, 1)

        # next page button
        self.nextbutton_layout = QtWidgets.QGridLayout()
        nextbutton_label = QtWidgets.QLabel('',parent=self)
        self.nextbutton_layout.addWidget(nextbutton_label, 0, 0, 1, 1)
        self.nextbutton_layout.addWidget(nextbutton_label, 0, 1, 1, 1)
        self.nextbutton_layout.addWidget(nextbutton_label, 0, 2, 1, 1)
        self.nextbutton_layout.addWidget(nextbutton_label, 0, 3, 1, 1)
        self.buttonBox = QtWidgets.QPushButton('下一步', parent=self)
        self.buttonBox.clicked.connect(self.close)
        self.buttonBox.clicked.connect(self.nextPageButtonOfFirstPage)
        self.nextbutton_layout.addWidget(self.buttonBox, 0, 4, 1, 1)
        grid.addLayout(self.nextbutton_layout, 3, 1, 1, 1)

        layout = QtWidgets.QVBoxLayout()
        layout.addLayout(grid)
        self.setLayout(layout)

    def createSubLayout(self):
        '''为第一个分页创建初始化界面
        '''
        self.sub_layout = QtWidgets.QGridLayout()
        if not os.path.exists('./cache/valid_date_set_firstpage.pkl'):
            self.sub_layout_input_0_1 = QtWidgets.QLineEdit(parent=self)
            self.sub_layout.addWidget(self.sub_layout_input_0_1, 0,0,1,1)
            self.sub_layout_input_0_1.textChanged.connect(partial(self.leAct,'0')) # 用partial传参，不用lambda

            self.sub_layout_label_0 = QtWidgets.QLabel('-',parent=self)
            self.sub_layout.addWidget(self.sub_layout_label_0, 0,1,1,1)

            self.sub_layout_input_0_2 = QtWidgets.QLineEdit(parent=self)
            self.sub_layout.addWidget(self.sub_layout_input_0_2, 0,2,1,1)
            self.sub_layout_input_0_2.textChanged.connect(partial(self.leAct,'0'))

            self.button = QtWidgets.QPushButton('添加', parent=self)
            self.button.clicked.connect(self.add)
            self.sub_layout.addWidget(self.button,0, 3, 1, 1)
            self.sub_layout.addItem(self.spacerItem)
            self.input_id = 0
            self.cache_data_set.update({'valid_date_'+str(self.input_id) : ''})
            self.sub_layout.addItem(self.spacerItem)
        else:
            with open('./cache/valid_date_set_firstpage.pkl','rb') as f:
                self.cache_data_set = pickle.load(f)
            with open('./cache/tableWidget1_firstpage.pkl','rb') as f:
                cache_df = pickle.load(f)
            for Id, date in enumerate(self.cache_data_set):
                try:
                    date_1, date_2 = self.cache_data_set[date].split('-')
                except:
                    date_1, date_2 = '', ''
                exec("self.sub_layout_input_"+str(Id)+"_1 = QtWidgets.QLineEdit(parent=self)")
                exec("self.sub_layout.addWidget(self.sub_layout_input_"+str(Id)+"_1, "+str(Id)+",0,1,1)")
                exec("self.sub_layout_input_"+str(Id)+"_1.textChanged.connect(partial(self.leAct,'"+str(Id)+"'))")

                exec("self.sub_layout_label_0 = QtWidgets.QLabel('-',parent=self)")
                exec("self.sub_layout.addWidget(self.sub_layout_label_0, "+str(Id)+",1,1,1)")

                exec("self.sub_layout_input_"+str(Id)+"_2 = QtWidgets.QLineEdit(parent=self)")
                exec("self.sub_layout.addWidget(self.sub_layout_input_"+str(Id)+"_2, "+str(Id)+",2,1,1)")
                exec("self.sub_layout_input_"+str(Id)+"_2.textChanged.connect(partial(self.leAct,'"+str(Id)+"'))")
                
                exec("self.sub_layout_input_"+str(Id)+"_1.setText(date_1)")
                exec("self.sub_layout_input_"+str(Id)+"_2.setText(date_2)")
            self.input_id = Id
            self.sub_layout.addItem(self.spacerItem,Id+1,0,1,1)
            # 填入tableWidget的值
            if not cache_df.empty:
                for row_id, value in enumerate(cache_df.values):
                    for col_id, item_value in enumerate(value):
                        if not isnull(item_value):
                            newItem = QTableWidgetItem(item_value)
                            self.tableWidget1.setItem(row_id, col_id, newItem)
        self.button = QtWidgets.QPushButton('添加', parent=self)
        self.button.clicked.connect(self.add)
        self.sub_layout.addWidget(self.button,0, 3, 1, 1)

    def leAct(self,valid_date_id):
        exec("self.cache_data_set['valid_date_'+valid_date_id] = str(self.sub_layout_input_" + valid_date_id + "_1.text()) + \
            '-' + str(self.sub_layout_input_" + valid_date_id + "_2.text())")
        # print(self.cache_data_set)

    def add(self):
        '''实现当按“添加”按钮的时候，不断往下加载LineEdit和label
        '''
        self.sub_layout.removeItem(self.spacerItem)
        exec("self.sub_layout_input_" + str(self.input_id+1) + "_1 = QtWidgets.QLineEdit(parent=self)")
        exec("self.sub_layout.addWidget(self.sub_layout_input_" + str(self.input_id+1) + "_1, self.input_id+1,0,1,1)")
        exec("self.sub_layout_input_" + str(self.input_id+1) + "_1.textChanged.connect(partial(self.leAct,str(self.input_id+1)))")
        exec("self.sub_layout_label_" + str(self.input_id+1) + " = QtWidgets.QLabel('-',parent=self)")
        exec("self.sub_layout.addWidget(self.sub_layout_label_" + str(self.input_id+1) + ", self.input_id+1,1,1,1)")
        # self.sub_layout_label = QtWidgets.QLabel('-',parent=self)
        # self.sub_layout.addWidget(self.sub_layout_label, self.input_id+1,1,1,1)
        exec("self.sub_layout_input_" + str(self.input_id+1) + "_2 = QtWidgets.QLineEdit(parent=self)")
        exec("self.sub_layout.addWidget(self.sub_layout_input_" + str(self.input_id+1) + "_2, self.input_id+1,2,1,1) ")
        exec("self.sub_layout_input_" + str(self.input_id+1) + "_2.textChanged.connect(partial(self.leAct,str(self.input_id+1)))")
        self.sub_layout.addItem(self.spacerItem,self.input_id+2,0,1,1) # 一定要加self.input_id+2,0,1,1
        self.cache_data_set.update({'valid_date_'+str(self.input_id+1) : ''})
        self.input_id += 1

    def nextPageButtonOfFirstPage(self):
        '''跳转下一个页面对话，并在跳转之前，1. 把当前页面的数据序列化到缓存文件夹,
           2. 检查第二个页面是否存在序列化文件，如果存在就加载进来
        '''
        # 序列化cache_data_set到pkl文件
        with open('./cache/valid_date_set_firstpage.pkl','wb') as cache_file:
            pickle.dump(self.cache_data_set, cache_file)
        
        # 将tableWidget1中的数据抽取保存成DataFrame格式
        cache_df = DataFrame(columns=self.HeaderLabelsList)
        for row_Id in range(self.setRowNum): #setRowNum=30
            for col_Id, col_name in enumerate(self.HeaderLabelsList):
                if self.tableWidget1.item(row_Id, col_Id) is not None:
                    cache_df.loc[row_Id,col_name] = self.tableWidget1.item(row_Id, col_Id).text()

        # 序列化DataFrame格式的tableWidget1数据到pkl文件
        with open('./cache/tableWidget1_firstpage.pkl','wb') as cache_file:
            pickle.dump(cache_df, cache_file)

        # 如果序列化文件存在且为非空文件，则重新设置第二个对话页面
        if os.path.exists("./cache/valid_date_set_secondpage.pkl"):
            with open("./cache/valid_date_set_secondpage.pkl","rb") as f:
                valid_date_set_secondpage = pickle.load(f)
            
            with open('./cache/tableWidget2_secondpage.pkl','rb') as cache_file:
                cache_df = pickle.load(cache_file)

            class resetSecondDiaglog(createTableSecondDiaglog):
                def __init__(self,parent=None):
                    createTableSecondDiaglog.__init__(self,parent)
                    self.sub_layout.removeItem(self.spacerItem)
                    for Id, csd in enumerate(valid_date_set_secondpage):
                        try:
                            date_1, date_2 = valid_date_set_secondpage[csd].split('-')
                        except:
                            date_1, date_2 = '', ''
                        exec("self.sub_layout_input_" + str(Id) + "_1 = QtWidgets.QLineEdit(parent=self)")
                        exec("self.sub_layout.addWidget(self.sub_layout_input_" + str(Id) + "_1, Id,0,1,1)")
                        exec("self.sub_layout_input_" + str(Id) + "_1.textChanged.connect(partial(self.leAct,str(Id)))")
                        exec("self.sub_layout_label_" + str(Id) + " = QtWidgets.QLabel('-',parent=self)")
                        exec("self.sub_layout.addWidget(self.sub_layout_label_" + str(Id) + ", Id,1,1,1)")
                        exec("self.sub_layout_input_" + str(Id) + "_2 = QtWidgets.QLineEdit(parent=self)")
                        exec("self.sub_layout.addWidget(self.sub_layout_input_" + str(Id) + "_2, Id,2,1,1) ")
                        exec("self.sub_layout_input_" + str(Id) + "_2.textChanged.connect(partial(self.leAct,str(Id)))")
                        # 注意这里setText要放在最后
                        exec("self.sub_layout_input_" + str(Id) + "_1.setText(date_1)")
                        exec("self.sub_layout_input_" + str(Id) + "_2.setText(date_2)")
                        self.input_id += 1
                    self.sub_layout.addItem(self.spacerItem,Id+1,0,1,1) 
                    # 填入tableWidget的值
                    if not cache_df.empty:
                        for row_id, value in enumerate(cache_df.values):
                            for col_id, item_value in enumerate(value):
                                if not isnull(item_value):
                                    newItem = QTableWidgetItem(item_value)
                                    self.tableWidget2.setItem(row_id, col_id, newItem)
                        
            # 重新创建第二个页面
            dialog = resetSecondDiaglog()
            if dialog.exec_():
                pass
        else:
            # 创建第二个页面
            dialog = createTableSecondDiaglog()
            if dialog.exec_():
                pass

    def center(self):
        screen = QDesktopWidget().screenGeometry()
        return int(screen.height()), int(screen.width())

    def resize_screen(self, a, b, c, d):
        return a*self.screen_width/1920, b*self.screen_height/1080, c*self.screen_width/1920, d*self.screen_height/1080


class createTableSecondDiaglog(QDialog):
    def __init__(self,parent=None):
        QDialog.__init__(self, parent)
        self.cache_data_set = {}

        # 设置窗口大小和屏幕位置
        self.font_size = 12
        self.screen_height, self.screen_width = self.center()
        self.resize(1000*self.screen_width/1920, 600*self.screen_height/1080)
        self.setWindowTitle('请输入简要信息创建数据表')

        self.setColNum = 2 # 设置多少列的表格
        self.setRowNum = 30 # 设置多少行的表格
        self.HeaderLabelsList = ['打包组合房型全称','打包所含费用']

        self.spacerItem = QtWidgets.QSpacerItem(0, 0, QtWidgets.QSizePolicy.Minimum, \
                            QtWidgets.QSizePolicy.Expanding)

        grid = QtWidgets.QGridLayout()

        # label 1
        self.label1 = QtWidgets.QLabel('打包情况',parent=self)
        self.label1.setStyleSheet("font: 15pt \"Adobe Arabic\";")
        self.label1.setObjectName("label1")
        grid.addWidget(self.label1, 0, 0, 1, 1)

        # label 2
        self.label2 = QtWidgets.QLabel('*有效日期范围',parent=self)
        self.label2.setStyleSheet("font: 10pt \"Adobe Arabic\";")
        self.label2.setObjectName("label1")
        grid.addWidget(self.label2, 1, 0, 1, 1)

        # inserted layout (sub-layout) 
        self.createSubLayout()
        grid.addLayout(self.sub_layout, 2, 0, 1, 1)

        # table
        self.tableWidget2 = QtWidgets.QTableWidget(parent=self)
        left, top, width, height = self.resize_screen(20, 10, 200, 21)
        self.tableWidget2.setGeometry(QtCore.QRect(left, top, width, height)) #370, 40, 361, 661
        self.tableWidget2.setColumnCount(self.setColNum)
        self.tableWidget2.setRowCount(self.setRowNum)
        self.tableWidget2.setColumnWidth(0,200*width/200) #将第一列的单元宽度设置为210
        self.tableWidget2.setColumnWidth(1,160*width/200)
        self.tableWidget2.setColumnWidth(2,160*width/200) 
        self.tableWidget2.setHorizontalHeaderLabels(self.HeaderLabelsList)
        grid.addWidget(self.tableWidget2, 2, 1, 1, 1)

        # back page button and next page button
        self.nextbutton_layout = QtWidgets.QGridLayout()
        nextbutton_label = QtWidgets.QLabel('',parent=self)

        self.buttonBox_back = QtWidgets.QPushButton('上一步', parent=self)
        self.buttonBox_back.clicked.connect(self.close)
        self.buttonBox_back.clicked.connect(self.backPageButtonOfSecondPage)
        self.nextbutton_layout.addWidget(self.buttonBox_back, 0, 0, 1, 1)

        self.nextbutton_layout.addWidget(nextbutton_label, 0, 1, 1, 1)
        self.nextbutton_layout.addWidget(nextbutton_label, 0, 2, 1, 1)
        self.nextbutton_layout.addWidget(nextbutton_label, 0, 3, 1, 1)
        
        self.buttonBox = QtWidgets.QPushButton('下一步', parent=self)
        self.buttonBox.clicked.connect(self.close)
        self.buttonBox.clicked.connect(self.nextPageButtonOfSecondPage)
        self.nextbutton_layout.addWidget(self.buttonBox, 0, 4, 1, 1)
        grid.addLayout(self.nextbutton_layout, 3, 1, 1, 1)

        layout = QtWidgets.QVBoxLayout()
        layout.addLayout(grid)
        self.setLayout(layout)

    def createSubLayout(self):
        '''为第一个分页创建初始化界面
        '''
        self.sub_layout = QtWidgets.QGridLayout()

        self.sub_layout_input_0_1 = QtWidgets.QLineEdit(parent=self)
        self.sub_layout.addWidget(self.sub_layout_input_0_1, 0,0,1,1)
        self.sub_layout_input_0_1.textChanged.connect(partial(self.leAct,'0')) # 用partial传参，不用lambda

        self.sub_layout_label_0 = QtWidgets.QLabel('-',parent=self)
        self.sub_layout.addWidget(self.sub_layout_label_0, 0,1,1,1)

        self.sub_layout_input_0_2 = QtWidgets.QLineEdit(parent=self)
        self.sub_layout.addWidget(self.sub_layout_input_0_2, 0,2,1,1)
        self.sub_layout_input_0_2.textChanged.connect(partial(self.leAct,'0'))

        self.button = QtWidgets.QPushButton('添加', parent=self)
        self.button.clicked.connect(self.add)
        self.sub_layout.addWidget(self.button,0, 3, 1, 1)
        self.sub_layout.addItem(self.spacerItem)
        self.input_id = 0
        self.cache_data_set.update({'valid_date_'+str(self.input_id) : ''})

    def leAct(self,valid_date_id):
        exec("self.cache_data_set['valid_date_'+valid_date_id] = str(self.sub_layout_input_" + valid_date_id + "_1.text()) + \
            '-' + str(self.sub_layout_input_" + valid_date_id + "_2.text())")
        # print(self.cache_data_set)

    def add(self):
        '''实现当按“添加”按钮的时候，不断往下加载LineEdit和label
        '''
        self.sub_layout.removeItem(self.spacerItem)
        exec("self.sub_layout_input_" + str(self.input_id+1) + "_1 = QtWidgets.QLineEdit(parent=self)")
        exec("self.sub_layout.addWidget(self.sub_layout_input_" + str(self.input_id+1) + "_1, self.input_id+1,0,1,1)")
        exec("self.sub_layout_input_" + str(self.input_id+1) + "_1.textChanged.connect(partial(self.leAct,str(self.input_id+1)))")
        exec("self.sub_layout_label_" + str(self.input_id+1) + " = QtWidgets.QLabel('-',parent=self)")
        exec("self.sub_layout.addWidget(self.sub_layout_label_" + str(self.input_id+1) + ", self.input_id+1,1,1,1)")
        exec("self.sub_layout_input_" + str(self.input_id+1) + "_2 = QtWidgets.QLineEdit(parent=self)")
        exec("self.sub_layout.addWidget(self.sub_layout_input_" + str(self.input_id+1) + "_2, self.input_id+1,2,1,1) ")
        exec("self.sub_layout_input_" + str(self.input_id+1) + "_2.textChanged.connect(partial(self.leAct,str(self.input_id+1)))")
        self.sub_layout.addItem(self.spacerItem,self.input_id+2,0,1,1) # 一定要加self.input_id+2,0,1,1
        self.cache_data_set.update({'valid_date_'+str(self.input_id+1) : ''})
        self.input_id += 1

    def nextPageButtonOfSecondPage(self):
        '''跳转下一个页面对话，并在跳转之前，1. 把当前页面的数据序列化到缓存文件夹,
           2. 如果第三个页面有缓存序列文件，则加载进来
        '''
        # 序列化cache_data_set到pkl文件
        with open('./cache/valid_date_set_secondpage.pkl','wb') as cache_file:
            pickle.dump(self.cache_data_set, cache_file)
        
        # 将tableWidget2中的数据抽取保存成DataFrame格式
        cache_df = DataFrame(columns=self.HeaderLabelsList)
        for row_Id in range(self.setRowNum): #setRowNum=30
            for col_Id, col_name in enumerate(self.HeaderLabelsList):
                if self.tableWidget2.item(row_Id, col_Id) is not None:
                    cache_df.loc[row_Id,col_name] = self.tableWidget2.item(row_Id, col_Id).text()

        # 序列化DataFrame格式的tableWidget2数据到pkl文件
        with open('./cache/tableWidget2_secondpage.pkl','wb') as cache_file:
            pickle.dump(cache_df, cache_file)

        if os.path.exists('./cache/transport_thirdpage.pkl'):
            with open('./cache/transport_thirdpage.pkl','rb') as f:
                transport_data_set = pickle.load(f)
            with open('./cache/dinnertype_thirdpage.pkl','rb') as f:
                dinnertype_data_set = pickle.load(f)
            with open('./cache/child_age_stratification.pkl','rb') as f:
                child_age_stratification = pickle.load(f)

            class resetThirdDiaglog(createTableThirdDiaglog):
                def __init__(self,parent=None):
                    createTableThirdDiaglog.__init__(self,parent)
                    # 交通组合方式的sub_layout_0
                    self.sub_layout_0.removeItem(self.spacerItem_1)
                    for Id, csd in enumerate(transport_data_set):
                        try:
                            transport = transport_data_set[csd]
                        except:
                            transport = ''
                        exec("self.sub_layout_input_0_" + str(Id) + " = QtWidgets.QLineEdit(parent=self)")
                        exec("self.sub_layout_0.addWidget(self.sub_layout_input_0_" + str(Id) + ", Id,0,1,1)")
                        exec("self.sub_layout_input_0_" + str(Id) + ".textChanged.connect(partial(self.leAct,0,str(Id)))")
                        # 注意这里setText要放在最后
                        exec("self.sub_layout_input_0_" + str(Id) + ".setText(transport)")
                        self.input_id_0 += 1
                    self.sub_layout_0.addItem(self.spacerItem_1,Id+1,0,1,1)

                    # 餐费类型的sub_layout_1
                    self.sub_layout_1.removeItem(self.spacerItem_2)
                    for Id, csd in enumerate(dinnertype_data_set):
                        try:
                            dinnertype = dinnertype_data_set[csd]
                        except:
                            dinnertype = ''
                        exec("self.sub_layout_input_1_" + str(Id) + " = QtWidgets.QLineEdit(parent=self)")
                        exec("self.sub_layout_1.addWidget(self.sub_layout_input_1_" + str(Id) + ", Id,0,1,1)")
                        exec("self.sub_layout_input_1_" + str(Id) + ".textChanged.connect(partial(self.leAct,1,str(Id)))")
                        # 注意这里setText要放在最后
                        exec("self.sub_layout_input_1_" + str(Id) + ".setText(dinnertype)")
                        self.input_id_1 += 1
                    self.sub_layout_1.addItem(self.spacerItem_2,Id+1,0,1,1)

                    # 儿童年龄分层的sub_layout_2
                    self.sub_layout_2.removeItem(self.spacerItem_3)
                    for Id, csd in enumerate(child_age_stratification):
                        try:
                            child_age = child_age_stratification[csd]
                        except:
                            child_age = ''
                        exec("self.sub_layout_input_2_" + str(Id) + " = QtWidgets.QLineEdit(parent=self)")
                        exec("self.sub_layout_2.addWidget(self.sub_layout_input_2_" + str(Id) + ", Id,0,1,1)")
                        exec("self.sub_layout_input_2_" + str(Id) + ".textChanged.connect(partial(self.leAct,2,str(Id)))")
                        # 注意这里setText要放在最后
                        exec("self.sub_layout_input_2_" + str(Id) + ".setText(child_age)")
                        self.input_id_2 += 1
                    self.sub_layout_2.addItem(self.spacerItem_3,Id+1,0,1,1)

            # 重新创建第三个页面
            dialog = resetThirdDiaglog()
            if dialog.exec_():
                pass    
        else:    
            # 创建第三个页面
            dialog = createTableThirdDiaglog()
            if dialog.exec_():
                pass
        
    def backPageButtonOfSecondPage(self):
        '''回转到上一个页面对话，并在跳转之前，1.把当前页面的数据序列化到缓存文件夹，
           2.并把之前页面的数据从缓存文件夹中加载进来
        '''
        # 序列化cache_data_set到pkl文件
        with open('./cache/valid_date_set_secondpage.pkl','wb') as cache_file:
            pickle.dump(self.cache_data_set, cache_file)
        
        # 将tableWidget2中的数据抽取保存成DataFrame格式
        cache_df = DataFrame(columns=self.HeaderLabelsList)
        for row_Id in range(self.setRowNum): #setRowNum=30
            for col_Id, col_name in enumerate(self.HeaderLabelsList):
                if self.tableWidget2.item(row_Id, col_Id) is not None:
                    cache_df.loc[row_Id,col_name] = self.tableWidget2.item(row_Id, col_Id).text()

        # 序列化DataFrame格式的tableWidget2数据到pkl文件
        with open('./cache/tableWidget2_secondpage.pkl','wb') as cache_file:
            pickle.dump(cache_df, cache_file)

        # 反序列化pkl文件到cache_data_set
        with open('./cache/valid_date_set_firstpage.pkl','rb') as cache_file:
            cache_data_set = pickle.load(cache_file)
        
        # 反序列化tableWidget1数据到cache_df
        with open('./cache/tableWidget1_firstpage.pkl','rb') as cache_file:
            cache_df = pickle.load(cache_file)

        # 重新创建第一个页面
        class resetFirstDiaglog(createTableFirstDiaglog):
            def __init__(self,parent=None):
                createTableFirstDiaglog.__init__(self,parent)
                self.sub_layout.removeItem(self.spacerItem)
                # 填入lineEdit的值
                for Id, csd in enumerate(cache_data_set):
                    try:
                        date_1, date_2 = cache_data_set[csd].split('-')
                    except:
                        date_1, date_2 = '', ''
                    exec("self.sub_layout_input_" + str(Id) + "_1 = QtWidgets.QLineEdit(parent=self)")
                    exec("self.sub_layout.addWidget(self.sub_layout_input_" + str(Id) + "_1, Id,0,1,1)")
                    exec("self.sub_layout_input_" + str(Id) + "_1.textChanged.connect(partial(self.leAct,str(Id)))")
                    exec("self.sub_layout_label_" + str(Id) + " = QtWidgets.QLabel('-',parent=self)")
                    exec("self.sub_layout.addWidget(self.sub_layout_label_" + str(Id) + ", Id,1,1,1)")
                    exec("self.sub_layout_input_" + str(Id) + "_2 = QtWidgets.QLineEdit(parent=self)")
                    exec("self.sub_layout.addWidget(self.sub_layout_input_" + str(Id) + "_2, Id,2,1,1) ")
                    exec("self.sub_layout_input_" + str(Id) + "_2.textChanged.connect(partial(self.leAct,str(Id)))")
                    # 注意这里setText要放在最后
                    exec("self.sub_layout_input_" + str(Id) + "_1.setText(date_1)")
                    exec("self.sub_layout_input_" + str(Id) + "_2.setText(date_2)")
                    self.input_id += 1
                self.sub_layout.addItem(self.spacerItem,Id+1,0,1,1) 
                # 填入tableWidget的值
                if not cache_df.empty:
                    for row_id, value in enumerate(cache_df.values):
                        for col_id, item_value in enumerate(value):
                            if not isnull(item_value):
                                newItem = QTableWidgetItem(item_value)
                                self.tableWidget1.setItem(row_id, col_id, newItem)
                        
        dialog = resetFirstDiaglog()
        if dialog.exec_():
            pass

    def center(self):
        screen = QDesktopWidget().screenGeometry()
        return int(screen.height()), int(screen.width())

    def resize_screen(self, a, b, c, d):
        return a*self.screen_width/1920, b*self.screen_height/1080, c*self.screen_width/1920, d*self.screen_height/1080


class createTableThirdDiaglog(QDialog):
    def __init__(self,parent=None):
        QDialog.__init__(self, parent)
        self.cache_data_set_0 = {'0':'单程内飞快艇和单程水飞','1':'单程快艇和单程水飞',\
                                 '2':'往返快艇','3':'往返内飞','4':'往返水飞','5':'往返内飞快艇'}
        self.cache_data_set_1 = {'0':'BB','1':'HB','2':'FB','3':'AI','4':'PAI'}
        self.cache_data_set_2 = {}

        # 设置窗口大小和屏幕位置
        self.font_size = 12
        self.screen_height, self.screen_width = self.center()
        self.resize(750*self.screen_width/1920, 600*self.screen_height/1080)
        self.setWindowTitle('请输入简要信息创建数据表')

        self.spacerItem_1 = QtWidgets.QSpacerItem(0, 0, QtWidgets.QSizePolicy.Minimum, \
                            QtWidgets.QSizePolicy.Expanding)
        self.spacerItem_2 = QtWidgets.QSpacerItem(0, 0, QtWidgets.QSizePolicy.Minimum, \
                            QtWidgets.QSizePolicy.Expanding)
        self.spacerItem_3 = QtWidgets.QSpacerItem(0, 0, QtWidgets.QSizePolicy.Minimum, \
                            QtWidgets.QSizePolicy.Expanding)
        grid = QtWidgets.QGridLayout()

        # label 1
        self.label1 = QtWidgets.QLabel('其他信息',parent=self)
        self.label1.setStyleSheet("font: 15pt \"Adobe Arabic\";")
        self.label1.setObjectName("label1")
        grid.addWidget(self.label1, 0, 0, 1, 1) # 0, 0, 1, 1

        # 表格名称输入框和label_name
        self.label_name = QtWidgets.QLabel('表格名称：',parent=self)
        self.label_name.setStyleSheet("font: 8pt \"Adobe Arabic\";")
        self.label_name.setObjectName("label_name")
        grid.addWidget(self.label_name, 1, 0, 1, 1)
        self.input = QtWidgets.QLineEdit(parent=self)
        grid.addWidget(self.input,2, 0, 1, 1)

        # label 2
        self.label2 = QtWidgets.QLabel('*交通组合方式',parent=self)
        self.label2.setStyleSheet("font: 10pt \"Adobe Arabic\";")
        self.label2.setObjectName("label2")
        grid.addWidget(self.label2, 3, 0, 1, 1) #1, 0, 1, 1

        # inserted layout (sub-layout) 
        self.createSubLayout(0)
        grid.addLayout(self.sub_layout_0, 4, 0, 1, 1) #2, 0, 1, 1

        # label 3
        self.label3 = QtWidgets.QLabel('*餐费类型',parent=self)
        self.label3.setStyleSheet("font: 10pt \"Adobe Arabic\";")
        self.label3.setObjectName("label3")
        grid.addWidget(self.label3, 3, 1, 1, 1) # 1, 1, 1, 1

        # inserted layout (sub-layout) 
        self.createSubLayout(1)
        grid.addLayout(self.sub_layout_1, 4, 1, 1, 1) #2, 1, 1, 1

        # label 4
        self.label4 = QtWidgets.QLabel('*不同收费年龄层(儿童)',parent=self)
        self.label4.setStyleSheet("font: 10pt \"Adobe Arabic\";")
        self.label4.setObjectName("label4")
        grid.addWidget(self.label4, 5, 0, 1, 1) # 1, 1, 1, 1

        # inserted layout (sub-layout) 
        self.createSubLayout(2)
        grid.addLayout(self.sub_layout_2, 6, 0, 1, 1) #

        # back page button and next page button
        self.nextbutton_layout = QtWidgets.QGridLayout()
        nextbutton_label = QtWidgets.QLabel('',parent=self)

        self.buttonBox_back = QtWidgets.QPushButton('上一步', parent=self)
        self.buttonBox_back.clicked.connect(self.close)
        self.buttonBox_back.clicked.connect(self.backPageButtonOfThirdPage)
        self.nextbutton_layout.addWidget(self.buttonBox_back, 0, 0, 1, 1)

        self.nextbutton_layout.addWidget(nextbutton_label, 0, 1, 1, 1)
        self.nextbutton_layout.addWidget(nextbutton_label, 0, 2, 1, 1)
        self.nextbutton_layout.addWidget(nextbutton_label, 0, 3, 1, 1)
        
        self.buttonBox = QtWidgets.QPushButton('完成', parent=self)
        self.buttonBox.clicked.connect(self.close)
        self.buttonBox.clicked.connect(self.finishPageButtonOfThirdPage)
        self.nextbutton_layout.addWidget(self.buttonBox, 0, 4, 1, 1)
        grid.addLayout(self.nextbutton_layout, 7, 1, 1, 1) #5, 1, 1, 1

        layout = QtWidgets.QVBoxLayout()
        layout.addLayout(grid)
        self.setLayout(layout)

    def createSubLayout(self,sub_layout_id):
        '''为第三个分页创建初始化界面
        '''
        exec("self.sub_layout_" + str(sub_layout_id) + " = QtWidgets.QGridLayout()")
        if sub_layout_id == 0:
            for i in self.cache_data_set_0: 
                exec("self.sub_layout_input_" + str(sub_layout_id) + "_" + i + " = QtWidgets.QLineEdit(parent=self)")
                exec("self.sub_layout_" + str(sub_layout_id) + ".addWidget(self.sub_layout_input_" + str(sub_layout_id) + "_" + i + ", " + i + ",0,1,1)")
                exec("self.sub_layout_input_" + str(sub_layout_id) + "_" + i + ".textChanged.connect(partial(self.leAct," + str(sub_layout_id) + ",'" + i + "'))")
                exec("self.sub_layout_input_" + str(sub_layout_id) + "_" + i + ".setText(self.cache_data_set_0['" + i + "'])")
        elif sub_layout_id == 1:
            for i in self.cache_data_set_1: 
                exec("self.sub_layout_input_" + str(sub_layout_id) + "_" + i + " = QtWidgets.QLineEdit(parent=self)")
                exec("self.sub_layout_" + str(sub_layout_id) + ".addWidget(self.sub_layout_input_" + str(sub_layout_id) + "_" + i + ", " + i + ",0,1,1)")
                exec("self.sub_layout_input_" + str(sub_layout_id) + "_" + i + ".textChanged.connect(partial(self.leAct," + str(sub_layout_id) + ",'" + i + "'))")
                exec("self.sub_layout_input_" + str(sub_layout_id) + "_" + i + ".setText(self.cache_data_set_1['" + i + "'])")
        else:
            exec("self.sub_layout_input_" + str(sub_layout_id) + "_0 = QtWidgets.QLineEdit(parent=self)")
            exec("self.sub_layout_" + str(sub_layout_id) + ".addWidget(self.sub_layout_input_" + str(sub_layout_id) + "_0, 0,0,1,1)")
            exec("self.sub_layout_input_" + str(sub_layout_id) + "_0.textChanged.connect(partial(self.leAct," + str(sub_layout_id) + ",'0'))")
    
        exec("self.button_" + str(sub_layout_id) + " = QtWidgets.QPushButton('添加',parent=self)")
        exec("self.button_" + str(sub_layout_id) + ".clicked.connect(partial(self.add,sub_layout_id))")
        exec("self.sub_layout_" + str(sub_layout_id) + ".addWidget(self.button_" + str(sub_layout_id) + ", 0, 1, 1, 1)")
        exec("self.sub_layout_" + str(sub_layout_id) + ".addItem(self.spacerItem_" + str(sub_layout_id+1) + ")")
        if sub_layout_id == 0:
            self.input_id_0 = len(self.cache_data_set_0) - 1
        elif sub_layout_id == 1:
            self.input_id_1 = len(self.cache_data_set_1) - 1
        else:
            self.input_id_2 = 0

    def leAct(self,sub_layout_id,input_id):
        if sub_layout_id == 0:
            exec("self.cache_data_set_0[input_id] = str(self.sub_layout_input_0_" + str(input_id) + ".text())")
        elif sub_layout_id == 1:
            exec("self.cache_data_set_1[input_id] = str(self.sub_layout_input_1_" + str(input_id) + ".text())")                
        else:
            exec("self.cache_data_set_2[input_id] = str(self.sub_layout_input_2_" + str(input_id) + ".text())")

    def add(self,sub_layout_id):
        '''实现当按“添加”按钮的时候，不断往下加载LineEdit
        '''
        if sub_layout_id == 0:
            self.sub_layout_0.removeItem(self.spacerItem_1)
            exec("self.sub_layout_input_0_" + str(self.input_id_0+1) + \
                " = QtWidgets.QLineEdit(parent=self)")
            exec("self.sub_layout_0.addWidget(self.sub_layout_input_0_" + \
                str(self.input_id_0+1) + ", self.input_id_0+1,0,1,1)")
            exec("self.sub_layout_input_0_" + str(self.input_id_0+1) + ".textChanged.connect(partial(self.leAct,0,str(self.input_id_0+1)))")
        
            self.sub_layout_0.addItem(self.spacerItem_1,self.input_id_0+2,0,1,1) # 一定要加self.input_id+2,0,1,1
            self.cache_data_set_0.update({str(self.input_id_0+1) : ''})
            self.input_id_0 += 1
        elif sub_layout_id == 1:
            self.sub_layout_1.removeItem(self.spacerItem_2)
            exec("self.sub_layout_input_1_" + str(self.input_id_1+1) + \
                " = QtWidgets.QLineEdit(parent=self)")
            exec("self.sub_layout_1.addWidget(self.sub_layout_input_1_" + \
                str(self.input_id_1+1) + ", self.input_id_1+1,0,1,1)")
            exec("self.sub_layout_input_1_" + str(self.input_id_1+1) + ".textChanged.connect(partial(self.leAct,1,str(self.input_id_1+1)))")
        
            self.sub_layout_1.addItem(self.spacerItem_2,self.input_id_1+2,0,1,1) # 一定要加self.input_id+2,0,1,1
            self.cache_data_set_1.update({str(self.input_id_1+1) : ''})
            self.input_id_1 += 1
        else:
            if eval("'-' not in self.sub_layout_input_2_" + str(self.input_id_2) + ".text()"):
                self.msg('错误','儿童年龄分层数据输入格式有误！')
            else:
                self.sub_layout_2.removeItem(self.spacerItem_3)
                exec("self.sub_layout_input_2_" + str(self.input_id_2+1) + \
                    " = QtWidgets.QLineEdit(parent=self)")
                exec("self.sub_layout_2.addWidget(self.sub_layout_input_2_" + \
                    str(self.input_id_2+1) + ", self.input_id_2+1,0,1,1)")
                exec("self.sub_layout_input_2_" + str(self.input_id_2+1) + ".textChanged.connect(partial(self.leAct,2,str(self.input_id_2+1)))")
            
                self.sub_layout_2.addItem(self.spacerItem_3,self.input_id_2+2,0,1,1) # 一定要加self.input_id+2,0,1,1
                self.cache_data_set_2.update({str(self.input_id_2+1) : ''})
                self.input_id_2 += 1

    def finishPageButtonOfThirdPage(self):
        '''1. 将序列化文件重新加载进来，并据此生成初步数据表
           2. 删除缓存文件夹里面的序列化数据
        '''
        with open('./cache/valid_date_set_firstpage.pkl','rb') as f:
            valid_date_set_firstpage = pickle.load(f)
        
        with open('./cache/valid_date_set_secondpage.pkl','rb') as f:
            valid_date_set_secondpage = pickle.load(f)
        
        with open('./cache/tableWidget1_firstpage.pkl','rb') as f:
            tableWidget1_firstpage = pickle.load(f)
        
        with open('./cache/tableWidget2_secondpage.pkl','rb') as f:
            tableWidget2_secondpage = pickle.load(f)
        
        try:
            self.createTableForIsland(valid_date_set_firstpage,\
                                    valid_date_set_secondpage,\
                                    tableWidget1_firstpage,\
                                    tableWidget2_secondpage,\
                                    self.cache_data_set_0,\
                                    self.cache_data_set_1,\
                                    self.cache_data_set_2)
        except:
            with open('./cache/transport_thirdpage.pkl','wb') as cache_file:
                pickle.dump(self.cache_data_set_0, cache_file)
            with open('./cache/dinnertype_thirdpage.pkl','wb') as cache_file:
                pickle.dump(self.cache_data_set_1, cache_file)
            with open('./cache/child_age_stratification.pkl','wb') as cache_file:
                pickle.dump(self.cache_data_set_2, cache_file)  
            self.msg("错误","表格未创建成功，请重新创建！")

    def backPageButtonOfThirdPage(self):
        '''回转到上一个页面对话，并在跳转之前，1.把当前页面的数据序列化到缓存文件夹，
           2.并把之前页面的数据从缓存文件夹中加载进来
        '''
        # 序列化交通组合方式和餐费类型数据到pkl文件
        with open('./cache/transport_thirdpage.pkl','wb') as cache_file:
            pickle.dump(self.cache_data_set_0, cache_file)
        with open('./cache/dinnertype_thirdpage.pkl','wb') as cache_file:
            pickle.dump(self.cache_data_set_1, cache_file)
        with open('./cache/child_age_stratification.pkl','wb') as cache_file:
            pickle.dump(self.cache_data_set_2, cache_file)  
        
        # 反序列化pkl文件到cache_data_set
        with open('./cache/valid_date_set_secondpage.pkl','rb') as cache_file:
            cache_data_set = pickle.load(cache_file)
        
        # 反序列化tableWidget2数据到cache_df
        with open('./cache/tableWidget2_secondpage.pkl','rb') as cache_file:
            cache_df = pickle.load(cache_file)

        # 重新创建第二个页面
        class resetSecondDiaglog(createTableSecondDiaglog):
            def __init__(self,parent=None):
                createTableSecondDiaglog.__init__(self,parent)
                self.sub_layout.removeItem(self.spacerItem)
                # 填入lineEdit的值
                for Id, csd in enumerate(cache_data_set):
                    try:
                        date_1, date_2 = cache_data_set[csd].split('-')
                    except:
                        date_1, date_2 = '', ''
                    exec("self.sub_layout_input_" + str(Id) + "_1 = QtWidgets.QLineEdit(parent=self)")
                    exec("self.sub_layout.addWidget(self.sub_layout_input_" + str(Id) + "_1, Id,0,1,1)")
                    exec("self.sub_layout_input_" + str(Id) + "_1.textChanged.connect(partial(self.leAct,str(Id)))")
                    exec("self.sub_layout_label_" + str(Id) + " = QtWidgets.QLabel('-',parent=self)")
                    exec("self.sub_layout.addWidget(self.sub_layout_label_" + str(Id) + ", Id,1,1,1)")
                    exec("self.sub_layout_input_" + str(Id) + "_2 = QtWidgets.QLineEdit(parent=self)")
                    exec("self.sub_layout.addWidget(self.sub_layout_input_" + str(Id) + "_2, Id,2,1,1) ")
                    exec("self.sub_layout_input_" + str(Id) + "_2.textChanged.connect(partial(self.leAct,str(Id)))")
                    # 注意这里setText要放在最后
                    exec("self.sub_layout_input_" + str(Id) + "_1.setText(date_1)")
                    exec("self.sub_layout_input_" + str(Id) + "_2.setText(date_2)")
                    self.input_id += 1
                self.sub_layout.addItem(self.spacerItem,Id+1,0,1,1) 
                # 填入tableWidget的值
                if not cache_df.empty:
                    for row_id, value in enumerate(cache_df.values):
                        for col_id, item_value in enumerate(value):
                            if not isnull(item_value):
                                newItem = QTableWidgetItem(item_value)
                                self.tableWidget2.setItem(row_id, col_id, newItem)
                        
        dialog = resetSecondDiaglog()
        if dialog.exec_():
            pass

    def createTableForIsland(self,valid_date_1,valid_date_2,\
                            table_1,table_2,transport,dinnertype,childage):
        '''根据序列化缓存数据，生成具体的数据表格
        '''
        people_type_list = ['成人','儿童','婴儿'] # 设定人群类型

        # 第一页面lineEdit数据整理成规范DataFrame
        df_date_contract = DataFrame({'起始日期':[],'终止日期':[]})
        for date in valid_date_1:
            try:
                date_1, date_2 = valid_date_1[date].split('-')
                if date_1 != '' and date_2 != '':
                    df_date_contract = df_date_contract.append({'起始日期':date_1,'终止日期':date_2},ignore_index=True)
            except:
                pass

        # 第一页面tableWidget数据整理成规范DataFrame
        if len(table_1[table_1['房型名称']==''].index) != 0:
            table_1.drop(table_1[table_1['房型名称']==''].index,inplace=True)

        # 第二页面lineEdit数据整理成规范DataFrame
        df_date_package = DataFrame({'起始日期':[],'终止日期':[]})
        for date in valid_date_2:
            try:
                date_1, date_2 = valid_date_2[date].split('-')
                if date_1 != '' and date_2 != '':
                    df_date_package = df_date_package.append({'起始日期':date_1,'终止日期':date_2},ignore_index=True)
            except:
                pass

        # 第二页面tableWidget数据整理成规范DataFrame
        if len(table_2[table_2['打包组合房型全称']==''].index) != 0:
            table_2.drop(table_2[table_2['打包组合房型全称']==''].index,inplace=True)

        # 第三页面交通组合方式数据整理成规范DataFrame
        df_traffic_combination = DataFrame({'交通组合方式':[]})
        for tp in transport:
            df_traffic_combination = df_traffic_combination.append({'交通组合方式':transport[tp]},ignore_index=True)
        if len(df_traffic_combination[df_traffic_combination['交通组合方式']==''].index) != 0:
            df_traffic_combination.drop(df_traffic_combination[df_traffic_combination['交通组合方式']==''].index,inplace=True)
            df_traffic_combination = df_traffic_combination.reset_index(drop=True)

        # 第三页面餐费类型数据整理成规范DataFrame
        df_fee_type = DataFrame({'餐费类型':[]})
        for feetype in dinnertype:
            df_fee_type = df_fee_type.append({'餐费类型':dinnertype[feetype]},ignore_index=True)
        if len(df_fee_type[df_fee_type['餐费类型']==''].index) != 0:
            df_fee_type.drop(df_fee_type[df_fee_type['餐费类型']==''].index,inplace=True)
            df_fee_type = df_fee_type.reset_index(drop=True)
        
        # 第三页面儿童年龄分层数据整理成规范DataFrame
        child_age_stratification = {}
        for pt in people_type_list:
            child_age_stratification.update({pt:[]})
        age = set()
        for ca in childage:
            try:
                a,b = childage[ca].split('-')
                child_age_stratification["儿童"].append(childage[ca])
                age.add(float(a))
                age.add(float(b))
            except:
                pass
        child_age_stratification["婴儿"].append('<'+str(int(min(age))))
        child_age_stratification["成人"].append('>'+str(int(ceil(max(age)))))

        df_output_all = {}
        # 创建合同房型价格表           
        df_output = DataFrame(columns=['房型','起始日期','终止日期','最大容纳人数(无附加费)', \
                                        '最大容纳人数(有附加费)','单价'])
        for k,house_type in enumerate(table_1['房型名称']):
            sheet_1 = DataFrame(columns=['房型','起始日期','终止日期','最大容纳人数(无附加费)', \
                                        '最大容纳人数(有附加费)','单价'])
            min_num = 1
            max_num = int(table_1['最大容纳人数(无附加费)'].iloc[k])
            for i in range(df_date_contract.shape[0]):
                for j, num in enumerate(range(min_num,max_num+1)):
                    idx = i*len(range(min_num,max_num+1))+j
                    sheet_1.loc[idx,'最大容纳人数(无附加费)'] = num
                    sheet_1.loc[idx,'起始日期'] = df_date_contract['起始日期'].iloc[i]
                    sheet_1.loc[idx,'终止日期'] = df_date_contract['终止日期'].iloc[i]
                    sheet_1.loc[idx,'房型'] = house_type
                    sheet_1.loc[idx,'最大容纳人数(有附加费)'] = table_1['最大容纳人数(有附加费)'].iloc[k]
                    sheet_1.loc[idx,'单价'] = 0
            df_output = df_output.append(sheet_1)
        df_output.set_index(list(df_output.keys())[0], inplace=True)
        df_output_all.update({'合同房型价格表':df_output})

        # 创建合同第三人费用表
        df_output = DataFrame(columns=['房型名称','人群类型','年龄分层','起始日期','终止日期','单价'])
        idx = 0
        for k, house_type in enumerate(table_1['房型名称']):
            sheet_1 = DataFrame(columns=['房型名称','人群类型','年龄分层','起始日期','终止日期','单价'])
            for j, people_type in enumerate(people_type_list):
                for _, age in enumerate(child_age_stratification[people_type]):
                    for i in range(df_date_contract.shape[0]):
                        idx += i
                        sheet_1.loc[idx,'房型名称'] = house_type
                        sheet_1.loc[idx,'人群类型'] = people_type
                        sheet_1.loc[idx,'年龄分层'] = age
                        sheet_1.loc[idx,'起始日期'] = df_date_contract['起始日期'].iloc[i]
                        sheet_1.loc[idx,'终止日期'] = df_date_contract['终止日期'].iloc[i]
                        sheet_1.loc[idx,'单价'] = 0
                    idx += 1
            df_output = df_output.append(sheet_1)
        df_output.set_index(list(df_output.keys())[0], inplace=True)
        df_output_all.update({'合同第三人费用表':df_output})

        # 创建打包房型价格表
        if not df_date_package.empty:
            col_name = ['打包组合房型全称','起始日期','终止日期']
            col_name.extend(list(df_fee_type['餐费类型']))
            df_output = DataFrame(columns=col_name)
            for k,house_type in enumerate(table_2['打包组合房型全称']):
                sheet_1 = DataFrame(columns=col_name)
                for i in range(df_date_package.shape[0]):
                    idx = i+k*df_date_package.shape[0]
                    sheet_1.loc[idx,'打包组合房型全称'] = house_type
                    sheet_1.loc[idx,'起始日期'] = df_date_package['起始日期'].iloc[i]
                    sheet_1.loc[idx,'终止日期'] = df_date_package['终止日期'].iloc[i]
                    for dt in df_fee_type['餐费类型']:
                        sheet_1.loc[idx,dt] = 0
                df_output = df_output.append(sheet_1)
        else:
            col_name = ['打包组合房型全称']
            col_name.extend(list(df_fee_type['餐费类型']))
            df_output = DataFrame(columns=col_name)
            try:
                for i,house_type in enumerate(table_2['打包组合房型全称']):
                    df_output.loc[i,'打包组合房型全称'] = house_type
                    for ft in list(df_fee_type['餐费类型']):
                        df_output.loc[i,ft] = 0
            except:
                pass
        df_output.set_index(list(df_output.keys())[0], inplace=True)
        df_output_all.update({'打包房型价格表':df_output})   

        # 创建打包第三人费用表
        col_name = ['打包组合房型全称','人群类型','年龄分层','起始日期','终止日期','打包所含费用名称']; col_name_1 = []
        col_name.extend(list(df_fee_type['餐费类型'])); col_name_1.extend(list(df_fee_type['餐费类型']))
        col_name.extend(list(df_traffic_combination['交通组合方式'])); col_name_1.extend(list(df_traffic_combination['交通组合方式']))
        df_output = DataFrame(columns=col_name)
        idx = 0
        for k,house_type in enumerate(table_2['打包组合房型全称']):
            sheet_1 = DataFrame(columns=col_name)
            for j, people_type in enumerate(people_type_list):
                for _, age in enumerate(child_age_stratification[people_type]):
                    for i in range(df_date_package.shape[0]):
                        idx += i
                        sheet_1.loc[idx,'打包组合房型全称'] = house_type
                        sheet_1.loc[idx,'人群类型'] = people_type
                        sheet_1.loc[idx,'年龄分层'] = age
                        sheet_1.loc[idx,'起始日期'] = df_date_package['起始日期'].iloc[i]
                        sheet_1.loc[idx,'终止日期'] = df_date_package['终止日期'].iloc[i]
                        sheet_1.loc[idx,'打包所含费用名称'] = table_2['打包所含费用'].iloc[k]
                        for name in col_name_1:
                            sheet_1.loc[idx,name] = 0
                    idx += 1
            df_output = df_output.append(sheet_1)
        df_output.set_index(list(df_output.keys())[0], inplace=True)
        df_output_all.update({'打包第三人费用表':df_output})

        # 创建餐费价格表
        col_name = ['人群类型','年龄分层','起始日期','终止日期']
        col_name.extend(list(df_fee_type['餐费类型']))
        df_output = DataFrame(columns=col_name)
        df_output.set_index(list(df_output.keys())[0], inplace=True)
        df_output_all.update({'餐费价格表':df_output})

        # 创建交通组合价格表
        col_name = ['人群类型','年龄分层','起始日期','终止日期']
        col_name.extend(list(df_traffic_combination['交通组合方式']))
        df_output = DataFrame(columns=col_name)
        df_output.set_index(list(df_output.keys())[0], inplace=True)
        df_output_all.update({'交通组合价格表':df_output})

        # 创建住付优惠表
        col_name = ['起始日期','终止日期','住付优惠']
        df_output = DataFrame(columns=col_name)
        df_output.set_index(list(df_output.keys())[0], inplace=True)
        df_output_all.update({'住付优惠表':df_output})

        # 创建节假日附加费和强制费用表
        col_name = ['人群类型','年龄分层','起始日期','终止日期']
        col_name.extend(list(df_fee_type['餐费类型']))
        col_name.extend(['附加费'])
        df_output = DataFrame(columns=col_name)
        df_output.set_index(list(df_output.keys())[0], inplace=True)
        df_output_all.update({'节假日附加费和强制费用表':df_output})

        # 创建EBD折扣表
        col_name = ['起始日期','终止日期','提前预定天数','房费折扣','餐费折扣',\
                    '能混用的优惠方式','是否适用于第三人房费','是否适用于餐费']
        df_output = DataFrame(columns=col_name)
        df_output.set_index(list(df_output.keys())[0], inplace=True)
        df_output_all.update({'EBD折扣表':df_output})

        # 创建LSD折扣表
        col_name = ['起始日期','终止日期','入住天数','房费折扣','餐费折扣',\
                    '能混用的优惠方式','是否适用于第三人房费','是否适用于餐费']
        df_output = DataFrame(columns=col_name)
        df_output.set_index(list(df_output.keys())[0], inplace=True)
        df_output_all.update({'LSD折扣表':df_output})

        # 创建其他信息表
        df_output = DataFrame(columns=['信息名称','数值'])
        df_output.loc[0,'信息名称'] = '房价是否包含环境税'; df_output.loc[0,'数值'] = '否'
        df_output.loc[1,'信息名称'] = '交通是否有优惠'; df_output.loc[1,'数值'] = '否'
        df_output.loc[2,'信息名称'] = '儿童费用是否有优惠'; df_output.loc[2,'数值'] = '否'
        df_output.loc[3,'信息名称'] = '混住费'; df_output.loc[3,'数值'] = '0'
        df_output.loc[4,'信息名称'] = '免费夜强制费'; df_output.loc[4,'数值'] = '0'
        df_output.loc[5,'信息名称'] = '免费夜固定房费价格'; df_output.loc[5,'数值'] = '0'
        df_output.loc[6,'信息名称'] = '第三成人免费夜固定房费价格'; df_output.loc[6,'数值'] = '0'
        df_output.loc[7,'信息名称'] = '成人环境税'; df_output.loc[7,'数值'] = '6'
        df_output.loc[8,'信息名称'] = '儿童环境税'; df_output.loc[8,'数值'] = '6'
        df_output.loc[9,'信息名称'] = '婴儿环境税'; df_output.loc[9,'数值'] = '6'
        df_output.set_index(list(df_output.keys())[0], inplace=True)
        df_output_all.update({'其他信息表':df_output})

        if (self.input.text() is None) or (self.input.text() == ''):
            with open('./cache/transport_thirdpage.pkl','wb') as cache_file:
                pickle.dump(self.cache_data_set_0, cache_file)
            with open('./cache/dinnertype_thirdpage.pkl','wb') as cache_file:
                pickle.dump(self.cache_data_set_1, cache_file)
            with open('./cache/child_age_stratification.pkl','wb') as cache_file:
                pickle.dump(self.cache_data_set_2, cache_file)  
            self.msg("错误","请输入数据表名称！")
        else:
            try:
                if not os.path.exists('./data'):
                    os.mkdir('./data')
                writer = ExcelWriter(os.path.join('./data/',self.input.text()+'.xlsx'))
                for _, item in enumerate(df_output_all):
                    df_output_all[item].to_excel(writer, sheet_name=item)
                writer.save()
                writer.close()
                _ = QMessageBox.information(self,'提示','已创建'+str(self.input.text())+'.xlsx数据表！')

                # 当正式创建数据表成功后，会将cache文件夹中的备份数据全部删除掉
                if len(os.listdir('./cache/')) != 0:
                    for f in os.listdir('./cache/'):
                        os.remove(os.path.join('./cache/',f))
            except Exception as e:
                if "[Errno 13] Permission denied: './data/"+self.input.text()+".xlsx'" in str(e):
                    # 如果未创建成功，所有已输入数据会备份到cache文件夹，当再次启动'创建'按钮，会立即加载所有数据
                    with open('./cache/transport_thirdpage.pkl','wb') as cache_file:
                        pickle.dump(self.cache_data_set_0, cache_file)
                    with open('./cache/dinnertype_thirdpage.pkl','wb') as cache_file:
                        pickle.dump(self.cache_data_set_1, cache_file)
                    with open('./cache/child_age_stratification.pkl','wb') as cache_file:
                        pickle.dump(self.cache_data_set_2, cache_file)  
                    self.msg("错误","与"+self.input.text()+".xlsx重名的文件已被打开，请先关闭该文件！")

    def msg(self,string1,string2):
        _ = QMessageBox.critical(self,string1,string2)

    def center(self):
        screen = QDesktopWidget().screenGeometry()
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
        self.pushButton_4.clicked.connect(self.crtTableForIsland)
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

    def crtTableForIsland(self):
        '''修改日期：2019/01/29
           修改内容：以分步输入简要信息的方式来创建数据表
        '''
        if self.datapath == '':
            self.msg("错误","请输入数据表文件夹路径！")
        else:
            diaglog = createTableFirstDiaglog()
            if diaglog.exec_():
                pass

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