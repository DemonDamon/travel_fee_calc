# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'travel_fee_calc_form.ui'
#
# Created by: PyQt5 UI code generator 5.9.2 
#
# Author: Damon Li
# 
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import os, sys
from PyQt5.QtWidgets import QMessageBox, QFileDialog
from pandas import read_excel, ExcelFile
from datetime import date, timedelta
from interval import Interval

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(397, 520)
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(70, -10, 271, 91))
        self.label.setStyleSheet("font: 75 28pt \"Adobe Arabic\";")
        self.label.setObjectName("label")
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(320, 70, 51, 41))
        self.pushButton.setObjectName("pushButton")
        self.textBrowser = QtWidgets.QTextBrowser(Form)
        self.textBrowser.setGeometry(QtCore.QRect(30, 130, 341, 341))
        self.textBrowser.setObjectName("textBrowser")
        self.lineEdit = QtWidgets.QLineEdit(Form)
        self.lineEdit.setGeometry(QtCore.QRect(30, 70, 291, 41))
        self.lineEdit.setObjectName("lineEdit")
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        self.pushButton_2.setGeometry(QtCore.QRect(300, 480, 71, 31))
        self.pushButton_2.setObjectName("pushButton_2")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label.setText(_translate("Form", "旅游计费小工具"))
        self.pushButton.setText(_translate("Form", "Search"))
        self.pushButton_2.setText(_translate("Form", "Output"))

class mywindow(QtWidgets.QWidget,Ui_Form):    
    def __init__(self):
        super(mywindow,self).__init__() 
        self.setupUi(self)
        self.setWindowTitle("Travel Fee Calculator")
        self.pushButton.clicked.connect(self.openDataFile)
        self.lineEdit.textChanged.connect(self.leAct) 
        self.pushButton_2.clicked.connect(self.travel_fee_calc_func)
        
    def openDataFile(self):
        filename, _ = QFileDialog.getOpenFileName(self, 'Open File', os.getenv('HOME'))  
        if filename:
           self.lineEdit.setText(filename)

    def msg(self,string1,string2):
        reply = QMessageBox.critical(self,string1,string2)

    def closeEvent(self, event):
        reply = QMessageBox.question(self, "Message", "Are you sure to quit?", QMessageBox.Yes |
                                     QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def leAct(self):
        if len(self.lineEdit.text()) != '':
            self.datapath = self.lineEdit.text()

    def travel_fee_calc_func(self):
        self.textBrowser.clear()
        self.string_outprint = ''
        sheet_names = ExcelFile(self.datapath).sheet_names  # see all sheet names
        table_set = {}
        for sn in sheet_names:
            table_set.update({sn:read_excel(self.datapath, sheet_name=sn)})

        check_in_nights = int(sum([table_set['客户信息数据']['数值'][Id] for Id, string in \
                         enumerate(table_set['客户信息数据']['名称']) if '房型' in string]))

        cal_unit = table_set['其他单价表']['单位'].iloc[0]

        start_date = str(table_set['客户信息数据'][table_set['客户信息数据']['名称'] == '入住时间']['数值'].values[0])
        start_date = date(int(start_date[0:4]), int(start_date[4:6]), int(start_date[6:]))
        end_date = start_date + timedelta(days=check_in_nights)
        end_date = end_date.year * 10000 + end_date.month * 100 + end_date.day
        start_date = start_date.year * 10000 + start_date.month * 100 + start_date.day

        house_abbr_name_list = [string.split('入住天数')[0] for Id, string in enumerate(table_set['客户信息数据']['名称']) \
                           if '房型' in string and table_set['客户信息数据']['数值'].iloc[Id] != 0]

        house_whole_name_list = [table_set['房型简称'][table_set['房型简称']['房型简称'] == abbr_name]['房型全称'].values[0] \
                                 for abbr_name in house_abbr_name_list]

        adult_amount = table_set['客户信息数据'][table_set['客户信息数据']['名称'] == '成人人数']['数值'].values[0]
        child_amount = table_set['客户信息数据'][table_set['客户信息数据']['名称'] == '儿童人数']['数值'].values[0]
        baby_amount = table_set['客户信息数据'][table_set['客户信息数据']['名称'] == '婴儿人数']['数值'].values[0]

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
            price_set.update({house_type:selectby_house_type[selectby_house_type['人数'] == adult_amount]['单价'].values[0]})

        days_set = {}
        for Id, house_type in enumerate(house_whole_name_list):
            selectby_house_type = table_set['客户信息数据'][table_set['客户信息数据']['名称'].str.contains(house_abbr_name_list[Id])==True]
            days_set.update({house_type:selectby_house_type['数值'].values[0]})

        house_cost = 0
        self.string_outprint += ' - 房费：'
        for Id, (house_type, house_price) in enumerate(price_set.items()):
            house_cost += house_price * days_set[house_type]
            if Id < len(price_set) - 1:
                self.string_outprint += house_type + '(' + str(house_price) + ') * ' +  str(days_set[house_type]) + '(天) + '
            else:
                self.string_outprint += house_type + '(' + str(house_price) + ') * ' +  str(days_set[house_type]) + '(天)'
        free_house_days = table_set['客户信息数据'][table_set['客户信息数据']['名称'] == '免房费天数']['数值'].values[0]
        if free_house_days != 0:
            house_cost -= free_house_days * price_set[min(price_set)]
            self.string_outprint += ' - ' + '最低房费单价(' + str(price_set[min(price_set)]) + ') * 免房费天数(' + str(free_house_days) + '天) = ' + str(house_cost)\
                            + ' ' + cal_unit + '\n'*2

        self.string_outprint += ' - 交通费：'
        adult_traffic_price = table_set['其他单价表'][table_set['其他单价表']['名称'] == '成人交通费']['价格'].values[0]
        child_traffic_price = table_set['其他单价表'][table_set['其他单价表']['名称'] == '儿童交通费']['价格'].values[0]
        traffic_cost = adult_traffic_price * adult_amount + child_traffic_price * child_amount
        self.string_outprint += '成人交通费单价(' + str(adult_traffic_price) + ') * ' + '成人数(' + str(adult_amount) + ') + ' + \
                        '儿童交通费单价(' + str(child_traffic_price) + ') * ' + '儿童数(' + str(child_amount) + ') = ' + \
                        str(traffic_cost) +  ' ' + cal_unit + '\n'*2

        self.string_outprint += ' - 环境税费：'
        adult_env_tax_price = table_set['其他单价表'][table_set['其他单价表']['名称'] == '成人环境税']['价格'].values[0]
        child_env_tax_price = table_set['其他单价表'][table_set['其他单价表']['名称'] == '儿童环境税']['价格'].values[0]
        env_tax_cost = (adult_env_tax_price * adult_amount + child_env_tax_price * child_amount) * check_in_nights
        self.string_outprint += '成人环境税费单价(' + str(adult_env_tax_price) + ') * ' + '成人数(' + str(adult_amount) + ') + ' + \
                        '儿童环境税费单价(' + str(child_env_tax_price) + ') * ' + '儿童数(' + str(child_amount) + ') = ' + \
                        str(env_tax_cost) +  ' ' + cal_unit + '\n'*2

        self.string_outprint += ' - 免费夜餐费：'
        food_price = table_set['其他单价表'][table_set['其他单价表']['名称'] == '免费夜餐费']['价格'].values[0]
        free_night_food_cost = food_price * adult_amount
        self.string_outprint += '成人免费夜餐费单价(' + str(food_price) + ') * ' + '成人数(' + str(adult_amount) + ') = ' + \
                        str(free_night_food_cost) +  ' ' + cal_unit + '\n'*2

        self.string_outprint += ' - 第三人费用：'
        third_person_price_child = table_set['其他单价表'][table_set['其他单价表']['名称'] == '儿童第三人费用']['价格'].values[0]
        third_person_price_baby = table_set['其他单价表'][table_set['其他单价表']['名称'] == '婴儿第三人费用']['价格'].values[0]
        third_person_cost = (third_person_price_child * child_amount + third_person_price_baby * baby_amount) * check_in_nights
        self.string_outprint += '[ 儿童第三人费用单价(' + str(third_person_price_child) + ') * ' + '儿童数(' + str(child_amount) + ') + ' + \
                        '婴儿第三人费用单价(' + str(third_person_price_baby) + ') * ' + '婴儿数(' + str(baby_amount) + ') ] * 入住天数(' + \
                        str(check_in_nights) + ') = ' + str(third_person_cost) +  ' ' + cal_unit + '\n'*2
                    
        whole_cost =  house_cost + traffic_cost + env_tax_cost + free_night_food_cost + third_person_cost

        self.string_outprint += " * 实际费用 = 房费 + 交通费 + 环境税费 + 免费夜餐费 + 第三人费用 = " + str(whole_cost) + ' ' + cal_unit
        self.textBrowser.append(self.string_outprint)
            

if __name__ == '__main__':    
    app = QtWidgets.QApplication(sys.argv)
    window = mywindow()
    window.show()
    sys.exit(app.exec_())