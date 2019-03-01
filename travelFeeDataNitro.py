# -*- coding: utf-8 -*-
"""
Created on Sat Dec 22 16:51:11 2018

@author: liziran (Damon Li)
"""

import datetime
import sys
import time

reload(sys)
sys.setdefaultencoding("utf8")


table_name_list = ["房型价格表","房型简称","其他单价表","客户信息数据"]

def stat_checkin_days(table_name):
	check_in_nights = 0
	for row_id in range(1,100):
		row_value = Cell(table_name,row_id,1).value
		if row_value is not None and u'房型' in row_value:
			check_in_nights += int(Cell(table_name,row_id,2).value)
		elif row_value is None:
			return check_in_nights

def get_calc_unit(table_name):
	return Cell(table_name,2,2).value

def get_start_end_date(table_name,check_in_nights):
	for row_id in range(1,100):
		row_value = Cell(table_name,row_id,1).value
		if row_value == '入住时间':
			start_date = str(Cell(table_name,row_id,2).value)
			start_date = datetime.date(int(start_date[0:4]), int(start_date[4:6]), int(start_date[6:]))
			end_date = start_date + datetime.timedelta(days=check_in_nights)
			end_date = end_date.year * 10000 + end_date.month * 100 + end_date.day
			start_date = start_date.year * 10000 + start_date.month * 100 + start_date.day
			return start_date, end_date

def get_house_abbr_names(table_name):
	house_abbr_name_list = []
	for row_id in range(1,100):
		row_value = Cell(table_name,row_id,1).value
		if row_value is not None and u'房型' in row_value and Cell(table_name,row_id,2).value != 0:
			house_abbr_name_list.append(row_value.split('入住天数')[0])
		elif row_value is None:
			return house_abbr_name_list

def get_house_whole_names(table_name,house_abbr_name_list):
	house_whole_name_list = []
	for abbr_name in house_abbr_name_list:
		for row_id in range(1,100):
			row_value = Cell(table_name,row_id,2).value
			if row_value is not None and row_value == abbr_name:
				house_whole_name_list.append(Cell(table_name,row_id,1).value)
				break
			elif row_value is None:
				break
	return house_whole_name_list

def get_people_amount(table_name):
	adult_amount = 0; child_amount = 0; baby_amount = 0
	for row_id in range(1,100):
		row_value = Cell(table_name,row_id,1).value
		if row_value == '成人人数':
			adult_amount = Cell(table_name,row_id,2).value
		elif row_value == '儿童人数':
			child_amount = Cell(table_name,row_id,2).value
		elif row_value == '婴儿人数':
			baby_amount = Cell(table_name,row_id,2).value
		elif (row_value is None) or (adult_amount != 0 and child_amount != 0 and \
			baby_amount != 0):
			return adult_amount, child_amount, baby_amount

def get_price_set(table_name,house_whole_name_list,start_date,end_date):
	price_set = {}
	for house_type in house_whole_name_list:
		for row_id in range(1,100):
			row_value = Cell(table_name,row_id,1).value
			if row_value is not None and row_value == house_type and \
			Cell(table_name,row_id,2).value <= start_date <= Cell(table_name,row_id,3).value and \
			Cell(table_name,row_id,2).value <= end_date <= Cell(table_name,row_id,3).value and \
			Cell(table_name,row_id,4).value == adult_amount:
				price_set.update({house_type:int(Cell(table_name,row_id,5).value)})
			elif row_value is None:
				break
	return price_set

def get_days_set(table_name,house_whole_name_list,house_abbr_name_list):
	days_set = {}
	for Id, house_type in enumerate(house_abbr_name_list):
		for row_id in range(1,100):
			row_value = Cell(table_name,row_id,1).value
			if row_value is not None and house_type in row_value:
				days_set.update({house_whole_name_list[Id]:Cell(table_name,row_id,2).value})
				break
	return days_set

def get_free_house_days(table_name):
	for row_id in range(1,100):
		row_value = Cell(table_name,row_id,1).value
		if row_value == '免房费天数':
			return Cell(table_name,row_id,2).value

def get_traffic_price(table_name):
	adult_traffic_price = 0
	child_traffic_price = 0
	for row_id in range(1,100):
		row_value = Cell(table_name,row_id,1).value
		if row_value == '成人交通费':
			adult_traffic_price = Cell(table_name,row_id,3).value
		elif row_value == '儿童交通费':
			child_traffic_price = Cell(table_name,row_id,3).value
		elif (row_value is None) or (adult_traffic_price != 0 and child_traffic_price != 0):
			return adult_traffic_price, child_traffic_price

def get_env_tax_price(table_name):
	adult_env_tax_price = 0
	child_env_tax_price = 0
	for row_id in range(1,100):
		row_value = Cell(table_name,row_id,1).value
		if row_value == '成人环境税':
			adult_env_tax_price = Cell(table_name,row_id,3).value
		elif row_value == '儿童环境税':
			child_env_tax_price = Cell(table_name,row_id,3).value
		elif (row_value is None) or (adult_env_tax_price != 0 and child_env_tax_price != 0):
			return adult_env_tax_price, child_env_tax_price

def get_free_night_food_price(table_name):
	for row_id in range(1,100):
		row_value = Cell(table_name,row_id,1).value
		if row_value == '免费夜餐费':
			return Cell(table_name,row_id,3).value

def get_third_person_fee(table_name):
	third_person_price_child = 0
	third_person_price_baby = 0
	for row_id in range(1,100):
		row_value = Cell(table_name,row_id,1).value
		if row_value == '儿童第三人费用':
			third_person_price_child = Cell(table_name,row_id,3).value
		elif row_value == '婴儿第三人费用':
			third_person_price_baby = Cell(table_name,row_id,3).value
		elif (row_value is None) or (third_person_price_child != 0 and third_person_price_baby != 0):
			return third_person_price_child, third_person_price_baby

if __name__ == '__main__':
	t1 = time.time()
	check_in_nights = stat_checkin_days(table_name_list[-1])
	cal_unit = get_calc_unit(table_name_list[2])
	start_date, end_date = get_start_end_date(table_name_list[-1],check_in_nights)
	house_abbr_name_list = get_house_abbr_names(table_name_list[-1])
	house_whole_name_list = get_house_whole_names(table_name_list[1],house_abbr_name_list)
	adult_amount, child_amount, baby_amount = get_people_amount(table_name_list[-1])
	price_set = get_price_set(table_name_list[0],house_whole_name_list,start_date,end_date)
	days_set = get_days_set(table_name_list[-1],house_whole_name_list,house_abbr_name_list)
	free_house_days = get_free_house_days(table_name_list[-1])
	adult_traffic_price, child_traffic_price = get_traffic_price(table_name_list[2])
	adult_env_tax_price, child_env_tax_price = get_env_tax_price(table_name_list[2])
	food_price = get_free_night_food_price(table_name_list[2])
	third_person_price_child, third_person_price_baby = get_third_person_fee(table_name_list[2])

	house_cost = 0
	string_outprint = ' - 房费：'
	for Id, (house_type, house_price) in enumerate(price_set.items()):
		house_cost += house_price * days_set[house_type]
		if Id < len(price_set) - 1:
			string_outprint += house_type + '(' + str(house_price) + ') * ' +  str(days_set[house_type]) + '(天) + '
		else:
			string_outprint += house_type + '(' + str(house_price) + ') * ' +  str(days_set[house_type]) + '(天)'
	if free_house_days != 0:
		house_cost -= free_house_days * price_set[min(price_set, key=price_set.get)]
		string_outprint += ' - ' + '最低房费单价(' + str(price_set[min(price_set, key=price_set.get)]) + ') * 免房费天数(' + str(free_house_days) + '天) = ' + str(house_cost)\
						+ ' ' + cal_unit
	string_outprint += '\n'

	string_outprint += ' - 交通费：'
	traffic_cost = adult_traffic_price * adult_amount + child_traffic_price * child_amount
	string_outprint += '成人交通费单价(' + str(adult_traffic_price) + ') * ' + '成人数(' + str(adult_amount) + ') + ' + \
					'儿童交通费单价(' + str(child_traffic_price) + ') * ' + '儿童数(' + str(child_amount) + ') = ' + \
					str(traffic_cost) +  ' ' + cal_unit
	string_outprint += '\n'

	string_outprint += ' - 环境税费：'
	env_tax_cost = (adult_env_tax_price * adult_amount + child_env_tax_price * child_amount) * check_in_nights
	string_outprint += '成人环境税费单价(' + str(adult_env_tax_price) + ') * ' + '成人数(' + str(adult_amount) + ') + ' + \
					'儿童环境税费单价(' + str(child_env_tax_price) + ') * ' + '儿童数(' + str(child_amount) + ') = ' + \
					str(env_tax_cost) +  ' ' + cal_unit
	string_outprint += '\n'

	string_outprint += ' - 免费夜餐费：'
	free_night_food_cost = food_price * adult_amount
	string_outprint += '成人免费夜餐费单价(' + str(food_price) + ') * ' + '成人数(' + str(adult_amount) + ') = ' + \
					str(free_night_food_cost) +  ' ' + cal_unit
	string_outprint += '\n'

	string_outprint += ' - 第三人费用：'
	third_person_cost = (third_person_price_child * child_amount + third_person_price_baby * baby_amount) * check_in_nights
	string_outprint += '[ 儿童第三人费用单价(' + str(third_person_price_child) + ') * ' + '儿童数(' + str(child_amount) + ') + ' + \
					'婴儿第三人费用单价(' + str(third_person_price_baby) + ') * ' + '婴儿数(' + str(baby_amount) + ') ] * 入住天数(' + \
					str(check_in_nights) + ') = ' + str(third_person_cost) +  ' ' + cal_unit
	string_outprint += '\n'  

	whole_cost =  house_cost + traffic_cost + env_tax_cost + free_night_food_cost + third_person_cost

	string_outprint += " * 实际费用 = 房费 + 交通费 + 环境税费 + 免费夜餐费 + 第三人费用 = " + str(whole_cost)

	Cell("C2").value = string_outprint
	t2 = time.time()
	Cell("D2").value = "共用了" + str(round(t2 - t1,2)) + "s"