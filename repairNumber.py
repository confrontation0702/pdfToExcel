# -*- coding: utf-8 -*-
import json
import time

def creatNo(customer_id): 
	customers = []
	with open("./config/customers.json",'rb') as cs:
		dicts = json.load(cs)
		customers = dicts['customers']
    #获取客户缩写
	c_id= customer_id.replace('.','')
	abbr = ''
	print(c_id)
	filterList = list(c for c in customers if c['id'] == c_id)

	if len(filterList) == 0:
		return 'id is not find'
	else :
		abbr = filterList[0]['abbr']
	#获取当前日期年月
	date = time.strftime('%Y%m',time.localtime(time.time()))
	#获取批次
	batch = ''
	with open("./config/batch",'r') as b:
		batch = b.read()
	serial = ''
	with open("./config/serialNumber",'rb') as s:
		#读取序号并加1
		tmp = s.read()
		serial = ('%d' % (int(tmp)+1))
	with open("./config/serialNumber",'w') as s:
		#保存序号
		s.write(serial)
	n = (abbr+'-'+date+batch+'-'+serial)
	print(n)
	return n
