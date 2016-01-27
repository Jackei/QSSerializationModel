#coding:utf-8

import xlrd
import os
import time

Father = 'JSONModel'
ISOTIMEFORMAT='%Y-%m-%d %X'
JSONType = ['string','[]','int','float','double','bool']


class InterfaceInfo:
	name = ''
	propertyDict = {}
	propertyFatherDict = {}
	isProtocol = False
	importName = []


def getData():
	data = xlrd.open_workbook('接口文档.xlsx')
	for table in data.sheets():
		if (table.name == '目录'.decode('utf-8')):
			continue
		col1 = table.col_values(0)
		startIndex = 0
		endIndex = 0
		first = True
		for index in range(0,len(col1)):
			if (table.cell(index,0).value == 'output'):
				startIndex = index
			if (table.cell(index,0).value == '#'):
				endIndex = index
				info = InterfaceInfo()
				if first:
					first = False
					info.name = table.cell(2,0).value
					info.isProtocol = False
				else:
					info.name = table.cell(startIndex,0).value
					info.isProtocol = True
				for i in range(startIndex,endIndex):
					info.propertyDict[table.cell(i,1).value] = table.cell(i,2).value
					info.propertyFatherDict[table.cell(i,1).value] = table.cell(i,3).value
					info.importName.append(table.cell(i,3).value)
				creatHfile(info)
				creatMfile(info)
				startIndex = index+1


def creatHfile(info):
	if (os.path.exists('./file/'+info.name+'.h')):
		os.popen('rm -rf '+'./file'+info.name+'.h')
	f = file('./file/'+info.name+'.h','w+')
	f.write('//'+'\r\n')
	f.write('// '+info.name+'.h'+'\r\n')
	f.write('// '+'SerializationModel 自动生成'+'\r\n')
	f.write('//'+'\r\n')
	f.write('// Create by qizhijian on '+time.strftime(ISOTIMEFORMAT,time.localtime())+'\r\n')
	f.write('// Copyright © 2015年 qizhijian. All rights reserved.'+'\r\n')
	f.write('//'+'\r\n')
	f.write('\r\n')
	f.write('#import <Foundation/Foundation.h>'+'\r\n')
	f.write('#import "JSONModel.h"'+'\r\n')
	for n in info.importName:
		if (n == ''):
			continue
		else:
			f.write('#import "'+n+'.h"'+'\r\n')
	if (info.isProtocol):
		f.write('\r\n')
		f.write('@protocol '+info.name+'\r\n')
		f.write('@end'+'\r\n')
	f.write('\r\n')
	f.write('@interface '+info.name+' : '+Father+'\r\n')
	f.write('\r\n')
	creatHProperty(info,f)
	f.write('@end'+'\r\n')
	f.close()


def creatHProperty(info,f):
	for key in info.propertyDict:
		value = info.propertyDict[key]
		f.write('@property (nonatomic')
		if (value == 'string'):
			f.write(', strong) NSString <Optional> *'+key+';'+'\r\n')
		elif (value == '[]'):
			f.write(', strong) NSMutableArray <'+info.propertyFatherDict[key]+',Optional> *'+key+';'+'\r\n')
			f.write
		elif (value == 'bool' or value == 'int' or value == 'float' or value == 'double'):
			specialValue = ''
			if (value == 'bool'): 
				specialValue = 'BOOL' 
			else: 
				specialValue = value
			f.write(') '+specialValue+' '+key+';'+'\r\n')
		else:
			f.write(', strong) '+value+' <Optional> *'+key+';'+'\r\n')
	f.write('\r\n')


def creatMfile(info):
	if (os.path.exists('./file/'+info.name+'.m')):
		os.popen('rm -rf '+'./file'+info.name+'.m')
	f = file('./file/'+info.name+'.m','w+')
	f.write('//'+'\r\n')
	f.write('// '+info.name+'.m'+'\r\n')
	f.write('// '+'SerializationModel 自动生成'+'\r\n')
	f.write('//'+'\r\n')
	f.write('// Create by qizhijian on '+time.strftime(ISOTIMEFORMAT,time.localtime())+'\r\n')
	f.write('// Copyright © 2015年 qizhijian. All rights reserved.'+'\r\n')
	f.write('//'+'\r\n')
	f.write('\r\n')
	f.write('#import '+'"'+info.name+'.h'+'"'+'\r\n')
	f.write('\r\n')
	f.write('@implementation '+info.name+'\r\n')
	f.write('\r\n')
	f.write('@end'+'\r\n')
	f.close()


def creatFolder():
	if (os.path.exists('./file')):
		os.popen('rm -rf ./file')
	os.mkdir('./file')


if __name__=='__main__':
	creatFolder()
	getData()
