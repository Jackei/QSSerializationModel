#coding:utf-8

import xlrd
import os
import time

Father = 'BaseResponseEntity'
ISOTIMEFORMAT='%Y-%m-%d %X'

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
				creatFile(info)
				startIndex = index+1
				info.propertyDict.clear()
				info.propertyFatherDict.clear()
				for i in range(len(info.importName)):
					info.importName.pop()


def creatFile(info):
	if (os.path.exists('./file/'+info.name+'.java')):
		os.popen('rm -rf '+'./file'+info.name+'.java')
	f = file('./file/'+info.name+'.java','w+')
	f.write('public class '+info.name+' extends '+Father+' {'+'\r\n')
	f.write('\r\n')
	creatProperty(info,f)
	f.write('}'+'\r\n')
	f.close()


def creatProperty(info,f):
	for key in info.propertyDict:
		value = info.propertyDict[key]
		if (value == 'string'):
			f.write('	private String '+key+';'+'\r\n')
		elif (value == '[]'):
			f.write('	private List<'+info.propertyFatherDict[key]+'> '+key+';'+'\r\n')
		elif (value == 'bool' or value == 'int' or value == 'float' or value == 'double'):
			specialValue = ''
			if (value == 'bool'):
				specialValue == 'boolean'
			else:
				specialValue == value
			f.write('	private '+value+' '+key+';'+'\r\n')
		else:
			f.write('	private '+info.propertyFatherDict[key]+' '+key+';'+'\r\n')
	f.write('\r\n')

	for key in info.propertyDict:
		value = info.propertyDict[key]
		print(value)
		specialValue = ''
		if (value == 'bool'):
			specialValue = 'boolean'
		elif (value == '[]'):
			specialValue = 'List<'+info.propertyFatherDict[key]+'>'
		elif (value == 'string'):
			specialValue = 'String'
		else:
			specialValue = value
		f.write('	public '+specialValue+' '+'get'+key.capitalize()+'()'+' { return '+key+'; }'+'\r\n')		
		f.write('\r\n')

		spe = ''
		if (value == '[]'):
			spe = 'List<'+info.propertyFatherDict[key]+'> '+key
		else:
			spe = specialValue+' '+key
		f.write('	public void set'+key.capitalize()+'('+spe+')'+' {'+'\r\n')
		f.write('		this.'+key+' = '+key+';'+'\r\n')
		f.write('	}'+'\r\n')
		f.write('\r\n')
	

def creatFolder():
	if (os.path.exists('./file')):
		os.popen('rm -rf ./file')
	os.mkdir('./file')


if __name__=='__main__':
	creatFolder()
	getData()
