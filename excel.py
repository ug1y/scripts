# -*- coding:utf-8 -*-
'''
Created on 2016-9-19

@author: yinhao
'''
import xml.dom.minidom, xlsxwriter, datetime, sys, getopt

class XmlDom():
	# 读取指定格式的xml文档
	def __init__(self, fid):
		self.fid = fid
		self.sheet = None
		self.parseRoot()
	# 打卡的xml用的是GB2312，重新编码成UTF-8
	def transcode(self):
		try:
			content = open(self.fid,"r").read()
		except IOError, e:
			print e
			sys.exit()
		else:
			if content.find('<?xml version="1.0" encoding="GB2312" ?>')!=-1:
				content = content.replace('<?xml version="1.0" encoding="GB2312" ?>','<?xml version="1.0" encoding="UTF-8" ?>')
				content = unicode(content,encoding='gb2312').encode('utf-8')
			return content
	# 解析xml文档生成dom树，选择“刷卡记录”表
	def parseRoot(self):
		f = self.transcode()
		dom = xml.dom.minidom.parseString(f)
		root = dom.documentElement
		e = root.getElementsByTagName('Worksheet')
		for i in e:
			if i.getAttribute('ss:Name') == '刷卡记录'.decode('utf-8'):
				self.sheet = i
				return 
	# 返回节点的值，若无则返回空字符
	def getText(self, node):
		return node.nodeValue if node != None else ''
	# 将xml格式的表转换成py中的二位数组处理，即excel中的可视化二维矩阵
	def getTable(self):
		ans = []
		table = self.sheet.getElementsByTagName('Table')[0]
		for row in table.getElementsByTagName('Row'):
			tmp = []
			for cell in row.getElementsByTagName('Cell'):
				tmp.append(self.getText(cell.getElementsByTagName('Data')[0].firstChild) if cell.hasChildNodes() else '')
			ans.append(tmp)
		return ans

# 获取表的属性值时间
def getTableTime(table):
	return str(table[2][1].split('~')[0].strip()).split('/')[0:1] +\
	str(table[2][1].split('~')[1].strip()).split('/')

# 按姓名获取对应的一行打卡时间
def getTableGuy(table, name):
	table = table[4:]
	for i in range(0,len(table),2):
		if name.decode('utf-8') in table[i]:
			return table[i+1]
	else:
		return 

# 按工号获取对应的一行打卡时间
def getTableId(table, id):
	table = table[4:]
	for i in range(0,len(table),2):
		if str(table[i][2]) == str(id):
			return table[i+1]
	else:
		return 

# 调整时间为上午下午的标准格式
def adjust(time):
	if len(time) == 0:
		return ['','','','']
	point = '13:00'
	before = [t for t in time if t <= point]
	after = [t for t in time if t > point]
	# 调整上午的时间，若只有一次打卡，则在11点30之前的记作开始，在11点30之后的记作离开
	if len(before) == 0:
		before = ['','']
	elif len(before) == 1:
		before = [before[0],'11:30'] if before[0] < '11:30' else ['08:30',before[0]]
	else:
		before = [before[0],before[-1]]
	# 调整下午的时间，若只有一次打卡，则在17点30之前的记作开始，在17点30之后的记作离开
	if len(after) == 0:
		after = ['','']
	elif len(after) == 1:
		after = [after[0],'17:30'] if after[0] < '17:30' else ['14:30',after[0]]
	else:
		after = [after[0],after[-1]]
	# 其他时间只算首尾
	return before + after

# 一行打卡时间转换成标准的四个时间点的格式
def transform(row):
	return [adjust([str(cell)[i*5:(i+1)*5] for i in range(len(str(cell))/5)] if cell != '' else '') for cell in row] if row != None else None


class XlsxMake():
	# 制作xlsx文档，传入文件名
	def __init__(self, fout):
		self.fout = fout
		self.book = None
		self.openXlsx()
	# 打开一个新的excel文档	
	def openXlsx(self):
		self.book = xlsxwriter.Workbook(self.fout)
	# 字符串返回格式化日期
	def getDate(self, string):
		return datetime.datetime.strptime(string, '%Y-%m-%d') if string != '' else ''
	# 字符串返回格式化时间
	def getTime(self, string):
		return datetime.datetime.strptime(string, '%H:%M') if string != '' else ''
	# 日期返回星期
	def getWeek(self, date):
		week_day_dict = {
		0 : '星期一',
		1 : '星期二',
		2 : '星期三',
		3 : '星期四',
		4 : '星期五',
		5 : '星期六',
		6 : '星期日',
		}
		day = date.weekday()
		return week_day_dict[day]
	# 制作指定格式的表格内容
	def makeSheet(self, title, row):
		sheet = self.book.add_worksheet("-".join(title))	# 新添加一个sheet
		# 按照规定格式制作表头
		cell_format = self.book.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
		sheet.merge_range('A1:A2','日期'.decode('utf-8'), cell_format)
		sheet.set_column('A:A',11)
		sheet.merge_range('B1:B2','星期'.decode('utf-8'), cell_format)
		sheet.set_column('B:B',11)
		sheet.merge_range('C1:D1','时段一'.decode('utf-8'), cell_format)
		sheet.merge_range('E1:F1','时段二'.decode('utf-8'), cell_format)
		sheet.write('C2','签到'.decode('utf-8'), cell_format)
		sheet.write('D2','签退'.decode('utf-8'), cell_format)
		sheet.write('E2','签到'.decode('utf-8'), cell_format)
		sheet.write('F2','签退'.decode('utf-8'), cell_format)
		sheet.merge_range('G1:G2','工作时间'.decode('utf-8'), cell_format)
		sheet.merge_range('H1:J2','说明'.decode('utf-8'), cell_format)
		sheet.set_column('C:J',9)
		sheet.merge_range('K1:K2','总时间'.decode('utf-8'), cell_format)
		sheet.set_column('K:K',15)
		# 日期格式，星期六和星期日用特别颜色标注出来
		date_format = self.book.add_format({'num_format': 'yyyy/mm/dd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
		week_format = self.book.add_format({'num_format': '[hh]:mm', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#E1E1E1'})
		day_format = self.book.add_format({'num_format': '[hh]:mm', 'align': 'center', 'valign': 'vcenter', 'border': 1})
		# 添加表格内容
		j=0
		for i in range(int(title[2])):
			time = self.getDate(title[0]+'-'+title[1]+'-'+str(1+i))
			sheet.write_datetime('A'+str(3+i), time, date_format)
			time_format = week_format if time.weekday() in [5,6] else day_format
			sheet.write('B'+str(3+i), self.getWeek(time).decode('utf-8'), time_format)
			sheet.write('C'+str(3+i), self.getTime(row[i][0]), time_format)
			sheet.write('D'+str(3+i), self.getTime(row[i][1]), time_format)
			sheet.write('E'+str(3+i), self.getTime(row[i][2]), time_format)
			sheet.write('F'+str(3+i), self.getTime(row[i][3]), time_format)
			sheet.write_formula('G'+str(3+i),'=IF(AND(ISNUMBER(C{num}),ISNUMBER(D{num})),D{num}-C{num},0)+IF(AND(ISNUMBER(F{num}),ISNUMBER(E{num})),F{num}-E{num},0)+COUNTIF(C{num}:F{num},"上课")*1/12'.format(num=str(3+i)).decode('utf-8'), time_format)
			if time.weekday()==6:
				sheet.merge_range('H{}:J{}'.format(3+j,3+i), '', cell_format)
				sheet.merge_range('K{}:K{}'.format(3+j,3+i), '=SUM(G{}:G{})'.format(3+j,3+i), day_format)
				j=i+1
		else:
			if j < i:
				sheet.merge_range('H{}:J{}'.format(3+j,3+i), '', cell_format)
				sheet.merge_range('K{}:K{}'.format(3+j,3+i), '=SUM(G{}:G{})'.format(3+j,3+i), day_format)
	# 保存并关闭文件
	def saveXlsx(self):
		try:
			self.book.close()
		except Exception, e:
			raise e
			sys.exit()

def usage():
	print '\
change record from xml into xlsx. author by ug1y\n\
\n\
用法：\n\
	python excel.py -i <infile> -o <outfile> -u <number>\n\
	python excel.py -i <infile> -o <outfile> -n <name>\n\
	python excel.py -c <char> -i <infile1><char><infile2> -o <outfile> -n <name>\n\
\n\
参数：\n\
	-h, --help			帮助\n\
	-i, --input			输入文件名[可以加多个，用选项-c的分隔符分开]\n\
	-o, --output			输出文件名，默认为"demo.xlsx"\n\
	-n, --name			指定导用户名导出对应的打卡时间\n\
	-u, --uid			指定工号导出对应的打卡时间\n\
'

def function():
	pass

if __name__ == '__main__':
	# 参数
	fid = ''
	fout = 'demo.xlsx'
	uid = -1
	uname = ''
	cut = ''
	# 读取参数
	try:
		options,args = getopt.getopt(sys.argv[1:],"hc:i:o:n:u:",["help","cut=","input=","output=","name=","uid="])
	except Exception, e:
		raise e
		sys.exit()
	# 
	for opt, val in options:
		if opt in ("-h","--help"):
			usage()
			sys.exit()
		if opt in ("-c","--cut"):
			cut = val
		if opt in ("-i","--input"):
			fid = val
		if opt in ("-o","--output"):
			fout = val
		if opt in ("-n","--name"):
			uname = val
		if opt in ("-u","--uid"):
			uid = int(val)
	# 输入验证
	if fid=='':
		print '[error!]没有输入文件，请输入-h参数查看帮助'
		sys.exit()
	elif uname=='' and uid==-1:
		print '[error!]没有指定用户，请输入-h参数查看帮助'
		sys.exit()
	# 分割文件
	files = fid.split(cut) if cut!='' else [fid]
	# 写入xlsx文件sheet中
	xm = XlsxMake(fout)
	for f in files:
		xd = XmlDom(f)
		table = xd.getTable()
		title = getTableTime(table)
		row = transform(getTableId(table, uid)) if uid!=-1 else transform(getTableGuy(table, uname))
		xm.makeSheet(title, row)
	xm.saveXlsx()
	print '***finished !***'