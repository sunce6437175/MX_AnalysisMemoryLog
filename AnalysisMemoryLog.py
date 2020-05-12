#!/usr/bin/python
# -*- coding: UTF-8 -*-

from __future__ import print_function
import os
import xlwt
import xlrd
from xlrd import xldate_as_tuple
import datetime
import json
import requests
import codecs
from bs4 import BeautifulSoup

# 搜索内存关键字
deWeightTime = "top -m | grep MXNavi"
# 原有setup配置项
setupFileName = "SETUP.txt"
# 最新批量json配置项
configJsonName = "config/config.json"

# 原有setup配置项地址获取(当前工作路径)
setupFilePath = os.path.join(os.getcwd(),setupFileName).replace("\\",'/')

# 最新批量json配置项地址获取(获取当前文件的绝对路径)
configJsonPath = os.path.join(os.path.abspath(os.path.dirname(__file__)),configJsonName).replace("\\",'/')

# 读取key和设置文档的类(不适用可以删除)
class MemorylogManager():
    def __init__(self,memoryKeyword,setupFilePath):
        self.memoryKeyword = memoryKeyword
        self.setupFilePath = setupFilePath

# 读取setup设置项参数(不适用可以删除)
def setup_Working_Directory(self):
    with open(self.setupFilePath) as read_file:
        for line in read_file:
            args = line.strip().split(',')
            startTime = args[0]
            finishTime = args[1]
            readFileName = args[2]
            writeFileName = args[3]
            exclFileName = args[4]
            
            readPath = os.path.join(os.getcwd(),readFileName).replace("\\",'/')
            writePath = os.path.join(os.getcwd(),writeFileName).replace("\\",'/')
            exclPath = os.path.join(os.getcwd(),exclFileName).replace("\\",'/')
    return(startTime,finishTime,readPath,writeFileName,exclPath)

#判断有效运行时间
def check_memoryTime(startTime,finishTime,readPath,writePath,exclPath):
    startLineNum = 0
    finishLineNum = 0
    tempLienNum = 0
    tempList = []
    templine = 0
    tempNum = 0
    kPI = 800
#     # 由字符串格式转化为日期格式的函数为: datetime.datetime.strptime()
    vSt = datetime.datetime.strptime(startTime.replace("/",'-'),"%Y-%m-%d %H:%M:%S")
    vFt = datetime.datetime.strptime(finishTime.replace("/",'-'),"%Y-%m-%d %H:%M:%S")
#     # print("开始时间：%s 和对应格式 %s"%(vSt,type(vSt)))
#     # print("结束时间：%s 和对应格式 %s"%(vSt,type(vFt)))
#     # 2.1）抽取开始和结束的时间戳，判断有效运行时间（单位：小时）
    dayCtimes = ((vSt - vFt).seconds)/3600
    print("本次有效时间-----共%.2f小时-----"%(dayCtimes))

#     # 2.2）以及当内存超1024MB时所需时间。（超1024MB才需判断）
    with open(readPath,'r',encoding = "UTF-8",errors = "ignore") as read_file:
        for line in readPath:
            tempLienNum = tempLienNum + 1
            if startTime in line:
                if deWeightTime in line:
                    continue
                else:
                    startLineNum = tempLienNum
            if finishTime in line:
                if "END" in line:
                    continue
                else:
                    finishLineNum = tempLienNum
            if startLineNum <= finishLineNum and finishLineNum > startLineNum:
                with open(readPath,'r',encoding='UTF-8',errors="ignore") as read_file:
                    for xline in read_file:
                        xline = xline.strip('\n')
                        if tempNum >= startLineNum and tempNum < finishLineNum:
                            a = xline.split()
                            b = a[5]
                            tempList.append(int(b[:-1]))
                        else:
                            pass
                        tempNum = tempNum + 1
                    for xKPI in range(len(tempList)):
                        if tempList[xKPI] > kPI :
                            print("大于的num:%s"%(tempList[xKPI]))
                        else:
                            print("不存在大于%s的本次"%(kPI))
                        
            else:
                pass
    
    return(dayCtimes)
# 判断内存开始和结束以及最大值
def check_memorylog(startTime,finishTime,readPath,writePath,exclPath):
    startLineNum = 0
    finishLineNum = 0
    tempLienNum = 0
    tempList = []
    templine = 0
    tempNum = 0

    #取得开始结束以及最大值 
    with open(readPath,'r',encoding = 'UTF-8',errors = "ignore") as read_file:

        for line in read_file:
            
            tempLienNum = tempLienNum + 1
            if startTime in line:
                if deWeightTime in line:
                    continue
                else:
                    startLineNum = tempLienNum
                    # print(startLineNum)
            if finishTime in line:
                if "END" in line:
                    continue
                else:
                    finishLineNum = tempLienNum
                    # print(finishLineNum)

        if startLineNum <= finishLineNum and finishLineNum > startLineNum:
            with open(readPath,'r',encoding='UTF-8',errors="ignore") as read_file:

                for xline in read_file:
                    xline = xline.strip('\n')
                    if tempNum >= startLineNum and tempNum < finishLineNum:

                        a = xline.split()
                        b = a[5]
                        tempList.append(int(b[:-1]))
                    else:
                        pass
                    tempNum = tempNum + 1
        else:
            print("！输入参时间错误！")
    print("起始内存值：%s MB,结束内存值：%s MB,最大内存值：%s MB"%(tempList[0],tempList[-1],max(tempList)))

    startNum = '起始内存值: ' + str(tempList[0]) + ' MB'
    endNum = '结束内存值: ' + str(tempList[-1]) + ' MB'
    maxNum = '最大内存值: ' + str(max(tempList)) + ' MB'
    # 

    return(startNum,endNum,maxNum)

# 读取excel的类
class ExcelData():
    # 初始化方法
    def __init__(self, data_path, sheetname):
        #定义一个属性接收文件路径
        self.data_path = data_path
        # 定义一个属性接收工作表名称
        self.sheetname = sheetname
        # 使用xlrd模块打开excel表读取数据
        self.data = xlrd.open_workbook(self.data_path)
        # 根据工作表的名称获取工作表中的内容（方式①）
        self.table = self.data.sheet_by_name(self.sheetname)
        # 根据工作表的索引获取工作表的内容（方式②）
        # self.table = self.data.sheet_by_name(0)
        # 获取第一行所有内容,如果括号中1就是第二行，这点跟列表索引类似
        self.keys = self.table.row_values(0)
        # 获取工作表的有效行数
        self.rowNum = self.table.nrows
        # 获取工作表的有效列数
        self.colNum = self.table.ncols
    # 判断有效行数
    def readRowValues(self):
        rows = self.rowNum
        # self.table = self.data.sheet_by_name(0)
        # 获取第一行所有内容,如果括号中1就是第二行，这点跟列表索引类似
        # print(self.table.row(rows-1))
        return rows
    # 定义一个读取excel表的方法
    def readExcel(self):
        # 定义一个空列表
        datas = []
        for i in range(1, self.rowNum):
            # 定义一个空字典
            sheet_data = {}
            for j in range(self.colNum):
                # 获取单元格数据类型
                c_type = self.table.cell(i,j).ctype
                # 获取单元格数据
                c_cell = self.table.cell_value(i, j)
                if c_type == 2 and c_cell % 1 == 0:  # 如果是整形
                    c_cell = int(c_cell)
                elif c_type == 3:
                    # 转成datetime对象
                    date = datetime.datetime(*xldate_as_tuple(c_cell,0))
                    c_cell = date.strftime('%Y/%d/%m %H:%M:%S')
                elif c_type == 4:
                    c_cell = True if c_cell == 1 else False
                sheet_data[self.keys[j]] = c_cell
                # 循环每一个有效的单元格，将字段与值对应存储到字典中
                # 字典的key就是excel表中每列第一行的字段
                # sheet_data[self.keys[j]] = self.table.row_values(i)[j]
            # 再将字典追加到列表中
            datas.append(sheet_data)
        # 返回从excel中获取到的数据：以列表存字典的形式返回
        return datas

# 覆盖写入的类
class ExcelWrite(object):
    def __init__(self,write_Path):
        self.write_Path = write_Path  # # excel的存放路径
        self.excel = xlwt.Workbook()  # 创建一个工作簿
        self.sheet = self.excel.add_sheet('Sheet1')  # 创建一个工作表
    
    # 写入单个值
    def write_value(self, cell, value):
        '''
            - cell: 传入一个单元格坐标参数，例如：cell=(0,0),表示修改第一行第一列
        '''
        self.sheet.write(*cell, value)
        # （覆盖写入）要先用remove(),移动到指定路径，不然第二次在同一个路径保存会报错
        os.remove(self.write_Path)
        self.excel.save(self.write_Path)
        
    # 写入多个值
    def write_values(self, cells, values):
        '''
            - cells: 传入一个单元格坐标参数的list，
            - values: 传入一个修改值的list，
            例如：cells = [(0, 0), (0, 1)],values = ('a', 'b')
            表示将列表第一行第一列和第一行第二列，分别修改为 a 和 b
        '''
        # 判断坐标参数和写入值的数量是否相等
        if len(cells) == len(values):
            for i in range(len(values)):
                self.write_value(cells[i], values[i])
        else:
            print("传参错误,单元格：%i个,写入值：%i个" % (len(cells), len(values)))

# UCS-2 little endian方法A(不适用可以删除)
def parseFileA(filepath):
    linelist = []
    try:
        with open(filepath,'r') as fp:
            temp = 0
            encoding = 'utf-16-le'
            with codecs.open(filepath, 'r', encoding) as fp2:
                soup = BeautifulSoup(fp2)
                # print(soup)
                print(type(soup))
                linelist.append(soup)
                # print(linelist)
        return(linelist)
    except Exception:
        print('[ERROR]')
# UCS-2 little endian方法B(不适用可以删除)
def parseFileB(filepath):
    try:
        lineList = [] # 存放每一行的内容
        with open(filepath, 'r') as fp:
            line = fp.read()
            print(line)
            if line.startswith('\xff\xfe'):
                encoding = 'utf-16-le'
                fp2 = codecs.open(filepath, 'r', encoding)
                lineList = fp2.readlines()
                fp2.stream.close()
        for i in lineList: # 打印每一行
            print(i)
    except Exception:
        print('[ERROR]')


if __name__ == '__main__':
    # sheetname = "Sheet1"
    with open(configJsonPath) as c:
        config = json.load(c)
    #判断 Output_Path 文件夹是否存在
        if not os.path.exists(config['Output_Path']):
            print("Output_Path 已存在")
        else:
            #结果False 就创建文件夹 
            os.makedirs(config['Output_Path'])
        for d in (config.keys()):
            if d != "Output_Path" and d != "Valgrind_File" :
                data_perison = config.get(d)
                for item in data_perison.keys():
                    if item == "grade":
                        data_grade = data_perison["grade"]
                        data_name = data_perison["name"]
                        data_startTime = data_grade['startTime']
                        data_endTime = data_grade['endTime']
                        data_setupFileName = data_grade['setupFilePath']
                        data_setupFilePath = os.path.join(os.path.abspath(os.path.dirname(__file__)),data_setupFileName).replace("\\",'/')
                        print('==============================')
                        print(data_name)
                        print(data_startTime)
                        print(data_endTime)
                        print(data_setupFilePath)
                        startNum,endNum,maxNum = check_memorylog(data_startTime,data_endTime,data_setupFilePath,config['Output_Path'],config['Valgrind_File'])
                        validTime = check_memoryTime(data_startTime,data_endTime,data_setupFilePath,config['Output_Path'],config['Valgrind_File'])
                        print('==============================')
                    else:
                        pass
                        # print("JSON文件设置项配置错误")
                    
            else:
                pass
                # print("不存在Output_Path和Valgrind_File文件夹")



    # 旧的设置项管理模块(不适用可以删除)
    # memoryInfo = MemorylogManager(deWeightTime,setupFilePath)
    # startTime,finishTime,readPath,writeFileName,exclPath = setup_Working_Directory(memoryInfo)

    # excel表的方法分析类
    # get_data = ExcelData(exclPath,sheetname)
    # datarows = get_data.readRowValues()
    # excel表的写入类
    # start = ExcelWrite(writeFileName)
    # cells1 = [(0,0),(1,1),(2,2)]
    # values1 = (startNum,endNum,maxNum)
    # start.write_values(cells1,values1)



