#!/usr/bin/python
# -*- coding: UTF-8 -*-

from __future__ import print_function
import os
import xlwt
import xlrd
from xlrd import xldate_as_tuple
import datetime
import json
import pandas as pd
import time
import xlsxwriter

class KeyType:
    # 搜索内存关键字
    deWeightTime = "top -m | grep MXNavi"
    # 原有setup配置项(当前工作路径不采用,可考虑删除)
    setupFileName = "SETUP.txt"
    # 最新批量json配置项
    configJsonName = "config/config.json"
    # 原有setup配置项地址获取(当前工作路径不采用,可考虑删除)
    setupFilePath = os.path.join(os.getcwd(),setupFileName).replace("\\",'/')
    # 最新批量json配置项地址获取(获取当前文件的绝对路径)
    configJsonPath = os.path.join(os.path.abspath(os.path.dirname(__file__)),configJsonName).replace("\\",'/')
class Animalm(object):
    startLineNum = 0    # 开始行数
    finishLineNum = 0   # 结束行数

    kPI = 1024                    # 内存KPI 阀值
    secondaryMaximum = 1000      # 次峰值范围值
    divide_The_Value = 300      # 开始与结束的落差 阀值
    divide_The_Time = 60      # 保持时间≥60s 阈值

    def __init__(self,name,startTime,finishTime,readPath,writePath,exclPath):
        self.name = name
        self.startTime = startTime
        self.finishTime = finishTime
        self.readPath = readPath
        self.writePath = writePath
        self.exclPath = exclPath

    def check_tmpLine(self):
        # tempList = []       # 标注列表
        tempLienNum = 0     # 标注行数
        # tempNum = 0         # 对比行数
    
        with open(self.readPath,'r',encoding = "UTF-8",errors = "ignore") as read_file:
            for line in read_file:
                tempLienNum = tempLienNum + 1
                if self.startTime in line:
                    if KeyType.deWeightTime in line:
                        continue
                    else:
                        startLineNum = tempLienNum
                if self.finishTime in line:
                    if "END" in line:
                        continue
                    else:
                        finishLineNum = tempLienNum
        return(startLineNum,finishLineNum)

class CheckAmemory(Animalm):

    def __init__(self,name,startTime,finishTime,readPath,writePath,exclPath,type='虚构类'):
        self.type = type
        Animalm.__init__(self,name,startTime,finishTime,readPath,writePath,exclPath)

    #判断有效运行时间
    def check_memoryTime(self):
        tempList = []       # 标注列表
        tempLienNum = 0     # 标注行数
        tempNum = 0         # 对比行数
    
        # 由字符串格式转化为日期格式的函数为: datetime.datetime.strptime()
        vSt = datetime.datetime.strptime(self.startTime.replace("/",'-'),"%Y-%m-%d %H:%M:%S")
        vFt = datetime.datetime.strptime(self.finishTime.replace("/",'-'),"%Y-%m-%d %H:%M:%S")
        # print("开始时间：%s 和对应格式 %s"%(vSt,type(vSt)))
        # print("结束时间：%s 和对应格式 %s"%(vFt,type(vFt)))
        # 2.1）抽取开始和结束的时间戳，判断有效运行时间（单位：小时）
        # dayCtimes = ((vSt - vFt).total_seconds())/3600
        dayCtimes = ((vFt - vSt).total_seconds())/3600
        print("本次有效时间-----共%.2f小时-----"%(dayCtimes))
        # 2.2）以及当内存超1024MB时所需时间。（超 kpi xxxxMB才需判断）
        stmpLine = Animalm.check_tmpLine(self)[0]      # 调用父类check_tmpLine方法并赋值起始行，防止初始化循环调用
        ftmpLine = Animalm.check_tmpLine(self)[1]      # 调用父类check_tmpLine方法并赋值结束行，防止初始化循环调用
        # print(Animalm.check_tmpLine(self)[0],Animalm.check_tmpLine(self)[1])
        if stmpLine <= ftmpLine and ftmpLine > stmpLine:
            # print(time.strftime("%Y-%m-%d %H:%M:%S",time.localtime()))
            with open(self.readPath,'r',encoding='UTF-8',errors="ignore") as read_file:
                for xline in read_file:
                    xline = xline.strip('\n')
                    # print(xline)
                    if tempNum >= stmpLine and tempNum < ftmpLine:
                        a = xline.split()
                        b = a[5]
                        # tempList.append(int(b[:-1]))
                        if int(b[:-1]) == CheckAmemory.kPI :
                            surpassDay = str((a[0]).strip('['))
                            surpassHour = (a[1])
                            surpassStr = (surpassDay + ' '+ surpassHour).rsplit(']')
                            surpassTime = datetime.datetime.strptime((surpassStr[0]).replace("/",'-'),"%Y-%m-%d %H:%M:%S")
                            print("开始超过 %sMB 的时间戳为 %s"%(CheckAmemory.kPI,a[0] + ' ' +a[1]))
                            surpassTimeS = ((vSt - surpassTime).seconds)/3600
                            print("导航放置 %.2f 小时到达 %s MB"%(surpassTimeS,CheckAmemory.kPI))
                            break
                        else:
                            pass
                    else:
                        pass
                    tempNum = tempNum + 1
            # print(time.strftime("%Y-%m-%d %H:%M:%S",time.localtime()))

        else:
            pass
        return(dayCtimes)

    # 判断内存开始和结束以及最大值
    def check_memorylog(self):
        tempLienNum = 0        # 标注行数
        tempList = []          # 标注列表
        # templine = 0
        tempNum = 0            # 对比行数
        KPIList = []         # 获取超KPI的数据列表

        stempLienNum = 0
        stempList = []
        stempNum = 0
        stempNumOneT = 0
        stempNumi = 0
        
        dtvList = []

        #取得开始结束值以及最大值 
        stmpLine = Animalm.check_tmpLine(self)[0]      # 调用父类check_tmpLine方法并赋值起始行，防止初始化循环调用
        ftmpLine = Animalm.check_tmpLine(self)[1]      # 调用父类check_tmpLine方法并赋值结束行，防止初始化循环调用
        # print(Animalm.check_tmpLine(self)[0],Animalm.check_tmpLine(self)[1])
        if stmpLine <= ftmpLine and ftmpLine > stmpLine:
            with open(self.readPath,'r',encoding='UTF-8',errors="ignore") as read_file:
                for xline in read_file:
                    xline = xline.strip('\n')
                    if tempNum >= stmpLine and tempNum < ftmpLine:
                        a = xline.split()
                        b = a[5]
                        tempList.append(int(b[:-1]))
                    else:
                        pass
                    tempNum = tempNum + 1
        else:
            print("！输入参时间错误！")
        print("起始内存值：%s MB,结束内存值：%s MB,最大内存值：%s MB"%(tempList[0],tempList[-1],max(tempList)))
        sNum = tempList[0]
        eNum = tempList[-1]
        mNum  = max(tempList)
        startNum = '起始内存值: ' + str(tempList[0]) + ' MB'
        endNum = '结束内存值: ' + str(tempList[-1]) + ' MB'
        maxNum = '最大内存值: ' + str(max(tempList)) + ' MB'

        # 实现场景1：判断峰值或结束值是已超1024MB
        if mNum >= CheckAmemory.kPI and eNum >= CheckAmemory.kPI :
            print("实现场景1：判断峰值和结束已超1024MB")
            self.write_to_excel()
            # 取出全部数据生成Excel sheet    
            if stmpLine <= ftmpLine and ftmpLine > stmpLine:
                with open(self.readPath,'r',encoding='UTF-8',errors="ignore") as read_file:
                    for kline in read_file:
                        if "BEGIN" in kline:
                            continue
                        elif "END" in kline:
                            continue
                        else:
                            kline = kline.strip('\n').split()
                            KPIList.append(kline)
            else:
                print("起始结束行位置错误")

        # 实现场景2：未超1024MB，但长时间保持在1000MB，也就是在1000-1024之间长时间保持（保持时间 暂定≥60s，后期改为可配置）
        elif CheckAmemory.secondaryMaximum <= mNum < CheckAmemory.kPI or CheckAmemory.secondaryMaximum <= eNum < CheckAmemory.kPI :
            if stmpLine <= ftmpLine and ftmpLine > stmpLine:
                with open(self.readPath,'r',encoding='UTF-8',errors="ignore") as read_file:
                    i = 0
                    b = 0
                    for sline in read_file:
                        sline = sline.strip('\n')
                        if stempNum >= stmpLine and stempNum < ftmpLine:
                            sa = sline.split()
                            sb = sa[5]
                            stempNumOneT = int(sb[:-1])
                            # 判断范围在1000 - 1024之间的时间和信息   方法：不连续，用“1”来间隔
                            if stempNumOneT >= CheckAmemory.secondaryMaximum and stempNumOneT < CheckAmemory.kPI:
                                
                                # 加入所在目标行位置
                                sa.append(i+1)
                                # print(sa)
                                stempList.append(sa)
                                # print(stempList)
                                # print("1000<=X<1024 所在行位置： %s  的行数 = %d"%(sa,i+1))
                                b = b+1
                            else:
                                pass
                                # print("不在1000<=X<1024范围内")
                        else:
                            pass
                            # print("！输入参数错误！")
                        stempNum = stempNum + 1
                        i = i +1
                        if i != ftmpLine :
                            continue
                        elif i == ftmpLine + 1 :
                            break
                        print('超过在1000 - 1024之间的连续时间：%d s'%b)
                        if b >=  CheckAmemory.divide_The_Time :
                            print("实现场景2：未超1024MB，但长时间保持在1000MB，也就是在1000-1024之间长时间保持（保持时间 暂定≥%ds）"%CheckAmemory.divide_The_Time)
                        else:
                            print("未实现场景2：未超1024MB，但长时间保持在1000MB，也就是在1000-1024之间长时间保持未≥%ds）"%CheckAmemory.divide_The_Time)

                # break
            # else:
            #     print("！输入参时间错误！")

        # 实现场景3：内存一直未超1000MB,但开始与结束的落差值在xxx（divide_The_Value =300mb,后期改为可配置在json文件中）
        elif mNum < CheckAmemory.secondaryMaximum and (int(eNum) - int(sNum)) >= CheckAmemory.divide_The_Value :
            print("实现场景3：内存一直未超1000MB,但开始与结束的实际落差值在%sMB,已超过KPI: %s MB"%((int(eNum) - int(sNum)),CheckAmemory.divide_The_Value))
            # self.write_to_excel()
            # 取出存在落差的全部数据创新sheet,并创建柱状图
            if stmpLine <= ftmpLine and ftmpLine > stmpLine:
                with open(self.readPath,'r',encoding='UTF-8',errors="ignore") as read_file:
                    for xline in read_file:
                        xline = xline.strip('\n')
                        dtvList.append(xline)
                # print(dtvList)
        # 实现场景4：判断峰值或结束值是未超1024MB，且未长时间1000MB和且未超开始结束落差值
        else:
            print("实现场景4：本次内存峰值和结束值不存在超%sMB的测试场景"%(CheckAmemory.kPI))

        return(startNum,endNum,maxNum,KPIList)
    # @CallingCounter
    def write_to_excel(self):
        print('写入函数被调用：1次')
        alist = []      
        blist = []
        flist = []
        # lenlist = []  
        count = 0
        headings = ['年月日','小时','内存值(MB)']
        sheet = self.name
        print('--write_to_excel--')
        # rb = xlrd.open_workbook(self.writePath)
        workbook = xlsxwriter.Workbook(self.writePath)
        workbooksheet = workbook.add_worksheet(sheet)
        workbooksheet.write_row('A1',headings)
        # workbooksheet.write_row()
        with open(self.readPath,'r',encoding='UTF-8',errors="ignore") as read_file:
            for kline in read_file:
                # lenlist.append(len(kline))
                if "BEGIN" in kline:
                    continue
                elif len(kline) >= 84 :
                    kline = kline.strip('\n').split()
                    print(type((kline[0]).strip()))
                    alist.append((kline[0]).strip())
                    # print(kline[0])
                    print(type((kline[1]).strip()))
                    blist.append((kline[1]).strip())
                    # print(kline[1])
                    if kline[5] != None:
                        # print(kline[5])
                        a = (kline[5].strip())
                        flist.append(int(a[:-1]))
                    else:
                        break
                elif "END" in kline:
                    continue
                else:
                    pass
                    # print("数据不满足11位")
        workbooksheet.write_column('A2',alist)
        workbooksheet.write_column('B2',blist) 
        workbooksheet.write_column('C2',flist) 
        # workbooksheet.write_column('D2',lenlist)
        print('sheet 名称：%s'%(sheet))
        categoriesLen = len(blist)
        valuesLen = len(flist)
        print(categoriesLen)
        print(valuesLen)
        chart_col = workbook.add_chart({'type':'line'})
        chart_col.add_series(
            {
            'name':'= sheet!$B$1',
            'categories':'= sheet!$B$2:$B$10',
            'values':'= sheet!$C$2:$C$10',
            'line':{'color':'red'},
            }
        )
        chart_col.set_title({'name':'测试'})
        chart_col.set_x_axis({'name':'运行时间'})
        chart_col.set_y_axis({'name':'内存值'})
        chart_col.set_style(1)
        # 放置位置
        workbooksheet.insert_chart('E2',chart_col,{'x_offset':25,'y_offset':10})

        count += 1
        workbook.close() 

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
    # 将字典列表导出到excel文件中：待验证

# 统计函数被调用的次数类
class CallingCounter(object):
    def __init__(self,func):
        self.func = func
        self.count = 0

    def __call__(self,*args,**kwargs):
        self.count += 1
        return self.func(*args,**kwargs)




if __name__ == '__main__':
    # 读取json 配置文件路径
    with open(KeyType.configJsonPath) as c:
        config = json.load(c)
        writeFileName = config['Output_Path']
        print(writeFileName)
    #判断 Output_Path 文件夹是否存在
        if os.path.exists(config['Output_Path']):
            print("Output_Path 已存在")
        else:
            #结果False 就创建文件夹 
            os.makedirs(config['Output_Path'])
        for d in (config.keys()):
            if d != "Output_Path" and d != "Valgrind_File" :
                data_perison = config.get(d)
                print(d)
                for item in data_perison.keys():
                    if item == "grade":
                        data_grade = data_perison["grade"]
                        data_name = data_perison["name"]
                        data_startTime = data_grade['startTime']
                        data_endTime = data_grade['endTime']
                        data_setupFileName = data_grade['setupFilePath']
                        data_setupFilePath = os.path.join(os.path.abspath(os.path.dirname(__file__)),data_setupFileName).replace("\\",'/')
                        print('==============================start==============================')
                        print(data_name)
                        print(data_startTime)
                        print(data_endTime)
                        print(data_setupFilePath)
                        ap = Animalm(data_name,data_startTime,data_endTime,data_setupFilePath,config['Output_Path'],config['Valgrind_File'])
                        # ap.check_tmpLine()
                        # persion = CheckAmemory(data_startTime,data_endTime,data_setupFilePath,config['Output_Path'],config['Valgrind_File'])
                        persion = CheckAmemory(data_name,data_startTime,data_endTime,data_setupFilePath,config['Output_Path'],config['Valgrind_File'])
                        persion.check_tmpLine()
                        persion.check_memoryTime()
                        persion.check_memorylog()
                        # a = Check_Amemory(data_startTime,data_endTime,data_setupFilePath,config['Output_Path'],config['Valgrind_File'])
                        # validTime = check_memoryTime(data_startTime,data_endTime,data_setupFilePath,config['Output_Path'],config['Valgrind_File'])
                        # Check_Amemory.
                        print('==============================end==============================')
                    else:
                        pass
                        # print("JSON文件设置项配置错误")  
            else:
                pass
                # print("不存在Output_Path和Valgrind_File文件夹")

    # print(f'函数被调用了：{write_to_excel.count}次')
    # excel表的方法分析类
    # get_data = ExcelData(exclPath,sheetname)
    # datarows = get_data.readRowValues()
    # excel表的写入类
    # start = ExcelWrite(writeFileName)
    # cells1 = [(0,0),(1,1),(2,2)]
    # values1 = (startNum,endNum,maxNum)
    # start.write_values(cells1,values1)





