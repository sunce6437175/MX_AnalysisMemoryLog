※当前文件夹一定要放在   非中文   路径下
文件夹和脚本文件支持改名 （※非中文）

准备环节
1.	导航连接Xshell后，在Xshell输入top -m | grep MXNavi, 日志启动保存，放置完毕后control + C 关闭。
2.	Xshell日志中将生成源文件（如：CNS3.0_2020-03-25_18-28-34.log）放在与脚本工具[AnalysisMemoryLog.exe] 同一路径下，然后配置入参文件
3.	使用Notepad++ 分别打开SETUP.txt 和CNS3.0_2020-03-25_18-28-34.log，在CNS3.0_2020-03-25_18-28-34.log中获取如下参数拷贝到SETUP.txt对应位置中
※每一行四个参数，参数之间用  英文逗号  隔开※
如：2020/3/25 18:20:22,2020/3/26 9:08:07,CNS3.0_2020-03-25_18-28-34.log,CNS3MemoryNew.xls,CNS3.0_ST_Checklist_WDXTest_SOP1&SOP1.5.xlsx
按照入参文件步骤，
将第一个参数 ：开始时间字段，如： 2020/3/25 18:20:22 (注意：[2020/3/21 16:46:18] [root@vw-infotainment-036761:~]# top -m | grep MXNavi不是这个前面的时间点，是下一个的时间点)。
第二个参数 ：结束时间字段，如：2020/3/26 9:08:07取出后保存在SETUP.txt中第一个参数和第二个参数位置上，SETUP.txt进行文件保存后关闭。
第三个参数: 源文件名称，如：CNS3.0_2020-03-25_18-28-34.log（源文件）。
4.	之前Notepad++ 打开CNS3.0_2020-03-25_18-28-34.log 先不关闭，点击 编码（N）,选择 使用UTF-8编码 后，在保存关闭
5.	双击运行 AnalysisMemoryLog.exe，会在CNS3MemoryNew.xls 生成开始/结束以及峰值的内存值大小（CNS3MemoryNew.xls 每次覆盖前一次的结果）

入参文件：SETUP.txt
※每一行四个参数，参数之间用  英文逗号  隔开
如：2020/3/25 18:20:22,2020/3/26 9:08:07,CNS3.0_2020-03-25_18-28-34.log,CNS3MemoryNew.xls,CNS3.0_ST_Checklist_WDXTest_SOP1&SOP1.5.xlsx

第一个参数：内存记录Xshell Log截取开始时间字段（2020/3/25 18:28:39）
第二个参数：内存记录Xshell Log截取开始时间字段（2020/3/26 9:08:04）
第三个参数：启动时间源文件名称(要加文件后缀名，如：CNS3.0_2020-03-25_18-28-34.log)
第四个参数：解析内存值后生成Excel表格名称，名称不可重复，既有文件无需删除，每次生成最新一次的解析结果(要加文件后缀名，如：CNS3MemoryNew.xls)


环境：
Window OS
PYTHON 2.X&3.X

脚本文件：AnalysisMemoryLog.exe
运行方法：双击即可运行
内存解析表格生成文件路径：当前文件夹下“CNS3MemoryNew.xls”（每次覆盖前一次的结果）

