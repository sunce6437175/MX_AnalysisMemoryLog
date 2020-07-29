※稳定性log分析及填写报告自动化工具※

背景：测试人员每天在放置稳定性后收集整理稳定性结果时耗费一定时间和精力，稳定性工具抽取log并自动生成报告的省去大量人工收集、效验、分析等工作，降低了人工分析遗漏，问题发现和结果通知更加及时和系统。

准备环节
1.导航连接Xshell后，在Xshell输入top -m | grep MXNavi, 日志启动保存（使用UTF-8编码保存日志），放置完毕后control + C 关闭。
2.Xshell日志中将生成源文件（如：CNS3.0_2020-03-25_18-28-34.log）放在服务器路径下：\\192.168.2.7\BugInfo_2020(7月29日启用)\CNS3.0 Sop1.5\非功能测试结果\稳定性测试\MX_AnalysisMemoryLog\Log_File
3.将生成源文件按照自定规则修改名称（如：CNS3.0_2020-03-25_18-28-34.log → sunce.log），修改名称时注意重复命名。（建议可按照项目名称+测试人员简拼+目标机序号 等信息进行修改，同一项目可制定统一规则）

配置文件【config.json】
※参数之间用  英文逗号  隔开 需要注意大小写※
配置文件由“测试人员基本放置信息”、“log和结果文档输出路径”、“可调节关键字参数”、“抄送人员邮件设置” 这四部分构成。
一、第一部分“测试人员基本放置信息”
	如下json结构体是每个测试人员的放置稳定性的基本信息：
	"personne1":{
		"name":"孙策",                      → 稳定性放置人员
		"project":"SOP1.5",                 → 稳定性放置人员所属项目（目前不保存在数据统计当中）
		"mailPass":"sunc@meixing.com",      → 稳定性放置人员邮件
		"grade":{
			"startTime":" ",                → 无需填写开始放置时间已实现自动获取，此处无需填写
			"endTime":" ",                  → 无需填写结束放置时间已实现自动获取，此处无需填写
			"setupFilePath":"//192.168.2.7/BugInfo_2020(7月29日启用)/CNS3.0 Sop1.5/非功能测试结果/稳定性测试/MX_AnalysisMemoryLog/Log_File/sunce.log",    → xshelllog放置路径和log名称（一次设定完毕，当无服务器和log名称规则确认后无需再次修改，可自动读取该路径下的log）
			"placingState":"normal"         → 本次稳定性放置状态（目前支持 normal→OK/NA/第三方NG 三种状态，可结合每次放置状态进行回填）
				}
			},
	以上属于单人的信息结构体，每个项目可根据人员的不同增减人员基本放置信息，例如有添加就按照相同格式添加"personne2":{}、"personne3":{}等等。（注意：在添加和删除结构体时，每个结构体之间需要用英文逗号间隔）

二、第二部分“log和结果文档输出路径”
	如下是log放置位置和稳定性文件和存在问题的曲线图表格读取位置和文件名称配置（一次设定完毕，当无服务器和log名称规则以及输出结果文件确认后无需再次修改，可自动读取输出文档，每次覆盖前一次的结果）
	"Output_Path":"//192.168.2.7/BugInfo_2020(7月29日启用)/CNS3.0 Sop1.5/非功能测试结果/稳定性测试/MX_AnalysisMemoryLog/output/CNS3MemoryNew.xls",
	"Valgrind_File":"Log_File",
	"WDX_Output_Path":"//192.168.2.7/BugInfo_2020(7月29日启用)/CNS3.0 Sop1.5/非功能测试结果/稳定性测试/MX_AnalysisMemoryLog/output/CNS3.0_ST_Checklist_WDXTest_SOP1&SOP1.5.xlsx",
	
	注意：输出稳定结果excel不可删除，因每次更新后会覆盖前一次的结果，如果删除掉之后之前的结果不在保存。

三、第三部分“可调节关键字参数”
	如下是可根据不同项目配置关键KPI，以下是SOP1.5项目总结的KPI参数
	"MEMkPI":"1024",                   → 内存峰值
	"SecondaryMaximum":"1000",         → 内存次峰值区间范围
	"DivideTheValue":"300",            → 起始结果落差值
	"DivideTheTime":"60",              → 导航保持在内存峰值与次峰值之期的时间范围

四、第四部分“抄送人员邮件设置”
	如下需要抄送人员邮件信息，如果需要添加更多成员，就在下放添加"lead5":"xx@xx.com","lead6":"xx@xx.com" 程序会自动添加新抄送人员名单
	"mailpassCc":{
		"lead1":"liujun@meixing.com",
		"lead2":"gaojy@meixing.com",
		"lead3":"zhaotj@meixing.com",
		"lead4":"wuj@meixing.com"
	}


运行方法：定时自动运行脚本，并自动发送结果邮件
	（1.当测试结果存在问题时，收件人为对应稳定性放置人员，其他人员在抄送人员名单内。
	  2.当测试结果全部OK时，收件为管理人员，其他人员在抄送人员名单内。）
	  3.邮件发送测试结果包括CNS3MemoryNew.xls （存在问题人员和对应log名称、导出log信息以及内存分析曲线图
	  4.CNS3.0_ST_Checklist_WDXTest_SOP1&SOP1.5.xlsx 按照原有稳定性结果文档式样生成对应信息，比如：“开始结束时间、有效运行时间、开始结束内存值和峰值、log放置路径、测试结果、错误原因、测试人姓名等等”
	  以上结果也会生成到服务器上【第二部分“log和结果文档输出路径”】每次覆盖前一次的结果）



环境：
Window OS
PYTHON 2.X&3.X

脚本文件：AnalysisMemoryLog.py
运行方法：定时自动运行


