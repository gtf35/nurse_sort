#!/usr/bin/env python3
# -*- coding: utf-8 -*-

' 常量 模块 '

__author__ = 'gtf35'

#home路径
homeHtmlPath = "./static/index.html"
#文件上传路径
fileUploadPath = "./temp/"
#上传文件拓展名限制
uploadExtensions = set(['xls', 'xlsx'])
#网络框架监听主机
flaskHost = '127.0.0.1'
#网络框架监听端口
flaskPort = 8888
#网络框架是否开启调试
flaskDebug = False
#工作目录
dataPath = "./data/"
#数据表名字默认颜色
workSheetDefaultColor = "1072BA"
#人员名单数据表的名字
nurseWorkSheetName = "人员名单"
#设置数据表的名字
settingsWorkSheetName = "设置"
#排序结果数据表的名字
outputWorkSheetName = "排序结果"
#人员清单提示单元格位置
nurseWorkSheetTipCellAddr = "A1"
#人员清单提示单元格内容
nurseWorkSheetTipCellText = "请勿删除此行！！！在这里的第一列(A)放入待排序的名单"
#demo 数据表文件保存目录
demoExcelPath = "./data/NurseDemo.xlsx"
#排班后的数据表文件保存目录
sortedExcelPath = "./data/NurseSorted.xlsx"
#设置提示单元格内容
settingsSheetTipCellText = "请勿删除此行！！！在对应条目下方填写对应的参数，注意参数不要跨年！！！否则数据不正常！！ 年为4位(2019)，月为1到2位(2或10这种)，随机开关只能写（开/关)"
#设置提示单元格位置
settingsSheetTipCellAddr = "A1"
#设置随机单元格位置
settingsSheetRandomCellAddr = "C3"
#设置随机单元格文字
settingsSheetRandomCellText = "随机开关"
#设置随机输入单元格位置
settingsSheetInputRandomCellAddr = "C4"
#设置随机输入单元格默认
settingsSheetInputRandomCellDefaultText = "关"
#开始年份单元格内容
settingsSheetBeginYearCellText = "开始年份"
#开始年份单元格位置
settingsSheetBeginYearCellAddr = "A3"
#开始年份单元格输入位置
settingsSheetInputBeginYearCellAddr = "A4"
#开始月份单元格内容
settingsSheetBeginMonthCellText = "开始月份"
#开始月份单元格位置
settingsSheetBeginMonthCellAddr = "B3"
#开始月份单元格输入位置
settingsSheetInputBeginMonthCellAddr = "B4"
#结束年份单元格内容
settingsSheetEndYearCellText = "结束年份"
#结束年份单元格位置
settingsSheetEndYearCellAddr = "A5"
#结束年份输入单元格位置
settingsSheetInputEndYearCellAddr = "A6"
#结束月份单元格内容
settingsSheetEndMonthCellText = "结束月份"
#结束月份单元格位置
settingsSheetEndMonthCellAddr = "B5"
#结束输入月份单元格位置
settingsSheetInputEndMonthCellAddr = "B6"
#
createDemoSuccessTip = "模板生成成功！\n在工作目录下的NurseDemo.xlsx\n请在根据里面的说明填入内容，填入之后不要改文件名，放入工作目录内"
sortExcelSuccessTip = "排班成功！\n在工作目录下的NurseSorted.xlsx"
UISuccessTip = "恭喜"
UIFailTip = "抱歉"
UITitle = "排班系统"
UIText = "排班系统 V1.0"
createDemoFailTip = "模板生成出错！！\n错误信息：\n"
openDataPathFailTip = "打开工作目录出错！！\n错误信息：\n"
checkExcelFailTip = "模板文件匹配失败，请检查模板文件是否存在，或者未按照内部要求修改，您可以尝试重新生成模板文件"
sortExcelFailTip = "排班出错！！\n错误信息：\n"