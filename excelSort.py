#!/usr/bin/env python3
# -*- coding: utf-8 -*-

' excle 排序处理模块 '

__author__ = 'gtf35'
#excle 处理框架
from openpyxl import *
import openpyxl
#随机
import random
#日期处理
import calendar
#配置文件
import config

class SortExcel():

    def __init__(self, path):
        self.mWorkBook = openpyxl.load_workbook(path)
        self.mNursesWorkSheet = self.mWorkBook.worksheets[0]
        self.mSettingsWorkSheet = self.mWorkBook.worksheets[1]
        self.mOutputWorkSheet = self.mWorkBook.worksheets[2]
        self.mBeginYear = self.readSettingsCell(config.settingsSheetInputBeginYearCellAddr)
        self.mBeginMonth = self.readSettingsCell(config.settingsSheetInputBeginMonthCellAddr)
        self.mEndYear = self.readSettingsCell(config.settingsSheetInputEndYearCellAddr)
        self.mEndMonth = self.readSettingsCell(config.settingsSheetInputEndMonthCellAddr)
        self.mSupportRamdom = self.getSupportRamdom()
        self.mNursesNum =  self.readNursesNum()
        self.readNurses()

    def sort(self):
        #大列表5一循环遍历位置标志
        days5LargeListIndex = 0
        #大列表2天一循环的遍历位置标志
        days2LargeListIndex = 0
        #存放排序列表
        sortList = []
        #循环月份计算排序列表
        for month in range(self.mBeginMonth, self.mEndMonth + 1, 1):
            #计算当前月份详情
            month = Month(self.mBeginYear, month)
            #存放月份整体
            monthItem = []
            #存放月份的第一行星期
            monthTitle = []
            #循环计算第一行日期
            for weekday in range(month.firstDayWeek, month.firstDayWeek + 7, 1):
                if weekday > 7 :
                    weekday = weekday -7
                if weekday == 1 :
                    monthTitle.append("一")
                if weekday == 2 :
                    monthTitle.append("二")
                if weekday == 3 :
                    monthTitle.append("三")
                if weekday == 4 :
                    monthTitle.append("四")
                if weekday == 5 :
                    monthTitle.append("五")
                if weekday == 6 :
                    monthTitle.append("六")
                if weekday == 7 :
                    monthTitle.append("日")
            #把计算出的日期写入到月份整体中
            monthItem.append(monthTitle)
            #计算一个月的名单
            for monthLineIndex in range(0, 5, 1):
                #存放一行名单（7天）
                monthLine = []
                #写出一行的人员排序
                #写出一行中的前5天
                if monthLineIndex != 4:
                    for dayOfWeekday in range(days5LargeListIndex, days5LargeListIndex + 5, 1):
                        monthLine.append(self.mNurseLargeList[dayOfWeekday])
                        #移动大列表的标志位
                        days5LargeListIndex = days5LargeListIndex + 1
                    #写出一行中的后2天
                    for dayOfWeekday in range(days2LargeListIndex, days2LargeListIndex + 2, 1):
                        monthLine.append(self.mNurseLargeList[dayOfWeekday])
                        pass
                    days2LargeListIndex = days2LargeListIndex + 2

                #是一个月的最后一行
                else:
                    for sengyu in range(35 - month.days - 1):
                        monthLine.append(self.mNurseLargeList[days5LargeListIndex])
                        days5LargeListIndex = days5LargeListIndex + 1
                        print(str(month.month) + "月最后一行第" + str(sengyu) + "位插入了：" + str(self.mNurseLargeList[sengyu]))
                    for kongwei in range(len(monthLine) + 1 , 8, 1):
                        monthLine.append(" ")
                        print(str(month.month) + "月最后一行第" + str(kongwei) + "位插入了：" + str("none"))

                #写出一行到月整体中
                monthItem.append(monthLine)
            #写出月整体到排序列表中
            sortList.append(monthItem)
        #返回排序列表
        return sortList

    def writeSortNurses(self, sortList, path):
        nowMonth = self.mBeginMonth
        nowYear = self.mBeginYear
        for monthItem in sortList:
            self.mOutputWorkSheet.append([])
            self.mOutputWorkSheet.append(["", "", "", str(nowYear) + "/" + str(nowMonth)])
            nowMonth = nowMonth + 1
            for row in monthItem:
                self.mOutputWorkSheet.append(row)
        self.mWorkBook.save(path)

    def readNurses(self):
        self.mNursesList = []
        for line in range(2, self.mNursesNum + 2 , 1):
            self.mNursesList.append(self.readNurseByLine(line))
        if self.mSupportRamdom :
            random.shuffle(self.mNursesList)
        mNurseLargeList = []
        for day in range(1, 360, self.mNursesNum):
            for nurse in self.mNursesList:
                mNurseLargeList.append(nurse)
        self.mNurseLargeList = mNurseLargeList

    def getSupportRamdom(self):
        input = self.readSettingsCell(config.settingsSheetInputRandomCellAddr)
        if input == config.settingsSheetInputRandomCellDefaultText:
            return False
        else:
            return True

    def readNursesNum(self):
        excelLineMax = self.mNursesWorkSheet.max_row
        line = excelLineMax
        while line > 0:
            if self.readNurseByLine(line) != "":
                return line -1
            line = line - 1
        return 0

    def readNurseByLine(self, index):
        return self.mNursesWorkSheet["A" + str(index)].value

    def readSettingsCell(self, addr):
        return self.mSettingsWorkSheet[addr].value

    def readOutputCell(self, addr):
        return self.mOutputWorkSheet[addr].value

class Month():
    def __init__(self, year,  month):
        self.month = month
        self.days = calendar.monthrange(year,month)[1]
        self.firstDayWeek = calendar.monthrange(year,month)[0] + 1
        print("月份详情计算：在" + str(month) + "月有" + str(self.days) + "天，1号是星期" + str(self.firstDayWeek))

if __name__ == '__main__':
    se = SortExcel(config.demoExcelPath)
    print("开始年：" + str(se.mBeginYear))
    print("开始月：" + str(se.mBeginMonth))
    print("结束年："  + str(se.mEndYear))
    print("结束月：" + str(se.mEndMonth))
    print("护士的个数：" + str(se.mNursesNum))
    print("护士清单：" + str(se.mNursesList))
    sortList = se.sort()
    print("排序结果：" + str(sortList))
    #se.writeSortNurses(sortList, config.demoExcelPath)
