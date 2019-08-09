#!/usr/bin/env python3
# -*- coding: utf-8 -*-

' ui 模块 '

__author__ = 'gtf35'


import tkinter as tk
import tkinter.messagebox as messagebox
import os
import config


# 设置窗口
window = tk.Tk()  # 建立一个窗口
window.title(config.UITitle)
window.geometry('300x200')  # 窗口大小为300x200


var = tk.StringVar()  # 文字变量储存器


# 设置标签
l = tk.Label(textvar=var, width=20, height=2)  # 参数textvar不同于text,bg是backgroud
l.pack()  # 放置标签
var.set(config.UIText)

def createDemo():
    try:
        import excelDemo
        excelDemo.createDemo()
        messagebox.showinfo(config.UISuccessTip, config.createDemoSuccessTip)
    except BaseException as e:
        messagebox.showinfo(config.UIFailTip, config.createDemoFailTip + str(e))

def openDataPath():
    try:
        os.system("explorer.exe %s" % os.path.abspath(config.dataPath))
    except BaseException as e:
        messagebox.showinfo(config.UIFailTip, config.openDataPathFailTip + str(e))

def sortExcel():
    try:
        import excelDemo
        if not excelDemo.checkIsDemo(config.demoExcelPath):
            messagebox.showinfo(config.UIFailTip, config.checkExcelFailTip)
            return
        import excelSort
        se = excelSort.SortExcel(config.demoExcelPath)
        sortList = se.sort()
        se.writeSortNurses(sortList, config.sortedExcelPath)
        messagebox.showinfo(config.UISuccessTip, config.sortExcelSuccessTip)
    except BaseException as e:
        messagebox.showinfo(config.UIFailTip, config.sortExcelFailTip + str(e))




createDemoBtn = tk.Button(text='生成模板', width=20, height=2, command=createDemo)
createDemoBtn.pack()

sortExcelBtn = tk.Button(text='将模板排班', width=20, height=2, command=sortExcel)
sortExcelBtn.pack()

openDataPathBtn = tk.Button(text='打开工作目录', width=20, height=2, command=openDataPath)
openDataPathBtn.pack()

window.mainloop()  # 循环，时刻刷新窗口