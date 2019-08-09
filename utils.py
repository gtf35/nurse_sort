#!/usr/bin/env python3
# -*- coding: utf-8 -*-

' 工具 模块 '

__author__ = 'gtf35'

#从指定路径读取文本
def readTxT(path):
    result = ""
    f = open(path)
    line = f.readline()
    while line:
        result = result + line;
        line = f.readline()
    f.close()
    return result;