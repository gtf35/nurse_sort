#!/usr/bin/env python3
# -*- coding: utf-8 -*-

' ui 模块 '

__author__ = 'gtf35'

#系统
import os

#网络框架
from flask import Flask, request, flash, redirect, url_for
from werkzeug.utils import secure_filename

#工具类
import utils
#配置文件
import config

def _test():
    init()

def _initHome(app):
    @app.route('/', methods=['GET', 'POST'])
    def home():
        if request.method == 'POST':
            #确认请求中有文件
            if 'file' not in request.files:
                return "没有上传文件"
            file = request.files['file']
            # 如果用户没有选择文件, 浏览器仍然提交了一个没有文件名的空的部分
            if file.filename == '':
                return "没有选择文件"
            if not allowed_file(file.filename):
                return "请上传指定类型的文件"
            if file:
                filename = secure_filename(file.filename)
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                return "上传成功"
        return utils.readTxT(config.homeHtmlPath)

def _runFlask(app):
    app.run(
        host = config.flaskHost,
        port = config.flaskPort,
        debug = config.flaskDebug
    )
def init():
    app = Flask(__name__)
    app.config['UPLOAD_FOLDER'] = config.fileUploadPath
    _initHome(app)
    _runFlask(app)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in config.uploadExtensions


if __name__=='__main__':
   _test()

