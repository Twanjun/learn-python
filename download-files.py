# -*- coding:utf-8 -*-
import requests
import urllib.parse
import re
import xlrd
import os

workbook_path = r'F:\python\**.xls' #url excel表路径
workbook = xlrd.open_workbook(workbook_path)
number_sheets = workbook.nsheets
worksheets = workbook.sheet_names()
n = 0
i = 0

while n < number_sheets:
    folder_name = worksheets[n] #获取第n个工作表名
    sheet_content = workbook.sheet_by_index(n)  #获取第n个工作表内容
    num_rows = sheet_content.nrows  #获取第n个工作表总行数
    folder_path = 'F:\python\\'+ folder_name + '\\' #按照第n个工作表名构建文件夹路径，不同工作表里的url列表下载的文件存放在各自的工作表名文件夹下
    try:
        os.mkdir(folder_path)   #创建文件夹
    except OSError as error:
        print(error)    #文件夹已存在时，返回错误
    while i < num_rows:
        row = sheet_content.row_values(i)   #返回结果是列表
        url = str(row[0])   #读取列表第1个，转换数据类型（这里这一步非必要，已是string），正则表达式只能按照某种模式匹配字符串
        matchobject = re.match(r'https://s3plus.sankuai.com/s3forkld/(.*).zip.*', url)  #构建从url字符串的起始位置匹配压缩文件名模式
        encodestring = matchobject.group(1)  #获取压缩文件名URL编码
        filename = urllib.parse.unquote(encodestring)  #URL解码成中文字符
        file = requests.get(url)  #请求连接url
        open(folder_path + filename + '_' + str(i+1) + '.zip', 'wb').write(file.content)  #打开文件写进url内容，文件不存在时会自动生成，即压缩包文件
        i = i + 1
    print(worksheets[n])
    n = n + 1
    i = 0








