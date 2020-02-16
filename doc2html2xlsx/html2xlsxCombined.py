#!/usr/bin/env python3
# -*- coding: UTF-8 -*-
# 批量将html中表格合并转换成一个xlsx的sheet，并保存
import fnmatch
import os
import pandas as pd

# pip3 install pandas xlsxwriter
curPath = os.path.abspath(os.path.dirname(__file__))
path = curPath + "/_results"
# 获取所有html
dirs = fnmatch.filter(os.listdir(path), '*.html')
if len(dirs) == 0:
    print(path + " does not have html to covert xlsx. pls check or replace it.")
    os._exit(0)
writer = pd.ExcelWriter(path + "/result.xlsx", engine='xlsxwriter')
for file in dirs:
    url = "file://" + path + "/" + file
    sheet = file.split('.')[0]
    print("converting " + file + " to sheet " + sheet + " of result.xlsx")
    # 读取第一张表
    table = pd.read_html(url)[0]
    table.to_excel(writer, sheet_name=sheet, index=False, header=False)
    # 设置样式
    format = writer.book.add_format({'text_wrap': True})
    worksheet = writer.sheets[sheet]
    # 设置row height
    worksheet.set_default_row(60)
    # 设置width
    for idx, col in enumerate(table):
        # set column width
        worksheet.set_column(idx, idx, 30, format)
    # 缩小第一行和列长度
    worksheet.set_column(0, 0, 10)
    worksheet.set_row(0, 15)
writer.save()
writer.close()
