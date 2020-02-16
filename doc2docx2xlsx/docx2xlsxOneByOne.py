#!/usr/bin/env python3
# -*- coding: UTF-8 -*-
# 将word中表格一个一个转换成对应的xlsx的sheet，并保存
from docx import Document
import pandas as pd
import fnmatch
import os

curPath = os.path.abspath(os.path.dirname(__file__))
# 设置待转换目录
path = curPath + "/_results"
dirs = fnmatch.filter(os.listdir(path), '*.docx')
if len(dirs) == 0:
    print(path + " does not have docx to covert xlsx. pls check or replace it.")
    os._exit(0)
for file in dirs:
    dPath = path + "/" + file
    document = Document(dPath)
    xName = file.split('.')[0]+".xlsx"
    print("converting " + file + " to " + xName)
    # 将word中表格一个一个转换成xlsx的sheet，并保存
    writer = pd.ExcelWriter(path + "/" + xName, engine='xlsxwriter')
    tables = []
    for index, table in enumerate(document.tables):
        df = []  # 用 list 套 list 的方法装二维表格内容
    for r in range(len(table.rows)):
        row = []
        for c in range(len(table.columns)):
            cell = table.cell(r, c)
            # 无内容用空格占位
            txt = cell.text if cell.text != '' else ' '
            row.append(txt)
        df.append(row)
        # 保存对应sheet
        pd.DataFrame(df).to_excel(writer, sheet_name=str(
            index), index=False, header=False)
        # 设置cell样式
        format = writer.book.add_format({'text_wrap': True})
        worksheet = writer.sheets[str(index)]
        # 设置row height
        worksheet.set_default_row(60)
        # 设置column width
        for idx, col in enumerate(pd.DataFrame(df)):
            worksheet.set_column(idx, idx, 30, format)
        # 缩小第一行和列长度设置
        worksheet.set_column(0, 0, 10)
        worksheet.set_row(0, 15)
    writer.save()
    writer.close()
