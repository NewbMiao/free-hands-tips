# only work in windows
# python -m pip install pypiwin32
from win32com import client as wc
# 打开word应用程序
word = wc.Dispatch("Word.Application") 
files=["bak/docs/1.doc"]
for file in files:
    # 打开word文件
    doc = word.Documents.Open(file)
    # 另存为后缀为".docx"的文件，其中参数12指docx文件 
    doc.SaveAs("{}x".format(file), 12)
    doc.Close()
word.Quit()
