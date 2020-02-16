# free-hands-tips

> 不限技术，只为解决问题。

记录一些解放双手、提升效率的小技巧

## doc2xlsx

**批量转换Word（.doc）到 Excel**

`doc`文件转为`xlsx`，可以先转为`docx`或`html`，然后用[pandas](https://www.pypandas.cn/)包转

mac上利用`applescript`实现`doc`批量另存为`docx`或`html`

![批量Word转Excel](graph/doc2xlsx.gif)

> 第一次选择目录后需要点击授权目录权限。
>
> Word另存运行过程在后台，动图上没显示出来，实际Word窗口会一次次打开-另存-关闭。

代码详见：

- [doc2docx2xlsx](doc2docx2xlsx)
- [doc2html2xlsx](doc2html2xlsx)

> 对应目录下shell可直接运行，第一次运行可能需要安装对应python包
>
> `pip3 install pandas xlsxwriter docx Document`