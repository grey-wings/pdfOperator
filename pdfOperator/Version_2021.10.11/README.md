# 2021.10.11版本  

已实现功能：  
1.去除pdf的限制，比如禁止打印、禁止复制等 对于有打开口令的pdf，只有知道打开口令才能去除限制，并且也能去除打开口令。  
2.ppt转pdf  

待实现功能：  
1.提取PDF中的文字和表格，该功能已经完成，待测试  
2.word转pdf，有设想但没有实现  
3.将一个文件夹中的图片转换和合并成pdf 初版用于转化笔记照片，因此进行了二值化处理和添加背景色，可能不具有普遍适用性，需要根据功能进行进一步改进。  

目前没有发布exe格式，复现需要自行搭建python环境  
复现需要安装的库（不保证统计完全，以pycharm中搜索的库为准）：  
1.pyqt5-tools  
2.pyuic5-tool  
3.PyPDF2  
4.comtypes  
5.pikepdf  
6.pdfplumber  
7.pandas  
8.python-docx   
9.xlwt  
