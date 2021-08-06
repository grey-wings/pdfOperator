import os
import PyPDF2
import pikepdf
import pdfplumber
import pandas as pd

"""pdf破解相关函数"""


def pdf_Crack(sourcePath, savePath, paswrd=""):
    """
    去除没有打开口令的pdf加密。
    不能覆盖原文件，支持中文名称和路径。
    :param sourcePath: 要破解的pdf文件路径
    :param savePath: 生成的pdf文件路径
    :return:生成的pdf文件
    """
    # with pikepdf.open(sourcePath) as pdf:
    #     pdf.save(savePath)
    #     print('successful cracked')
    #     return savePath
    fp = open(sourcePath, "rb+")
    pdfFile = PyPDF2.PdfFileReader(fp)

    filepath, tempfilename = os.path.split(sourcePath)
    if pdfFile.isEncrypted:
        try:
            pdf = pikepdf.open(sourcePath, password=paswrd)
            pdf.save(savePath)
            return savePath
        except:
            return 0
    else:
        with pikepdf.open(sourcePath) as pdf:
            pdf.save(savePath)
            print('successful cracked')
            return savePath


def batch_pdf_Crack(folderPath):
    """
    将一个文件夹中的文件全部解密，在main.py目录下保存为原名字。
    支持带中文的文件名字。
    支持带中文的路径。
    :param filePath: 要解密的文件夹
    :param savePath: 文件保存的文件夹
    :return:
    """
    filelist = os.listdir(folderPath)
    # os.listdir() 方法用于返回指定的文件夹包含的文件或文件夹的名字的列表。
    for file in filelist:
        if file.endswith(".pdf") and ("~$" not in file):
            filePath = folderPath + "\\" + file
            with pikepdf.open(filePath) as pdf:
                pdf.save(file)
    print('successful cracked')


"""pdf内容提取相关函数"""


def getTable(sourcePath, bp=0, ep=-1):
    """
    从pdf文件中取得表格。
    :param sourcePath:
    :param bp:beginPage
    :param ep:endPage
    :return :所取得表格的dataframe形式
    """
    pdf = pdfplumber.open(sourcePath)
    if ep == -1:
        ep = len(pdf.pages) - 1
    result = pd.DataFrame()
    for i in range(bp - 1, ep):
        page = pdf.pages[i]
        table = page.extract_table()
        table = pd.DataFrame(table[1:], columns=table[0])
        pd.concat([result, table], axis=0) #这步有错
        
    # print(type(table))
    # print(table)
    return result


def writeExcel(tabledf, savePath, sheet='Sheet1'):
    """
    将表格dataframe保存到excel文件中。
    :param tabledf: dataframe变量
    :param savePath: 要写入的excel文件
    :param sheet: 要写入的sheet
    :return:
    """
    # 使用to_excel写入xls文件要安装openpyxl库
    tabledf.to_excel(savePath, sheet)
