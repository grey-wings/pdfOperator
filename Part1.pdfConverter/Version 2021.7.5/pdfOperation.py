import os
import pikepdf
import PyPDF2
import pdfplumber
from PIL import Image


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
            print("error!")
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


# def split_PDF_pages(filePath, savePath, pageEnd, pageBeg):
#     """
#     将pdf的第pageBeg到第pageEnd页截取出来,保存至savePath.
#     :param filePath:
#     :param pageBeg:
#     :param pageEnd:
#     :return:
#     """
#     # 未成功
#     with pikepdf.open(filePath) as pdf:
#         dst = pikepdf.new()
#         for n, page in enumerate(pdf.pages[pageBeg-1:pageEnd]):
#             dst.pages.append(page)
#         dst.save(savePath)


def deletePages(filePath, savePath, begin, end):
    """
    删除pdf的第begin页到第end页，第end页也被删除，计数从1开始。
    :param filePath: 要删除页面的文件
    :param begin: 开始删除的页面（第1页为1）
    :param end: ……
    :return:
    """
    with pikepdf.open(filePath) as pdf:
        del pdf.pages[begin - 1:end]
        pdf.save(savePath)


def PDF_Merge(fileList, savePath):
    """
    合并几个
    :param fileList: 要合并的pdf文件名组成的列表
    :param savePath: 保存路径和文件名
    :return:
    """
    pdf = pikepdf.new()
    for file in fileList:
        src = pikepdf.open(file)
        pdf.pages.extend(src.pages)
    pdf.save(savePath)


def pdf_get_text(FilePath, SavePath):
    """
    提取一个pdf文档中的文本，保存到一个txt里面。
    :param FilePath: pdf文档路径
    :param SavePath: txt文档路径（不存在将被创建；若已有内容则添加在其后）
    :return:
    """
    pdf = pdfplumber.open(FilePath)
    with open(SavePath, 'a', encoding='utf-8') as f:
        for page in pdf.pages:
            text = page.extract_text()
            f.write(text)
            f.write("\n\r")


def clear_brackets(FilePath, SavePath):
    """
    将文本中两个括号之间的内容全部去掉。
    只支持英文括号。

    应当更新正则的版本。
    :param FilePath:
    :param SavePath:
    :return:
    """
    with open(FilePath, 'r+', encoding='utf-8') as f1:
        text = f1.read()
    with open(SavePath, 'w', encoding='utf-8') as f2:
        outText = ""
        i = 0
        while i < len(text):
            outText += text[i]
            if text[i] == '(':
                j = i + 1
                # ch = text[j]
                while text[j] != ')':
                    outText += ' '
                    # ch = text[j]
                    j += 1
                i = j - 1
            i += 1
        f2.write(outText)


def pictures2PDF(path, pdf_name):
    """
    将一个文件夹中的所有图片转化和合并成一个pdf
    :param path: 文件夹路径
    :param pdf_name: pdf文件路径
    :return:
    """
    ## change all png into jpg & delete the .png files
    names = os.listdir(path)
    for name in names:
        img = Image.open(path + '/' + name)
        name = name.split(".")
        if name[-1] == "png":
            name[-1] = "jpg"
            name_jpg = str.join(".", name)
            r, g, b, a = img.split()
            img = Image.merge("RGB", (r, g, b))
            to_save_path = path + '/' + name_jpg
            img.save(to_save_path)
            os.remove(path + '/' + name[0] + '.png')
        else:
            continue

    ## add jpg and jpeg to
    file_list = os.listdir(path)

    pic_name = []
    im_list = []
    for x in file_list:
        if "jpg" in x or 'jpeg' in x:
            pic_name.append(x)

    pic_name.sort()  # sorted
    new_pic = []

    for x in pic_name:
        if "jpg" in x:
            new_pic.append(x)

    im1 = Image.open(os.path.join(path, new_pic[0]))
    new_pic.pop(0)
    for i in new_pic:
        img = Image.open(os.path.join(path, i))
        # im_list.append(Image.open(i))
        if img.mode == "RGBA":
            r, g, b, a = img.split()
            img = Image.merge("RGB", (r, g, b))
            img = img.convert('RGB')
            im_list.append(img)
        else:
            im_list.append(img)

    im1.save(pdf_name, "PDF", resolution=100.0, save_all=True, append_images=im_list)
    print("输出文件名称：", pdf_name)


# if __name__ == '__main__':
#     # input the name for pdf like xxx.pdf
#     pdf_name = 'image2pdf.pdf'
#     mypath = 'image'
#     if ".pdf" in pdf_name:
#         rea(mypath, pdf_name=pdf_name)
#     else:
#         rea(mypath, pdf_name="{}.pdf".format(pdf_name))