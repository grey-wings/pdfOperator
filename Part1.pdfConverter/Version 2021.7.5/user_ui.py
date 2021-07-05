import sys
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QWidget, QDialog, QFileDialog

import Ui
import crackCompleteDialog
import crackSaveAs
import deleteDialog
import crackDialog
import mergeDialog
from pdfOperation import *


class completedWindow(QDialog, crackCompleteDialog.Ui_Dialog):
    """
    操作完成提示框。
    """

    def __init__(self, parent=None):
        super(completedWindow, self).__init__(parent)
        self.setWindowTitle('提示')
        self.setWindowIcon(QIcon('bitbug_favicon.ico'))
        self.setupUi(self)

        self.pushButton.clicked.connect(self.close)


class crackSaveAsWindow(QWidget, crackSaveAs.Ui_Dialog):
    """
    文件另存为路径选择界面。
    """
    fname = ""

    def __init__(self, parent=None):
        super(crackSaveAsWindow, self).__init__(parent)
        self.setWindowTitle('选择文件保存路径')
        self.setWindowIcon(QIcon('bitbug_favicon.ico'))
        self.setupUi(self)

        self.exitBtn.clicked.connect(self.close)
        self.selectFilleBtn.clicked.connect(self.openfile)
        # self.operateBtn.clicked.connect(self.crack)

    def openfile(self):
        self.fname = QFileDialog.getOpenFileName(self,
                                                 '打开',
                                                 './',
                                                 'PDF文档 (*.pdf)')[0]
        # getOpenFileName返回一个形式为('E:/SME/《C语言从入门到精通》.pdf', 'PDF文档 (*.pdf)')的元组
        self.textEdit.setPlainText(self.fname + '\n')


class mergeWindow(QDialog, mergeDialog.Ui_Dialog):
    """
    pdf合并功能界面。
    """
    sourceFileNames = []
    savePath = ""

    def __init__(self, parent=None):
        super(mergeWindow, self).__init__(parent)
        self.setWindowTitle('PDF破解')
        self.setWindowIcon(QIcon('bitbug_favicon.ico'))
        self.setupUi(self)

        self.exitBtn.clicked.connect(self.close)
        self.selectFileBtn.clicked.connect(self.openfile)
        self.selectPathBtn.clicked.connect(self.setSavePath)
        self.operateBtn.clicked.connect(self.merge)

    def openfile(self):
        self.sourceFileNames = QFileDialog.getOpenFileNames(self,
                                                            '打开',
                                                            './',
                                                            'PDF文档 (*.pdf)')[0]
        # getOpenFileNames支持一次选择多个文件，它返回一个元组。
        # 元组的第0个元素是一个列表，里面包括选择的每个文件的路径。
        # 第1个元素则是函参中设置的文件类型。
        s = '\n'.join(self.sourceFileNames)
        self.textEdit.setPlainText(s + '\n')

    def setSavePath(self):
        self.savePath = QFileDialog.getSaveFileName(self,
                                                    '打开',
                                                    './',
                                                    'PDF文档 (*.pdf)')[0]
        self.textEdit_2.setPlainText(self.savePath + '\n')

    def merge(self):
        PDF_Merge(self.sourceFileNames, self.savePath)
        completedDialog = completedWindow()
        completedDialog.setWindowModality(Qt.ApplicationModal)
        completedDialog.exec_()


class crackWindow(QDialog, crackDialog.Ui_Dialog):
    """
    pdf破解功能界面。
    """
    sourceFileName = ""
    savePath = ""
    password = ""

    def __init__(self, parent=None):
        super(crackWindow, self).__init__(parent)
        self.setWindowTitle('PDF破解')
        self.setWindowIcon(QIcon('bitbug_favicon.ico'))
        self.setupUi(self)

        self.exitBtn.clicked.connect(self.close)
        self.selectFileBtn.clicked.connect(self.openfile)
        self.selectPathBtn.clicked.connect(self.setSavePath)
        self.operateBtn.clicked.connect(self.crack)

    def openfile(self):
        self.sourceFileName = QFileDialog.getOpenFileName(self,
                                                          '打开',
                                                          './',
                                                          'PDF文档 (*.pdf)')[0]
        # getOpenFileName返回一个形式为('E:/SME/《C语言从入门到精通》.pdf', 'PDF文档 (*.pdf)')的元组
        self.textEdit.setPlainText(self.sourceFileName + '\n')

    def setSavePath(self):
        self.savePath = QFileDialog.getSaveFileName(self,
                                                    '打开',
                                                    './',
                                                    'PDF文档 (*.pdf)')[0]
        self.textEdit_2.setPlainText(self.savePath + '\n')

    def crack(self):
        self.password = self.passwordEdit.toPlainText()
        print(type(self.password))
        pdf_Crack(self.sourceFileName, self.savePath, paswrd=self.password)
        completedDialog = completedWindow()
        completedDialog.setWindowModality(Qt.ApplicationModal)
        completedDialog.exec_()


class deleteWindow(QDialog, deleteDialog.Ui_Dialog):
    """
    pdf页面删除功能界面。
    """
    sourceFileName = ""
    savePath = ""
    beginPage, endPage = 0, 0
    inputRange = ""

    def __init__(self, parent=None):
        super(deleteWindow, self).__init__(parent)
        self.setWindowTitle('PDF页面删除')
        self.setWindowIcon(QIcon('bitbug_favicon.ico'))
        self.setupUi(self)

        self.exitBtn.clicked.connect(self.close)
        self.selectFileBtn.clicked.connect(self.openfile)
        self.selectPathBtn.clicked.connect(self.setSavePath)
        self.operateBtn.clicked.connect(self.operateDelete)

    def openfile(self):
        self.sourceFileName = QFileDialog.getOpenFileName(self,
                                                          '打开',
                                                          './',
                                                          'PDF文档 (*.pdf)')[0]
        # getOpenFileName返回一个形式为('E:/SME/《C语言从入门到精通》.pdf', 'PDF文档 (*.pdf)')的元组
        self.textEdit.setPlainText(self.sourceFileName + '\n')

    def setSavePath(self):
        self.savePath = QFileDialog.getSaveFileName(self,
                                                    '打开',
                                                    './',
                                                    'PDF文档 (*.pdf)')[0]
        self.textEdit_2.setPlainText(self.savePath + '\n')

    def operateDelete(self):
        self.inputRange = self.textEdit_3.toPlainText()  # 获取文本框输入内容
        self.beginPage, self.endPage = (eval(i) for i in self.inputRange.split('-'))
        deletePages(self.sourceFileName, self.savePath, self.beginPage, self.endPage)
        completedDialog = completedWindow()
        completedDialog.setWindowModality(Qt.ApplicationModal)
        completedDialog.exec_()


class mainInterface(QDialog, Ui.Ui_Home):
    # 括号里是继承的父类，可以是object
    # object是python的默认类。python3默认加载object，即使没有写。
    def __init__(self, parent=None):
        super(mainInterface, self).__init__(parent)
        # 如果重写了init，调用子类时，不会调用父类的init
        # 如果重写了__init__ 时，要继承父类的构造方法，可以使用 super 关键字：
        # super(子类，self).__init__(参数1，参数2，....)

        # 设置窗口的标题
        self.setWindowTitle('Home')
        # 设置窗口的图标
        self.setWindowIcon(QIcon('bitbug_favicon.ico'))
        self.setupUi(self)

        # 点击按键弹出对话框
        # 子窗口和主窗口不能是同一类型（一般子窗口设成dialog），否则会崩溃
        self.mergeBtn.clicked.connect(self.showMergeDialog)
        self.crackerBtn.clicked.connect(self.showCrackDialog)
        self.pageDeleter.clicked.connect(self.showDeleteDialog)
        # 点击按钮退出界面
        # close后面不能加括号，否则报错
        self.exitBtn.clicked.connect(self.close)

    def showCrackDialog(self):
        crackDialogInstance = crackWindow()
        crackDialogInstance.resize(640, 580)
        crackDialogInstance.setWindowTitle('PDF去加密')
        crackDialogInstance.setWindowModality(Qt.ApplicationModal)
        crackDialogInstance.exec_()

    def showMergeDialog(self):
        mergeDialogInstance = mergeWindow()
        mergeDialogInstance.setWindowTitle('PDF合并')
        mergeDialogInstance.setWindowModality(Qt.ApplicationModal)
        mergeDialogInstance.exec_()

    def showDeleteDialog(self):
        deleteDialogInstance = deleteWindow()
        deleteDialogInstance.resize(640, 580)
        deleteDialogInstance.setWindowTitle('PDF页面删除')
        deleteDialogInstance.setWindowModality(Qt.ApplicationModal)
        deleteDialogInstance.exec_()
