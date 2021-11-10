import PyQt5.QtWidgets as QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from PyQt5 import QtCore
import PDFOperations

from DesignerUI import ppt2pdfWindow


class UserWord2pdfDialog(QtWidgets.QDialog, ppt2pdfWindow.Ui_Dialog):
    sourceFileNames = []
    savePath = []

    def __init__(self, parent=None):
        super(UserWord2pdfDialog, self).__init__(parent)
        self.setupUi(self)
        self.setWindowIcon(QIcon('favicon2.ico'))
        self.setWindowTitle("Word2pdf")
        _translate = QtCore.QCoreApplication.translate
        self.label_5.setText(_translate("Dialog", "将word文件批量转换为pdf文件。\n"
                                                  "1. 生成的word文件名与原word文件相同。\n"
                                                  "2. 保存路径不选默认为原路径（要求所选的所有word路径相同）。"))

        self.operateBtn.clicked.connect(self.operate)
        self.exitBtn.clicked.connect(self.close)
        self.selectFileBtn.clicked.connect(self.openFile)
        self.selectPathBtn.clicked.connect(self.setSavePath)

    def openFile(self):
        self.filetype = 'Word 档(*.docx);;' \
                        '启用宏的 Word 文档(*.docm);;' \
                        'Word 97-2003 演示文稿(*.doc);;'
        slist = QtWidgets.QFileDialog.getOpenFileNames(self,
                                                       '打开',
                                                       './',
                                                       self.filetype)[0]
        # getOpenFileNames支持一次选择多个文件，它返回一个元组。
        # 元组的第0个元素是一个列表，里面包括选择的每个文件的路径。
        # 第1个元素则是函参中设置的文件类型。
        self.sourceFileNames.extend(slist)
        self.sourceFileNames = list({}.fromkeys(self.sourceFileNames).keys())
        self.sourceFileNames = [st.replace('/', '\\') for st in self.sourceFileNames]
        s = '\n'.join(self.sourceFileNames)
        self.textEdit.setPlainText(s + '\n')
        # 如果路径包含'/'而不是'\\'，会报错：PowerPoint 无法将 ^0 保存到 ^1

    def setSavePath(self):
        self.savePath = QtWidgets.QFileDialog.getExistingDirectory(self,
                                                                   '选取文件夹',
                                                                   './')
        self.textEdit_2.setPlainText(self.savePath)
        self.savePath = self.savePath.replace('/', '\\')

    def operate(self):
        PDFOperations.batch_word2PDF(fileList=self.sourceFileNames,
                                     savePath=self.savePath)
        completeDialog = QtWidgets.QMessageBox.information(self, '提示',
                                                           "操作完成！")
        self.sourceFileNames = []
        self.textEdit.setPlainText('')
        self.savePath = ''
        self.textEdit_2.setPlainText('')
