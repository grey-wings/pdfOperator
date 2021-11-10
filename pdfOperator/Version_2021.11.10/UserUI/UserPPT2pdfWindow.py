import PyQt5.QtWidgets as QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
import PDFOperations

from DesignerUI import ppt2pdfWindow, OperationCompleteDialog


class UserPPT2pdfDialog(QtWidgets.QDialog, ppt2pdfWindow.Ui_Dialog):
    sourceFileNames = []
    savePath = []
    powerpoint = PDFOperations.init_powerpoint()

    def __init__(self, parent=None):
        super(UserPPT2pdfDialog, self).__init__(parent)
        self.setupUi(self)
        self.setWindowIcon(QIcon('favicon2.ico'))
        self.setWindowTitle("ppt2pdf")

        self.operateBtn.clicked.connect(self.operate)
        self.exitBtn.clicked.connect(self.close)
        self.selectFileBtn.clicked.connect(self.openFile)
        self.selectPathBtn.clicked.connect(self.setSavePath)

    def openFile(self):
        self.filetype = 'PowerPoint 演示文稿(*.pptx);;' \
                        '启用宏的 PowerPoint 演示文稿(*.pptm);;' \
                        'PowerPoint 97-2003 演示文稿(*.ppt)'
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

        PDFOperations.batch_ppt2PDF(powerpoint=self.powerpoint,
                                    fileList=self.sourceFileNames,
                                    savePath=self.savePath)
        completeDialog = QtWidgets.QMessageBox.information(self, '提示',
                                                           "操作完成！")
        self.sourceFileNames = []
        self.textEdit.setPlainText('')
        self.savePath = ''
        self.textEdit_2.setPlainText('')

