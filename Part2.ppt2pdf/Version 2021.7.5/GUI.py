from PyQt5.QtWidgets import QDialog, QFileDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
import crackCompleteDialog
import convert
import maingui


class completedWindow(QDialog, crackCompleteDialog.Ui_Dialog):
    """
    操作完成提示框。
    """

    def __init__(self, parent=None):
        super(completedWindow, self).__init__(parent)
        self.setWindowTitle('提示')
        self.setWindowIcon(QIcon('favicon.ico'))
        self.setupUi(self)

        self.pushButton.clicked.connect(self.close)


class operateWindow(QDialog, maingui.Ui_Dialog):
    sourceFileNames = []
    savePath = []
    powerpoint = convert.init_powerpoint()

    def __init__(self, parent=None):
        super(operateWindow, self).__init__(parent)
        self.setWindowTitle('PDF页面删除')
        self.setWindowIcon(QIcon('favicon.ico'))
        self.setupUi(self)

        self.selectFileBtn.clicked.connect(self.openfile)
        self.selectFileBtn_2.clicked.connect(self.setSavePath)
        self.operateBtn.clicked.connect(self.operate)

    def openfile(self):
        self.sourceFileNames = QFileDialog.getOpenFileNames(self,
                                                            '打开',
                                                            './',
                                                            'ppt文档 (*.ppt;*.pptx)')[0]
        # getOpenFileNames支持一次选择多个文件，它返回一个元组。
        # 元组的第0个元素是一个列表，里面包括选择的每个文件的路径。
        # 第1个元素则是函参中设置的文件类型。
        s = '\n'.join(self.sourceFileNames)
        self.textEdit.setPlainText(s + '\n')
        self.sourceFileNames = [st.replace('/', '\\') for st in self.sourceFileNames]
        # 如果路径包含'/'而不是'\\'，会报错：PowerPoint 无法将 ^0 保存到 ^1

    def setSavePath(self):
        self.savePath = QFileDialog.getSaveFileName(self,
                                                    '打开',
                                                    './',
                                                    'pdf文档 (*.pdf)')[0]
        self.textEdit_2.setPlainText(self.savePath + '\n')
        self.savePath = [st.replace('/', '\\') for st in self.savePath]

    def operate(self):
        convert.batch_convert(powerpoint=self.powerpoint,
                              fileList=self.sourceFileNames,
                              saveList=self.savePath)
        completedDialog = completedWindow()
        completedDialog.setWindowModality(Qt.ApplicationModal)
        completedDialog.exec_()