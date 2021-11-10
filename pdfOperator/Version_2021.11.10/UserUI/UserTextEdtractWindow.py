import os
import PyQt5.QtWidgets as QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from DesignerUI import OperationCompleteDialog, TextExtractWindow
import PDFOperations


class ExtCompletedWindow(QtWidgets.QDialog, OperationCompleteDialog.Ui_Dialog):
    """
    操作完成提示框。
    """

    def __init__(self, parent=None):
        super(ExtCompletedWindow, self).__init__(parent)
        self.setWindowTitle('提示')
        self.setWindowIcon(QIcon('favicon3.ico'))
        self.setupUi(self)

        self.pushButton.clicked.connect(self.close)


class UserTextExtractDialog(QtWidgets.QDialog, TextExtractWindow.Ui_Dialog):
    sourceFileName = None
    savePath = None
    beginPage = 1
    endPage = -1
    uformat = ''

    def __init__(self, parent=None):
        super(UserTextExtractDialog, self).__init__(parent)
        self.setupUi(self)
        self.setWindowIcon(QIcon('favicon3.ico'))
        self.setWindowTitle("TextExtracter")

        self.exitBtn.clicked.connect(self.close)
        self.selectFileBtn.clicked.connect(self.openFile)
        self.selectPathBtn.clicked.connect(self.setSavePath)
        self.operateBtn.clicked.connect(self.extOperate)

    def openFile(self):
        self.sourceFileName = QtWidgets.QFileDialog.getOpenFileName(self,
                                                                    '打开',
                                                                    './',
                                                                    'PDF文档 (*.pdf)')[0]
        self.textEdit.setPlainText(self.sourceFileName)

    def setSavePath(self):
        self.uformat = "Microsoft Word文件 (*.docx);;" \
                       "文本文档 (*.txt)"
        self.savePath = QtWidgets.QFileDialog.getSaveFileName(self,
                                                              '打开',
                                                              './',
                                                              self.uformat)[0]
        self.textEdit_2.setPlainText(self.savePath)

    def extOperate(self):
        self.beginPage = eval(self.beginPageEdit.toPlainText())
        self.endPage = eval(self.endPageEdit.toPlainText())
        s = PDFOperations.getText(self.sourceFileName,
                                  self.beginPage,
                                  self.endPage)
        PDFOperations.writeText('\n\r' + s, self.savePath)
        completedDialog = ExtCompletedWindow()
        completedDialog.setWindowModality(Qt.ApplicationModal)
        completedDialog.exec_()
