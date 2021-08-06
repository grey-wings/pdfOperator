import os
import PyQt5.QtWidgets as QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
import DesignerUI.MainCrackWindow
import DesignerUI.OperationCompleteDialog
import PDFOperations


class CrackCompletedWindow(QtWidgets.QDialog, DesignerUI.OperationCompleteDialog.Ui_Dialog):
    """
    操作完成提示框。
    """

    def __init__(self, parent=None):
        super(CrackCompletedWindow, self).__init__(parent)
        self.setWindowTitle('提示')
        self.setWindowIcon(QIcon('favicon4.ico'))
        self.setupUi(self)

        self.pushButton.clicked.connect(self.close)


class CrackFailedWindow(QtWidgets.QDialog, DesignerUI.OperationCompleteDialog.Ui_Dialog):
    """
    打开口令错误提示框。
    """
    def __init__(self, parent=None):
        super(CrackFailedWindow, self).__init__(parent)
        self.setWindowTitle('提示')
        self.setWindowIcon(QIcon('favicon4.ico'))
        self.setupUi(self)

        self.label.setText("操作失败\n请检查打开口令！")

        self.pushButton.clicked.connect(self.close)


class UserCrackMainDialog(QtWidgets.QDialog, DesignerUI.MainCrackWindow.Ui_Dialog):
    sourceFileNames = []
    targetFileNames = []

    def __init__(self, parent=None):
        super(UserCrackMainDialog, self).__init__(parent)
        self.setupUi(self)
        self.setWindowIcon(QIcon('favicon4.ico'))
        self.setWindowTitle("Cracker")

        self.exitBtn_2.clicked.connect(self.close)
        self.selectFileBtn.clicked.connect(self.openFiles)
        self.operateBtn_2.clicked.connect(self.crackOperate)

    def openFiles(self):
        self.sourceFileNames = QtWidgets.QFileDialog.getOpenFileNames(self,
                                                                      '打开',
                                                                      './',
                                                                      'PDF文档 (*.pdf)')[0]
        # getOpenFileNames支持一次选择多个文件，它返回一个元组。
        # 元组的第0个元素是一个列表，里面包括选择的每个文件的路径。
        # 第1个元素则是函参中设置的文件类型。
        s = '\n'.join(self.sourceFileNames)
        self.textEdit.setPlainText(s + '\n')

    def crackOperate(self):
        self.password = self.passwordEdit.toPlainText()
        for ifile in self.sourceFileNames:
            pos = ifile.rindex('/') + 1
            rawname = ifile[:pos] + "raw" + ifile[pos:-4] + ".pdf"
            os.rename(ifile, rawname)
            self.targetFileNames.append(rawname)
        flag = False
        for i in range(len(self.sourceFileNames)):
            if PDFOperations.pdf_Crack(self.targetFileNames[i],
                                       self.sourceFileNames[i],
                                       paswrd=self.password) == 0:
                completedDialog = CrackFailedWindow()
                completedDialog.setWindowModality(Qt.ApplicationModal)
                completedDialog.exec_()
            else:
                flag = True
        self.textEdit.setPlainText("")
        self.passwordEdit.setPlainText("")
        if flag:
            completedDialog = CrackCompletedWindow()
            completedDialog.setWindowModality(Qt.ApplicationModal)
            completedDialog.exec_()
