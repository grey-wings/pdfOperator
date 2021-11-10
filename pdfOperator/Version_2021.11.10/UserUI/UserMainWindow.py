import PyQt5.QtWidgets as QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon

import DesignerUI.MainWindow
from UserUI import UserTextEdtractWindow, UserMainCrackWindow, UserMainCrackWindow, \
    UserTableEdtractWindow, UserPPT2pdfWindow, UserWord2pdfWindow


class UserMainWindow(QtWidgets.QMainWindow, DesignerUI.MainWindow.Ui_MainWindow):
    """
        pdf破解功能界面。
        """
    sourceFileName = ""
    savePath = ""
    password = ""

    def __init__(self, parent=None):
        super(UserMainWindow, self).__init__(parent)
        self.setupUi(self)
        self.setWindowIcon(QIcon('favicon1.ico'))
        self.setWindowTitle("PDFOperator")

        self.action.triggered.connect(self.extractText)
        self.action_2.triggered.connect(self.extractTables)
        self.action_3.triggered.connect(self.crack)
        self.actionPPT_PDF.triggered.connect(self.ppt2pdf)
        self.actionWord_PDF.triggered.connect(self.word2pdf)

    def crack(self):
        crackDialogInstance = UserMainCrackWindow.UserCrackMainDialog()
        crackDialogInstance.resize(800, 600)
        crackDialogInstance.setWindowModality(Qt.ApplicationModal)
        crackDialogInstance.exec_()

    def extractTables(self):
        tableExtDialogInstance = UserTableEdtractWindow.UserTableExtractDialog()
        tableExtDialogInstance.resize(962, 835)
        tableExtDialogInstance.setWindowModality(Qt.ApplicationModal)
        tableExtDialogInstance.exec_()

    def extractText(self):
        textExtDialogInstance = UserTextEdtractWindow.UserTextExtractDialog()
        textExtDialogInstance.resize(962, 835)
        textExtDialogInstance.setWindowModality(Qt.ApplicationModal)
        textExtDialogInstance.exec_()

    def ppt2pdf(self):
        p2pDialogInstance = UserPPT2pdfWindow.UserPPT2pdfDialog()
        p2pDialogInstance.resize(851, 732)
        p2pDialogInstance.setWindowModality(Qt.ApplicationModal)
        p2pDialogInstance.exec_()

    def word2pdf(self):
        w2pDialogInstance = UserWord2pdfWindow.UserWord2pdfDialog()
        w2pDialogInstance.resize(851, 732)
        w2pDialogInstance.setWindowModality(Qt.ApplicationModal)
        w2pDialogInstance.exec_()
