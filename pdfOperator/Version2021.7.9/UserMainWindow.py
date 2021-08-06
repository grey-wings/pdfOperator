from PyQt5.QtCore import Qt
import PyQt5.QtWidgets as QtWidgets
from PyQt5.QtGui import QIcon
import DesignerUI.MainWindow
import UserMainCrackWindow, UserTableEdtractWindow


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

        self.action_2.triggered.connect(self.extractTables)
        self.action_3.triggered.connect(self.crack)

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
