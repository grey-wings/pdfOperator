import sys
from PyQt5.QtWidgets import QApplication
import UserUI.UserMainWindow


app = QApplication(sys.argv)
win = UserUI.UserMainWindow.UserMainWindow()
win.show()
sys.exit(app.exec_())
