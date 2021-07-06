import sys
from PyQt5.QtWidgets import QApplication
import GUI


app = QApplication(sys.argv)
win = GUI.operateWindow()
win.show()
sys.exit(app.exec_())
input("please input any key to exit!")
