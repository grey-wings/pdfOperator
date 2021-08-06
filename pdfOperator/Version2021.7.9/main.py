import sys
from PyQt5.QtWidgets import QApplication
import UserMainWindow


app = QApplication(sys.argv)
win = UserMainWindow.UserMainWindow()
win.show()
sys.exit(app.exec_())
# import PDFOperations
# df = PDFOperations.getTable("C:\\Users\\15594\\Desktop\\电子类-四川赛区（第一场）获奖名单.pdf")
# PDFOperations.writeExcel(df, "C:\\Users\\15594\\Desktop\\aaa.xlsx")
