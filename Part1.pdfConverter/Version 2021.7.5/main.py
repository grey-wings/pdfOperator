from PyQt5.QtWidgets import QApplication
from user_ui import *
from pdfOperation import *
from Transform import *


app = QApplication(sys.argv)
win = mainInterface()
win.show()
sys.exit(app.exec_())
