from opn1502_form import *


# app = QtGui.QApplication(sys.argv)
app = QtWidgets.QApplication(sys.argv)
main_form = opn1502_main_ui()
main_form.show()
sys.exit(app.exec_())