import sys
from PyQt5 import QtCore, QtGui, QtWidgets
import Image_rc
from PyQt5.Qt import QSize

class toolbar_UI(object):

    def setupUi(self, MainWindow):

        #define and create toolbar
        self.tool_bar = QtWidgets.QToolBar(MainWindow)
        self.tool_bar.setMovable(False)
        MainWindow.addToolBar(QtCore.Qt.TopToolBarArea, self.tool_bar)
        self.tool_bar.setIconSize(QSize(300, 150))

        #define icon
        self.search_icon = QtWidgets.QAction(MainWindow)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(':search.png'), QtGui.QIcon.Normal, QtGui.QIcon.On)
        self.search_icon.setIcon(icon)

        #define close icon
        self.close_icon = QtWidgets.QAction(MainWindow)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(':close.png'), QtGui.QIcon.Normal, QtGui.QIcon.On)
        self.close_icon.setIcon(icon)

        #define space between icons
        self.space = QtWidgets.QWidget()
        self.space.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        
        #add icons to toolbar
        self.tool_bar.addAction(self.search_icon)
        self.tool_bar.addWidget(self.space)
        self.tool_bar.addAction(self.close_icon)
     
        #call function
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    #add names to all components
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "KPMG Resignation Dashboard"))
        self.tool_bar.setWindowTitle(_translate("MainWindow", "ToolBar"))
        self.search_icon.setText(_translate("MainWindow", "Search"))
        self.close_icon.setText(_translate("MainWindow", "Close"))
        

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = toolbar_UI()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

