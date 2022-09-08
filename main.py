#xlrd version 1.2.0
#opencv-python version 4.4.0.46

import sys
import toolbar
import cv2
import xlrd
import math
import Image_rc
from pyzbar import pyzbar
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import Qt, QSize

#define global constant
EXCEL_PATH = 'C:/Users/cn222735/Desktop/Resignation Dashboard/test.xlsx'
row_count = 13

#create input keyboard
class NumInput(QWidget):

    def __init__(self):
        super(NumInput, self).__init__()
        self.initUI()
        self.center() #adjust to the center
        
    #define keyboard UI
    def initUI(self):

        #customise window title and icon
        self.setWindowTitle("Input Keyboard")
        self.setWindowIcon(QIcon(':KPMG_Logo.ico'))
        gridLayout = QGridLayout()

        #define and create input box
        self.display = QLineEdit("866") #set the prefix for all staff IDs as initial display value
        self.display.setFixedSize(QSize(1218, 200))
        self.display.setReadOnly(True)
        self.display.setAlignment(Qt.AlignRight)
        self.display.setFont(QFont("微软雅黑", 80, QFont.Bold))
        gridLayout.addWidget(self.display, 0, 0, 1, 3)

        #define key display
        keys = ['7', '8', '9',
                '4', '5', '6',
                '1', '2', '3',
                '0', '', 'del',
                'OK', '', 'Cancel']

        #define key position
        position = [(0, 0), (0, 1), (0, 2),
                    (1, 0), (1, 1), (1, 2),
                    (2, 0), (2, 1), (2, 2),
                    (3, 0), (3, 1), (3, 2),
                    (4, 0), (4, 1), (4, 2)]
        
        
        #define key font and size
        for item in range(len(keys)):
            btn = QPushButton(keys[item])
            btn.setFixedSize(QSize(400, 200))
            btn.setFont(QFont("微软雅黑", 70, QFont.Bold))
            btn.clicked.connect(self.btnClicked)
            if keys[item] == "0":
                gridLayout.addWidget(btn, 4, 0, 1, 2)
                btn.setFixedSize(QSize(809, 200))
            elif keys[item] == "Cancel":
                gridLayout.addWidget(btn, 5, 1, 1, 2)
                btn.setFixedSize(QSize(809, 200))
            elif keys[item] == "":
                continue
            else:
                gridLayout.addWidget(btn, position[item][0] + 1, position[item][1], 1, 1)
        self.setLayout(gridLayout)
    
    #define a function to adjust keyboard to the center of the screen
    def center(self):
        self.move(500,40)

    #define actions when keys are pressed
    def btnClicked(self):
        self.showNum = "866"
        sender = self.sender()
        symbols = ["del", "OK", "Cancel"]

        if sender.text() not in symbols:
            self.showNum += sender.text()
            self.display.setText(self.showNum)

        elif sender.text() == "del":
            self.display.setText("866")
            self.showNum = "866"

        elif sender.text() == "Cancel":
            self.close()

        elif sender.text() == "OK":

            self.found = []

             # use default path for the excel file 
            workbook = xlrd.open_workbook(EXCEL_PATH)

            # obtain sheet_name
            self.sheet_name = workbook.sheet_names()[0]

            # obtain sheet contents
            self.sheet = workbook.sheet_by_index(0)  # start from 0

            # obtain contents of cthe first column
            self.ids = self.sheet.col_values(0)

            #make all IDs integers
            for i, v in enumerate(self.ids):
                # skip first two values
                if i > 1:
                    self.ids[i] = str(int(v))

            #append to found list if an ID matches with keyboard input
            for i, v in enumerate(self.ids):
                if i > 0:
                    if (int(self.ids[i]) == int(self.display.text())):
                        self.found.append(i)
            
            #ID not matched
            if len(self.found) == 0:
                QMessageBox.warning(self, "Warning", "No Item Found Under this Staff ID")

            #ID matched
            else:
                self.close()
                self.ex = Excel(list = self.found, id = self.display.text(), sheet = self.sheet) #send arguments to class Excel
                self.ex.showFullScreen()

              
#create a new page to display values read from the excel file
class Excel(QMainWindow):

    def __init__(self, list, id, sheet): #receive arguments from class NumInput
        super(Excel, self).__init__()

        self.id = id
        self.list = list
        self.sheet = sheet

        self.setupUi(self)
        self.setWindowTitle("Item Possessed by " + str(self.id))
        self.setWindowIcon(QIcon(':KPMG_Logo.ico'))

        # define constant
        self.currentPage = 0
        self.pageSize = 8
        self.infoCols = 4
        self.fileName = ''
        self.Editable = False
        self.lastPageFlag = False

        #call functions to read excel file and display contents
        self.readExcel()
        self.getOnePage()
    
    #set up UI
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1244, 687)

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setEnabled(True)
        self.tableWidget.setGeometry(QtCore.QRect(250, 180, 1700, 1100))
        self.tableWidget.setRowCount(row_count) #set rows
        self.tableWidget.setColumnCount(6) #set columns

        #set cell size
        for i in range(row_count):
            self.tableWidget.setRowHeight(i, 200)
        
        #create header
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget.horizontalHeader().setStyleSheet('QHeaderView{font-size:30px};')
        self.tableWidget.setHorizontalHeaderLabels(['Staff ID', 'Model', 'Serial Number', 'Label', 'Item Status', 'Return Item'])
        
        #define exit button
        self.exitButton = QtWidgets.QPushButton(self.centralwidget)
        self.exitButton.setGeometry(QtCore.QRect(1730, 60, 100, 80))
        self.exitButton.setStyleSheet('QPushButton{font-size:30px};')
        self.exitButton.setObjectName("exitButton")
        
        #define copyright
        self.copyright = QtWidgets.QLabel(self.centralwidget)
        self.copyright.setGeometry(QtCore.QRect(900, 1320, 500, 41))
        self.copyright.setAlignment(QtCore.Qt.AlignCenter)
        self.copyright.setObjectName("copyright")

        #define total item text
        self.total_item = QtWidgets.QLabel(self.centralwidget)
        self.total_item.setGeometry(QtCore.QRect(900, 80, 500, 40))
        self.total_item.setStyleSheet('QLabel{font-size:30px};')
        self.total_item.setObjectName("total_item")

        #define number text
        self.total_item_num = QtWidgets.QLabel(self.centralwidget)
        self.total_item_num.setGeometry(QtCore.QRect(1350, 80, 31, 40))
        self.total_item_num.setStyleSheet('QLabel{font-size:30px};')
        self.total_item_num.setObjectName("total_item_num")

        MainWindow.setCentralWidget(self.centralwidget)

        #define and create menu bar
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1244, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)

        #define help icon
        self.help = QtWidgets.QMenu(self.menubar)
        self.help.setObjectName("Help")
        self.menubar.addAction(self.help.menuAction())

        #define and add user manual
        self.actiontips = QtWidgets.QAction(MainWindow)
        self.actiontips.setObjectName("actiontips")
        self.help.addAction(self.actiontips)

        self.retranslateUi()

        self.MainWindow = MainWindow

        # tie events to icons
        self.exitButton.clicked.connect(MainWindow.exit)
        self.actiontips.triggered.connect(MainWindow.openAbout)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    #add names 
    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
 
        __sortingEnabled = self.tableWidget.isSortingEnabled()
        self.tableWidget.setSortingEnabled(False)

        self.tableWidget.setSortingEnabled(__sortingEnabled)
        self.copyright.setText(_translate("MainWindow", "Copyright © 2022 KPMG Shanghai ITS Services Jacky Jiang"))
        self.total_item.setText(_translate("MainWindow", "Total Items Possessed by " + str(self.id) + " : "))
        self.total_item_num.setText(_translate("MainWindow", "0"))
        self.exitButton.setText(_translate("MainWindow", "EXIT"))
        self.help.setTitle(_translate("MainWindow", "Help"))
        self.actiontips.setText(_translate("MainWindow", "User Manual"))

    # generate a row for the table
    def generateRow(self, row, col, val, lastPageNum, lastPageFlag, trueRow):
        item = QtWidgets.QTableWidgetItem(val)
        item.setFont(QFont("微软雅黑", 15))
        item.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
        self.tableWidget.setItem(row, col, item)

        if row <= lastPageNum or lastPageFlag == False:
            # add scan button
            self.scanButton = QtWidgets.QPushButton('Scan')
            self.scanButton.setStyleSheet('QPushButton{margin:40px; font-size:30px};')
            self.scanButton.setObjectName("scanButton" + str(trueRow))
            self.tableWidget.setCellWidget(row, 5, self.scanButton)
            self.scanButton.clicked.connect(self.scan)
    
    #set item status
    def setStatus(self, row, col, status):
        self.Editable = True
        background_color = QColor()

        if status == True:
            item = QtWidgets.QTableWidgetItem("Returned")
            item.setFont(QFont("微软雅黑", 15))
            background_color.setNamedColor('green')

        else:
            item = QtWidgets.QTableWidgetItem("In Use")
            item.setFont(QFont("微软雅黑", 15))
            background_color.setNamedColor('red')

        self.tableWidget.setItem(row, col, item)
        item.setTextAlignment(QtCore.Qt.AlignCenter)
        item.setForeground(background_color)
               
 
    # set total item
    def setTotalItem(self, total_item):
        _translate = QtCore.QCoreApplication.translate
        self.total_item_num.setText(_translate("MainWindow", str(total_item)))


    # obtain values from cells
    def getTableWidgetItemContent(self, row, col):
        return self.tableWidget.item(row,col).text()
    
    #close current window
    def exit(self):
        msg = QMessageBox.information(self, "Information", "Are yuu sure to leave this page?", QMessageBox.Yes|QMessageBox.No,QMessageBox.No)
        if msg == QMessageBox.Yes:
            self.close()
            self.ex = NumInput()
            self.ex.show()

    # read Excel
    def readExcel(self):

        # obtain row length
        self.nrows = len(self.list)
        
        # obtain page number
        self.pageCount = math.ceil((self.nrows - 2) / self.pageSize)
        self.setTotalItem(len(self.list))   
        
        # obtain row numbers in last page
        self.lastPageCount = self.nrows - 2 - self.pageSize * (self.pageCount - 1)

    # recreate excel table 
    def getOnePage(self):
        self.Editable = False
        self.pageSize = len(self.list)
        for i in range(self.pageSize):
            for j in range(self.infoCols):
                if self.lastPageFlag and i >= self.lastPageCount:
                    val = ''
                else:
                    index = self.currentPage * self.pageSize + i + 1
                    val = self.sheet.cell_value(self.list[i], j)
                    
                    if isinstance(val, float):
                        if j == 0:
                            # remove zeros and decimal point
                            val = int(val)
                            # TableWidget needs string
                            val = str(val)
                        elif j == 2:
                            val = int(val)
                            val = str(val)
                        else:
                            val = str(val)
         
                # put in values
                self.generateRow(i, j, val, self.lastPageCount, self.lastPageFlag, index)
        
        #set inital status
        for i in range(len(self.list)):
                self.setStatus(i, 4, False)
                self.tableWidget.item(i, 4).setTextAlignment(Qt.AlignHCenter|Qt.AlignVCenter)

    # user manual
    def openAbout(self):
        # print('about...')
        QMessageBox.about(self,"User Manual Documentation","1. Please click the scan button to return the correspoding item.\n2. Once an item has been successfully scanned, the Item Status\n    column will record this item as Returned.\n3. After all items have been successfully scaned and recorded,\n    please click the submit button.")
    
    # scan QRcode or Bar code
    def scan(self):
    
        found = set()
        capture = cv2.VideoCapture(1)  # 0 for front camera, 1 for back camera

        while capture.isOpened():

            ret, frame = capture.read()

            camera = pyzbar.decode(frame)
        
    
            for tests in camera:
                # change code to string
                data = tests.data.decode('utf-8')
                testtype = tests.type
        
                # format into data and type
                printout = "{} ({})".format(data, testtype)

                if data not in found:
                    # print("[INFO] Found {} barcode: {}".format(testtype, data))
                    found.add(data)
                    
                    # cross check with values in table 
                    for i in range (len(self.list)):

                        # close camera and chaneg status if detected
                        if data == self.tableWidget.item(i,2).text():
                            found = True
                            capture.release()
                            cv2.destroyAllWindows()
                            self.setStatus(i, 4, True)
                            break

                        else:
                            found = False

                    if found == True:
                        QMessageBox.information(self, "Information", "Scanned Successfully")
                        self.check()
                    
                    else:
                        QMessageBox.warning(self, "Warning", "Item Not Found")
                
            cv2.imshow('Camera',frame)

            #stop camera by pressing  'q'
            if cv2.waitKey(1) == ord('q'):
                break

            #close camera by closing the window
            if cv2.getWindowProperty('Camera', cv2.WND_PROP_AUTOSIZE) < 1:
                break

        capture.release()
        cv2.destroyAllWindows()

    #check if all items have been successfully returned
    def check(self):
        checked = []
        for i in range(len(self.list)):
            if self.getTableWidgetItemContent(i,4) == "Returned":
                checked.append(i)
        if len(checked) == len(self.list):
            msg = QMessageBox.information(self, "Information", "All items have been successfully returned\nWould you like to stay on this page?", QMessageBox.Yes|QMessageBox.No,QMessageBox.No)
            if msg == QMessageBox.No:
                self.close()
                self.ex = NumInput()
                self.ex.show()

#define main dashboard
class Window(QMainWindow, toolbar.toolbar_UI):

    """Main Window."""
    def __init__(self):
        """Initializer.""" 
        super(Window, self).__init__() 
        self.initUI()

    def initUI(self):

        self.setWindowTitle("KPMG Resignation Dashboard")
        self.setWindowIcon(QIcon(':KPMG_Logo.ico'))
        self.resize(1000, 800)

        self.screen = QDesktopWidget().screenGeometry()
        self.setupUi(self)

        self.search_icon.triggered.connect(self.search)
        self.close_icon.triggered.connect(self.closeWindow)

    # set background
    def paintEvent(self, a0: QtGui.QPaintEvent) -> None:
        painter = QPainter(self)
        pixmap = QPixmap(':background')
        painter.drawPixmap(self.rect(), pixmap)

    # call input keyboard
    def search(self):
        self.ex = NumInput()
        self.ex.show()
    
    def closeWindow(self):
        msg = QMessageBox.information(self, "Information", "Would you like to close this window?", QMessageBox.Yes|QMessageBox.No,QMessageBox.No)
        if msg == QMessageBox.Yes:
            self.close()
        
        
    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = Window()
    win.showFullScreen()
    sys.exit(app.exec_())