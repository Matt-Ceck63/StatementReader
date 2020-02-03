##from fbs_runtime.application_context.PyQt5 import ApplicationContext
##from PyQt5.QtWidgets import QMainWindow
##
##import sys
##
##if __name__ == '__main__':
##    appctxt = ApplicationContext()       # 1. Instantiate ApplicationContext
##    window = QMainWindow()
##    window.resize(250, 150)
##    window.show()
##    exit_code = appctxt.app.exec_()      # 2. Invoke appctxt.app.exec_()
##    sys.exit(exit_code)

from fbs_runtime.application_context.PyQt5 import ApplicationContext
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox
from statementReader import StatementReader
from excelSaver import ExcelSaver
import pandas as pd
from os import environ
from pathlib import Path
import sys


class Ui_MainWindow(object):

    def get_data(self, path = 'statements.pdf'):

        reader = StatementReader(path)
        data = reader.retrieveTable(reader.get_page_text())
        data = reader.newLineClean(data)
        self.data = data

        print(data)
        return data

    def setupUi(self, MainWindow):

        self.excel_path = Path(environ["HOMEPATH"]  + "//OneDrive" + "//Documents" + "//Just Budget.xlsx") # Windows environ["ONEDRIVE"] + "\\Documents\\Just Budget.xlsx"

        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(953, 706)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")

        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")

        self.date_label = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(15)
        self.date_label.setFont(font)
        self.date_label.setObjectName("date_label")
        self.date_label.setAlignment(QtCore.Qt.AlignHCenter)
        self.horizontalLayout.addWidget(self.date_label)

        self.name_label = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(15)
        self.name_label.setFont(font)
        self.name_label.setObjectName("name_label")
        self.name_label.setAlignment(QtCore.Qt.AlignHCenter)
        self.horizontalLayout.addWidget(self.name_label)

        self.value_label = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(15)
        self.value_label.setFont(font)
        self.value_label.setObjectName("value_label")
        self.value_label.setAlignment(QtCore.Qt.AlignHCenter)
        self.horizontalLayout.addWidget(self.value_label)

        self.category_label = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(15)
        self.category_label.setFont(font)
        self.category_label.setObjectName("category_label")
        self.category_label.setAlignment(QtCore.Qt.AlignHCenter)
        self.horizontalLayout.addWidget(self.category_label)

        self.insert_label = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(15)
        self.insert_label.setFont(font)
        self.insert_label.setObjectName("insert_label")
        self.insert_label.setAlignment(QtCore.Qt.AlignHCenter)
        self.horizontalLayout.addWidget(self.insert_label)

        self.verticalLayout.addLayout(self.horizontalLayout)

        self.scrollArea = QtWidgets.QScrollArea(self.centralwidget)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setObjectName("scrollArea")

        self.scrollAreaWidgetContents = QtWidgets.QWidget()
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 1000, 800))
        self.scrollAreaWidgetContents.setObjectName("scrollAreaWidgetContents")

        self.gridLayout = QtWidgets.QGridLayout(self.scrollAreaWidgetContents)
        self.gridLayout.setObjectName("gridLayout")
        
        self.scrollArea.setWidget(self.scrollAreaWidgetContents)
        self.verticalLayout.addWidget(self.scrollArea)

        # Bottom Menu

        self.bottom_menu = QtWidgets.QGridLayout()
        self.verticalLayout.addLayout(self.bottom_menu)

        self.saveToButton = QtWidgets.QPushButton(self.centralwidget)
        self.saveToButton.setObjectName("ExcelButton")
        self.saveToButton.clicked.connect(self.open_excel)
        self.bottom_menu.addWidget(self.saveToButton, 0, 1, 1, 1, QtCore.Qt.AlignRight)

        self.openPdfButton = QtWidgets.QPushButton(self.centralwidget)
        self.openPdfButton.setObjectName("openPdfButton")
        self.openPdfButton.clicked.connect(self.open_pdf)
        self.bottom_menu.addWidget(self.openPdfButton, 1, 1, 1, 1, QtCore.Qt.AlignRight)

        self.SaveButton = QtWidgets.QPushButton(self.centralwidget)
        self.SaveButton.setObjectName("SaveButton")
        self.SaveButton.clicked.connect(self.save_to_excel)
        self.bottom_menu.addWidget(self.SaveButton, 2, 1, 1, 1, QtCore.Qt.AlignRight)

        self.save_to_label = QtWidgets.QLabel(self.centralwidget)
        self.save_to_label.setObjectName("save_to_label")
        self.bottom_menu.addWidget(self.save_to_label, 0, 0, 1, 1, QtCore.Qt.AlignLeft)

        MainWindow.setCentralWidget(self.centralwidget)

        # self.menubar = QtWidgets.QMenuBar(MainWindow)
        # self.menubar.setGeometry(QtCore.QRect(0, 0, 953, 21))
        # self.menubar.setObjectName("menubar")
        # MainWindow.setMenuBar(self.menubar)
        # self.statusbar = QtWidgets.QStatusBar(MainWindow)
        # self.statusbar.setObjectName("statusbar")
        # MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Statement Reader"))
        self.date_label.setText(_translate("MainWindow", "Date"))
        self.name_label.setText(_translate("MainWindow", "Name"))
        self.value_label.setText(_translate("MainWindow", "Value"))
        self.category_label.setText(_translate("MainWindow", "Category"))
        self.insert_label.setText(_translate("MainWindow", "Insert?"))
        self.SaveButton.setText(_translate("MainWindow", "Save"))
        self.openPdfButton.setText(_translate("MainWindow", "Open PDF"))
        self.saveToButton.setText(_translate("MainWindow", "Excel Path"))
        self.save_to_label.setText(_translate("MainWindow", "Path: " + str(self.excel_path)))

    def open_excel(self):
        path = Path(environ["HOMEPATH"] + "OneDrive" + "Documents" + "Just Budget.xlsx") # Windows environ["ONEDRIVE"] + "\\Documents\\"
        file_dialog = QtWidgets.QFileDialog()
        self.excel_path = Path(QtWidgets.QFileDialog.getOpenFileName(file_dialog, "Select Excel",
                                                                        str(path),
                                                                        "XLSX Files (*.xlsx)")[0])

        _translate = QtCore.QCoreApplication.translate
        self.save_to_label.setText(_translate("MainWindow", "Default path: " + str(self.excel_path)))


    def open_pdf(self):

        desktop_path = Path(environ["HOMEPATH"] + "/Desktop") # Windows environ["HOMEPATH"] + '\\Desktop'
        file_dialog = QtWidgets.QFileDialog()
        self.pdf_filepath = Path(QtWidgets.QFileDialog.getOpenFileName(file_dialog, "Open PDf",
                                                                  str(desktop_path),
                                                                  "PDF Files (*.pdf)")[0])
        print(self.pdf_filepath)
        self.data = self.get_data(self.pdf_filepath)
        try:
        	self.create_table()
        except FileNotFoundError:
        	msg = QMessageBox()
        	msg.setText("Excel file not found.")
        	msg.exec_()


    def save_to_excel(self):
        values = [v.text() for v in self.value_fields]
        combo_data = self.get_category_data(self.combo_boxes)
        check_data = self.get_check_data(self.check_boxes)
        file = ExcelSaver(self.excel_path)
        file.save(values, combo_data, check_data, self.months)

    @staticmethod
    def get_category_data(combo_list):
        result = list()
        for i in combo_list:
            result.append(i.currentText())
        return result

    @staticmethod
    def get_check_data(check_list):
        result = list()
        for i in check_list:
            result.append(i.isChecked())
        return result

    def get_categories(self, path):
        xls = pd.ExcelFile(path)
        last_sheet = len(xls.sheet_names) - 1 # -1 because sheet counting starts with 0
        df = pd.read_excel(xls, sheet_name=last_sheet)
        info = list(df["Month"].dropna().values)
        for s in ("Income", "Total Income", "Expenditure", "Total Expenditure", "Profit/Loss", "Closing balance"):
            info.remove(s)

        return info

    def get_full_month(self, date):
        full_months = ("September", "October", "November", "December", "January", "February", "March", "April", "May",
                       "June", "July", "August")
        short_months = ("Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug")

        if len(date) == 9:
            date = date[3:6]
        else:
            date = date[2:5]

        index = short_months.index(date)

        return full_months[index]

    def create_table(self):
        _translate = QtCore.QCoreApplication.translate
        combo_info = self.get_categories(self.excel_path)
        curr_month = ""
        self.months = []
        self.value_fields = []
        self.combo_boxes = []
        self.check_boxes = []
        row = 0
        col = 1

        for d in self.data:
            for e in d:
                if type(e) is list:
                    # e is not the date
                    for t in e:
                        col = 2
                        flag = 0
                        for n in t:
                            if flag == 0:  # it's a name create label
                                temp_label = QtWidgets.QLabel(self.scrollAreaWidgetContents)
                                temp_label.setObjectName(n)
                                self.gridLayout.addWidget(temp_label, row, col, 1, 1)
                                temp_label.setText(_translate("MainWindow", n))
                                flag = 1
                            elif flag == 1:  # it's a value create Line edit
                                temp_line = QtWidgets.QLineEdit(self.scrollAreaWidgetContents)
                                temp_line.setObjectName(n)
                                self.gridLayout.addWidget(temp_line, row, col, 1, 1, QtCore.Qt.AlignHCenter)
                                self.value_fields.append(temp_line)  # add a reference to the line edit object
                                temp_line.setFixedWidth(50)
                                temp_line.setText(n)
                                flag = 0
                            col += 1

                        # self.values.append(t[1])  # append the value of each transaction
                        self.months.append(self.get_full_month(curr_month))

                        temp_combo = QtWidgets.QComboBox(self.scrollAreaWidgetContents)
                        temp_combo.setObjectName(t[0] + "Combo")
                        temp_combo.insertItems(0, combo_info)
                        temp_combo.setEditable(True)
                        self.gridLayout.addWidget(temp_combo, row, col + 1, 1, 1)
                        self.combo_boxes.append(temp_combo)

                        temp_checkBox = QtWidgets.QCheckBox(self.scrollAreaWidgetContents)
                        temp_checkBox.setText("")
                        temp_checkBox.setObjectName(t[0] + "checkBox")
                        temp_checkBox.setChecked(True)
                        self.gridLayout.addWidget(temp_checkBox, row, col+2, 1, 1,
                                                  QtCore.Qt.AlignHCenter | QtCore.Qt.AlignTop)
                        self.check_boxes.append(temp_checkBox)

                        row += 1

                elif type(e) is str:
                    temp_label = QtWidgets.QLabel(self.scrollAreaWidgetContents)
                    temp_label.setObjectName(e)
                    self.gridLayout.addWidget(temp_label, row, col, 1, 1)
                    temp_label.setText(_translate("MainWindow", e))
                    curr_month = e

                col = 1

            # insert spacer
            empty_label = QtWidgets.QLabel(self.scrollAreaWidgetContents)
            empty_label.setObjectName("Empty")
            self.gridLayout.addWidget(empty_label, row, col, 1, 1)
            row += 1


if __name__ == "__main__":
    appctxt = ApplicationContext()

    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()

    exit_code = appctxt.app.exec_()
    sys.exit(exit_code)
