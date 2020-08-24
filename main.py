from PyQt5 import QtWidgets, Qt, QtGui, QtCore
from PyQt5.Qt import *
from PyQt5 import QtSql
import main
import sys
import os
import sqlite3
import subprocess
import pandas as pd
from openpyxl import load_workbook
import webbrowser
import win32com.client as win32


#Open Excel_file
try:
    wb_val = load_workbook(filename='Work_Hours.xlsx',
                           data_only=True)  # Open filename.xlsx

    sheet_val = wb_val['Hours']  # Get page "Hours"

    df = pd.DataFrame({sheet_val['A1'].value: [sheet_val['A2'].value, sheet_val['A3'].value],
                       sheet_val['B1'].value: [sheet_val['B2'].value, sheet_val['B3'].value],
                       sheet_val['C1'].value: [sheet_val['C2'].value, sheet_val['C3'].value, ]})

    bonus_day = pd.DataFrame({"A": [sheet_val['A19'].value, sheet_val['A20'].value, sheet_val['A21'].value],
                              "B": [sheet_val['B19'].value, sheet_val['B20'].value, sheet_val['B21'].value]})
except:
    pass



#Create AbstarctModel for display our Excel_file
class pandasModel(QAbstractTableModel):

    def __init__(self, data):
        QAbstractTableModel.__init__(self)
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, paren=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None


class mywindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(mywindow, self).__init__()
        self.ui = main.Ui_Form()
        self.ui.setupUi(self)

        conn = sqlite3.connect('filials.db')
        c = conn.cursor()

        #Create table admins
        c.execute("""CREATE TABLE IF NOT EXISTS admins(
            f_name text,
            filial text,
            data_start text,
            data_finish text)
            """),
        #Create table filials
        c.execute("""CREATE TABLE IF NOT EXISTS filials(
            PRM_key oid,
            filial text,
            print_server text,
            device_lock_server text)
            """)


class Ui_Form(object):

    def setupUi(self, Form):
        #Form
        Form.setObjectName("Form")
        Form.setFixedSize(864, 511)
        self.tabWidget = QtWidgets.QTabWidget(Form)
        self.tabWidget.setGeometry(QtCore.QRect(9, 9, 851, 491))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(10)
        self.tabWidget.setFont(font)
        self.tabWidget.setObjectName("tabWidget")

        #Tab2 Cities
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        #Search filial
        self.TextEdit_search = QtWidgets.QTextEdit(self.tab_2)
        self.TextEdit_search.setReadOnly(True)
        self.TextEdit_search.setGeometry(QtCore.QRect(20, 60, 391, 371))
        self.TextEdit_search.setObjectName("TextEdit_search")
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(10)
        self.TextEdit_search.setFont(font)
        self.btn_search = QtWidgets.QPushButton(self.tab_2)
        self.btn_search.setGeometry(QtCore.QRect(310, 20, 101, 31))
        self.btn_search.setObjectName("btn_search")
        self.btn_search.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Bold))
        self.search_box = QtWidgets.QLineEdit(self.tab_2)
        self.search_box.setGeometry(QtCore.QRect(20, 20, 211, 21))
        self.search_box.setObjectName("search_box")
        #Query list admins
        self.query_btn_2 = QtWidgets.QPushButton(self.tab_2)
        self.query_btn_2.setGeometry(QtCore.QRect(680, 20, 131, 31))
        self.query_btn_2.setObjectName("query_btn_2")
        self.query_btn_2.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Bold))
        self.TextEdit_query2 = QtWidgets.QTextEdit(self.tab_2)
        self.TextEdit_query2.setGeometry(QtCore.QRect(450, 60, 361, 371))
        self.TextEdit_query2.setReadOnly(True)
        self.TextEdit_query2.setObjectName("TextEdit_query2")
        self.label_discription2 = QtWidgets.QLabel(self.tab_2)
        self.label_discription2.setGeometry(QtCore.QRect(450, 30, 230, 20))
        self.label_discription2.setObjectName("label_discription2")
        self.tabWidget.addTab(self.tab_2, "")

        #Tab3 Work_hours
        self.tab_3 = QtWidgets.QWidget()
        self.tab_3.setObjectName("tab_3")
        self.toolBox = QtWidgets.QToolBox(self.tab_3)
        self.toolBox.setGeometry(QtCore.QRect(10, 5, 821, 421))
        self.toolBox.setObjectName("toolBox")
        #Page 1 (Work hours)
        self.page = QtWidgets.QWidget()
        self.page.setGeometry(QtCore.QRect(10, 10, 821, 391))
        self.page.setObjectName("page")
        self.TableView_1 = QtWidgets.QTableView(self.page)
        self.TableView_1.setGeometry(QtCore.QRect(0, 0, 821, 358))
        self.TableView_1.setObjectName("TableView_1")
        self.toolBox.addItem(self.page, "")
        self.tabWidget.addTab(self.tab_3, "")
        #Page 2 (days off)
        self.page_2 = QtWidgets.QWidget()
        self.page_2.setGeometry(QtCore.QRect(10, 10, 821, 391))
        self.page_2.setObjectName("page_2")
        self.TableView_2 = QtWidgets.QTableView(self.page_2)
        self.TableView_2.setGeometry(QtCore.QRect(0, 0, 821, 358))
        self.TableView_2.setObjectName("TableView_2")
        self.toolBox.addItem(self.page_2, "")
        #Button show
        self.btn_graphic = QtWidgets.QPushButton(self.tab_3)
        self.btn_graphic.setGeometry(QtCore.QRect(700, 430, 131, 30))
        self.btn_graphic.setObjectName("btn_graphic")
        self.btn_graphic.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Bold))

        #Button feedback
        self.btn_feedback = QtWidgets.QPushButton(Form)
        self.btn_feedback.setGeometry(QtCore.QRect(757, 5, 100, 25))
        self.btn_feedback.setObjectName("btn_feedback")

        #Tab 3 "Admin"
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")

        #labels
        #label name Administrator
        self.f_name_label = QtWidgets.QLabel(self.tab)
        self.f_name_label.setGeometry(QtCore.QRect(20, 40, 151, 21))
        self.f_name_label.setObjectName("f_name_label")
        #label City
        self.filial_label = QtWidgets.QLabel(self.tab)
        self.filial_label.setGeometry(QtCore.QRect(20, 90, 71, 21))
        self.filial_label.setObjectName("filial_label")
        #label Data start vacation
        self.data_start_label = QtWidgets.QLabel(self.tab)
        self.data_start_label.setGeometry(QtCore.QRect(20, 140, 131, 21))
        self.data_start_label.setObjectName("data_start_label")
        #label Data finish vacation
        self.data_finish_label = QtWidgets.QLabel(self.tab)
        self.data_finish_label.setGeometry(QtCore.QRect(20, 190, 151, 21))
        self.data_finish_label.setObjectName("data_finish_label")
        #label select ID
        self.label_5 = QtWidgets.QLabel(self.tab)
        self.label_5.setGeometry(QtCore.QRect(20, 300, 151, 21))
        self.label_5.setObjectName("label_5")
        #label list admins in vacation
        self.label_discription = QtWidgets.QLabel(self.tab)
        self.label_discription.setGeometry(QtCore.QRect(460, 40, 230, 21))
        self.label_discription.setObjectName("label_discription")
        self.tabWidget.addTab(self.tab, "")

        #TextEdit list admins
        self.Label_query = QtWidgets.QTextEdit(self.tab)
        self.Label_query.setGeometry(QtCore.QRect(460, 70, 361, 251))
        self.Label_query.setReadOnly(True)
        self.Label_query.setObjectName("Label_query")

        #Inputs
        #input name Administrator
        self.f_name = QtWidgets.QLineEdit(self.tab)
        self.f_name.setGeometry(QtCore.QRect(200, 40, 231, 21))
        self.f_name.setObjectName("f_name")
        #input City
        self.filial = QtWidgets.QLineEdit(self.tab)
        self.filial.setGeometry(QtCore.QRect(200, 90, 231, 21))
        self.filial.setObjectName("filial")
        #input Data start vacation
        self.data_start = QtWidgets.QLineEdit(self.tab)
        self.data_start.setGeometry(QtCore.QRect(200, 140, 231, 21))
        self.data_start.setObjectName("data_start")
        #input Data finish vacation
        self.data_finish = QtWidgets.QLineEdit(self.tab)
        self.data_finish.setGeometry(QtCore.QRect(200, 190, 231, 21))
        self.data_finish.setObjectName("data_finish")
        #input Delete record
        self.delete_box = QtWidgets.QLineEdit(self.tab)
        self.delete_box.setGeometry(QtCore.QRect(200, 300, 231, 21))
        self.delete_box.setText("")
        self.delete_box.setObjectName("delete_box")

        #Buttons
        #Submit btn
        self.submit_btn = QtWidgets.QPushButton(self.tab)
        self.submit_btn.setGeometry(QtCore.QRect(20, 240, 141, 31))
        self.submit_btn.setObjectName("submit_btn")
        self.submit_btn.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Bold))
        #Show records
        self.query_btn = QtWidgets.QPushButton(self.tab)
        self.query_btn.setGeometry(QtCore.QRect(680, 350, 141, 31))
        self.query_btn.setObjectName("query_btn")
        self.query_btn.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Bold))
        #Delete records
        self.delete_btn = QtWidgets.QPushButton(self.tab)
        self.delete_btn.setGeometry(QtCore.QRect(20, 350, 141, 31))
        self.delete_btn.setObjectName("delete_btn")
        self.delete_btn.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Bold))
        #btn open file
        self.btn_open_file = QtWidgets.QPushButton(self.tab)
        self.btn_open_file.setGeometry(QtCore.QRect(460, 350, 141, 31))
        self.btn_open_file.setObjectName("btn_open_excel")
        self.btn_open_file.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Bold))

        #Tab 4 "Control panel"
        self.tab_4 = QtWidgets.QWidget()
        self.tab_4.setObjectName("tab_4")
        self.tabWidget.addTab(self.tab_4, "")
        #Label name_PC
        self.Input_PC_label = QtWidgets.QLabel(self.tab_4)
        self.Input_PC_label.setGeometry(QtCore.QRect(620, 20, 101, 20))
        self.Input_PC_label.setObjectName("Input_PC_label")
        #Input PC_name_or_IP
        self.lbl = QtWidgets.QLineEdit(self.tab_4)
        self.lbl.setGeometry(690, 20, 141, 21)
        #TextEdit for your tasks
        self.TextEdit_search_2 = QtWidgets.QTextEdit(self.tab_4)
        self.TextEdit_search_2.setReadOnly(True)
        self.TextEdit_search_2.setGeometry(QtCore.QRect(380, 50, 450, 401))
        self.TextEdit_search_2.setObjectName("TextEdit_search")
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(10)
        self.TextEdit_search_2.setFont(font)
        #Groupbox_3 "Quick launch"
        self.groupBox_3 = QtWidgets.QGroupBox(self.tab_4)
        self.groupBox_3.setGeometry(QtCore.QRect(10, 40, 341, 181))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(10)
        self.groupBox_3.setFont(font)
        self.groupBox_3.setObjectName("groupBox_3")
        #button Outlook message
        self.btn_Outlook_message = QtWidgets.QPushButton(self.tab_4)
        self.btn_Outlook_message.setObjectName("Outlook message")
        self.btn_Outlook_message.setGeometry(30, 70, 131, 31)
        self.btn_Outlook_message.setText("Outlook message")
        self.btn_Outlook_message.setFont(
            QtGui.QFont("Arial", 10, QtGui.QFont.Bold))
        #button phone_list
        self.btn_phone_list = QtWidgets.QPushButton(self.tab_4)
        self.btn_phone_list.setGeometry(QtCore.QRect(190, 120, 131, 31))
        self.btn_phone_list.setObjectName("Your link")
        self.btn_phone_list.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Bold))
        #button db_learn
        self.btn_db_learn = QtWidgets.QPushButton(self.tab_4)
        self.btn_db_learn.setGeometry(QtCore.QRect(30, 170, 131, 31))
        self.btn_db_learn.setObjectName("DataBase Learn")
        self.btn_db_learn.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Bold))
        #button RDP
        self.btn_rdp = QtWidgets.QPushButton(self.tab_4)
        self.btn_rdp.setObjectName("run_rdp")
        self.btn_rdp.setGeometry(30, 120, 131, 31)
        self.btn_rdp.setText("run RDP")
        self.btn_rdp.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Bold))
        #button run programm
        self.btn_run_programm = QtWidgets.QPushButton(self.tab_4)
        self.btn_run_programm.setGeometry(QtCore.QRect(190, 70, 131, 31))
        self.btn_run_programm.setObjectName("Run programm")
        self.btn_run_programm.setFont(
            QtGui.QFont("Arial", 10, QtGui.QFont.Bold))

        self.retranslateUi(Form)
        self.tabWidget.setCurrentIndex(0)
        self.toolBox.setCurrentIndex(1)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        #btn feedback
        self.btn_feedback.setText(_translate("Form", "Feedback"))

        #Tab 1 "Cities"
        self.tabWidget.setTabText(self.tabWidget.indexOf(
            self.tab_2), _translate("Form", "Cities"))

        #Tab 2 "Work hours"
        self.toolBox.setItemText(self.toolBox.indexOf(
            self.page), _translate("Form", "Work hours    🠗"))
        self.toolBox.setItemText(self.toolBox.indexOf(
            self.page_2), _translate("Form", "Days off   🠗"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(
            self.tab_3), _translate("Form", "Work hours and holidays"))
        self.btn_graphic.setText(_translate("Form", "Show"))
        self.query_btn_2.setText(_translate("Form", "Show records"))

        #Tab 3 "Admin"
        self.f_name_label.setText(_translate("Form", "Name Administrator"))
        self.filial_label.setText(_translate("Form", "City"))
        self.data_start_label.setText(
            _translate("Form", "Date start vacation"))
        self.data_finish_label.setText(
            _translate("Form", "Date finish vacation"))
        self.submit_btn.setText(_translate("Form", "Add in DB"))
        self.query_btn.setText(_translate("Form", "Show record"))
        self.Label_query.setText(_translate("Form", ""))
        self.label_discription.setText(
            _translate("Form", "List admins in vacation"))
        self.btn_open_file.setText(_translate("Form", "Open file"))
        self.delete_btn.setText(_translate("Form", "Delete record"))
        self.label_5.setText(_translate("Form", "Select by ID"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(
            self.tab), _translate("Form", "Admin"))

        #Tab 4 "Control panel"
        self.tabWidget.setTabText(self.tabWidget.indexOf(
            self.tab_4), _translate("Form", "Control Panel"))
        self.TextEdit_query2.setText(_translate("Form", ""))
        self.label_discription2.setText(
            _translate("Form", "List admins in vacation"))
        self.btn_search.setText(_translate("Form", "Search"))
        #Groupbox "Quick launch"
        self.groupBox_3.setTitle(_translate("Form", "Quick launch"))
        self.Input_PC_label.setText(_translate("Form", "Name PC"))
        self.btn_run_programm.setText(_translate("Form", "run programm"))
        self.btn_db_learn.setText(_translate("Form", "DataBase Learn"))
        self.btn_phone_list.setText(_translate("Form", "Your link"))
        self.btn_rdp.setText(_translate("Form", "run RDP"))


if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    application = mywindow()
    application.show()
    sys.exit(app.exec())


