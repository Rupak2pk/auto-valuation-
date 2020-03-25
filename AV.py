#Python 3
#3/20/2020
#Rupak Kannan and Shyam Kannan
#Auto Valuation

import sys
import openpyxl
from openpyxl import *
import shutil
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QMessageBox, QDialogButtonBox, QFileDialog

book = openpyxl.load_workbook('Template.xlsx')
sheet = book.active

def copy_template():
    original = 'Template.xlsx'
    target = 'S_info.xlsx'
    
    shutil.copyfile(original, target)

copy_template()
class Ui_MainWindow(object):
    def setup_Ui(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(485, 545)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        
        self.title = QtWidgets.QLabel(self.centralwidget)
        self.title.setGeometry(QtCore.QRect(190, 40, 201, 51))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(16)
        self.title.setFont(font)
        self.title.setObjectName("title")
        
        self.logo = QtWidgets.QLabel(self.centralwidget)
        self.logo.setGeometry(QtCore.QRect(40, 10, 161, 111))
        self.logo.setText("")
        self.logo.setPixmap(QtGui.QPixmap("Capture.PNG"))
        self.logo.setObjectName("logo")
        
        self.Income_statement_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.Income_statement_txt.setGeometry(QtCore.QRect(180, 130, 201, 21))
        self.Income_statement_txt.setObjectName("Income_statement_txt")
        self.Income_statement_txt.setReadOnly(True)
        
        self.income_statement_btn = QtWidgets.QPushButton(self.centralwidget)
        self.income_statement_btn.setGeometry(QtCore.QRect(380, 130, 75, 23))
        self.income_statement_btn.setObjectName("income_statement_btn")
        self.income_statement_btn.clicked.connect(self.get_xl_income)
        
        self.income_statement_lbl = QtWidgets.QLabel(self.centralwidget)
        self.income_statement_lbl.setGeometry(QtCore.QRect(60, 130, 101, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.income_statement_lbl.setFont(font)
        self.income_statement_lbl.setObjectName("income_statement_lbl")
        
        self.balance_sheet_btn = QtWidgets.QPushButton(self.centralwidget)
        self.balance_sheet_btn.setGeometry(QtCore.QRect(380, 160, 75, 23))
        self.balance_sheet_btn.setObjectName("balance_sheet_btn")
        self.balance_sheet_btn.clicked.connect(self.get_xl_balance)
        
        self.balance_sheet_lbl = QtWidgets.QLabel(self.centralwidget)
        self.balance_sheet_lbl.setGeometry(QtCore.QRect(80, 160, 101, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.balance_sheet_lbl.setFont(font)
        self.balance_sheet_lbl.setObjectName("balance_sheet_lbl")
        
        self.balance_sheet_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.balance_sheet_txt.setGeometry(QtCore.QRect(180, 160, 201, 21))
        self.balance_sheet_txt.setObjectName("balance_sheet_txt")
        self.balance_sheet_txt.setReadOnly(True)      
        
        self.cash_file_btn = QtWidgets.QPushButton(self.centralwidget)
        self.cash_file_btn.setGeometry(QtCore.QRect(380, 190, 75, 23))
        self.cash_file_btn.setObjectName("cash_file_btn")
        self.cash_file_btn.clicked.connect(self.get_xl_cash)
        
        self.cash_file_lbl = QtWidgets.QLabel(self.centralwidget)
        self.cash_file_lbl.setGeometry(QtCore.QRect(100, 190, 101, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.cash_file_lbl.setFont(font)
        self.cash_file_lbl.setObjectName("cash_file_lbl")
        
        self.cash_file_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.cash_file_txt.setGeometry(QtCore.QRect(180, 190, 201, 21))
        self.cash_file_txt.setObjectName("cash_file_txt")
        self.cash_file_txt.setReadOnly(True)
        
        self.company_ticket_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.company_ticket_txt.setGeometry(QtCore.QRect(180, 280, 61, 20))
        self.company_ticket_txt.setObjectName("company_ticket_txt")
        
        self.company_ticket_lbl = QtWidgets.QLabel(self.centralwidget)
        self.company_ticket_lbl.setGeometry(QtCore.QRect(60, 280, 101, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.company_ticket_lbl.setFont(font)
        self.company_ticket_lbl.setObjectName("company_ticket_lbl")
        
        self.debt_spreadsheet_lbl = QtWidgets.QLabel(self.centralwidget)
        self.debt_spreadsheet_lbl.setGeometry(QtCore.QRect(60, 220, 101, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.debt_spreadsheet_lbl.setFont(font)
        self.debt_spreadsheet_lbl.setObjectName("debt_spreadsheet_lbl")
        
        self.debt_spreadsheet_btn = QtWidgets.QPushButton(self.centralwidget)
        self.debt_spreadsheet_btn.setGeometry(QtCore.QRect(380, 220, 75, 23))
        self.debt_spreadsheet_btn.setObjectName("debt_spreadsheet_btn")
        self.debt_spreadsheet_btn.clicked.connect(self.get_xl_debt)
        
        self.debt_spreadsheet_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.debt_spreadsheet_txt.setGeometry(QtCore.QRect(180, 220, 201, 21))
        self.debt_spreadsheet_txt.setObjectName("debt_spreadsheet_txt")
        self.debt_spreadsheet_txt.setReadOnly(True)
        
        self.key_ratios_lbl = QtWidgets.QLabel(self.centralwidget)
        self.key_ratios_lbl.setGeometry(QtCore.QRect(90, 250, 101, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.key_ratios_lbl.setFont(font)
        self.key_ratios_lbl.setObjectName("key_ratios_lbl")
        
        self.key_ratios_btn = QtWidgets.QPushButton(self.centralwidget)
        self.key_ratios_btn.setGeometry(QtCore.QRect(380, 250, 75, 23)) 
        self.key_ratios_btn.clicked.connect(self.get_xl_ratios)
        
        self.key_ratios_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.key_ratios_txt.setGeometry(QtCore.QRect(180, 250, 201, 21))
        self.key_ratios_txt.setReadOnly(True)  
        
        
        self.mrperp__txt = QtWidgets.QLineEdit(self.centralwidget)
        self.mrperp__txt.setGeometry(QtCore.QRect(180, 310, 61, 20))
        self.mrperp__txt.setText("")
        self.mrperp__txt.setObjectName("mrperp__txt")
        
        self.mrperp_lbl = QtWidgets.QLabel(self.centralwidget)
        self.mrperp_lbl.setGeometry(QtCore.QRect(100, 310, 61, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.mrperp_lbl.setFont(font)
        self.mrperp_lbl.setObjectName("mrperp_lbl")
        
        self.risk_free_rate_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.risk_free_rate_txt.setGeometry(QtCore.QRect(180, 340, 61, 20))
        self.risk_free_rate_txt.setText("")
        self.risk_free_rate_txt.setObjectName("risk_free_rate_txt")
        
        self.risk_free_rate_lbl = QtWidgets.QLabel(self.centralwidget)
        self.risk_free_rate_lbl.setGeometry(QtCore.QRect(80, 340, 81, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.risk_free_rate_lbl.setFont(font)
        self.risk_free_rate_lbl.setObjectName("risk_free_rate_lbl")
        
        self.growth_rate_lbl = QtWidgets.QLabel(self.centralwidget)
        self.growth_rate_lbl.setGeometry(QtCore.QRect(30, 370, 81, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.growth_rate_lbl.setFont(font)
        self.growth_rate_lbl.setObjectName("growth_rate_lbl")
        
        self.smallest_of_etc_rd = QtWidgets.QRadioButton(self.centralwidget)
        self.smallest_of_etc_rd.setGeometry(QtCore.QRect(60, 400, 231, 17))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(11)
        self.smallest_of_etc_rd.setFont(font)
        self.smallest_of_etc_rd.setObjectName("smallest_of_etc_rd")
        
        self.custom_rd = QtWidgets.QRadioButton(self.centralwidget)
        self.custom_rd.setGeometry(QtCore.QRect(60, 430, 231, 17))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(11)
        self.custom_rd.setFont(font)
        self.custom_rd.setObjectName("custom_rd")
        
        self.custom_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.custom_txt.setGeometry(QtCore.QRect(140, 430, 51, 20))
        self.custom_txt.setObjectName("custom_txt")
        
        self.run_btn = QtWidgets.QPushButton(self.centralwidget)
        self.run_btn.setGeometry(QtCore.QRect(50, 470, 75, 23))
        self.run_btn.setObjectName("run_btn")
        self.run_btn.clicked.connect(self.run)
        
        self.reset_btn = QtWidgets.QPushButton(self.centralwidget)
        self.reset_btn.setGeometry(QtCore.QRect(140, 470, 75, 23))
        self.reset_btn.setObjectName("reset_btn")
        self.reset_btn.clicked.connect(self.reset)
        
        self.cancel_btn = QtWidgets.QPushButton(self.centralwidget)
        self.cancel_btn.setGeometry(QtCore.QRect(320, 470, 75, 23))
        self.cancel_btn.setObjectName("cancel_btn")
        self.cancel_btn.clicked.connect(self.cancel)
        
        self.get_bond_data = QtWidgets.QPushButton(self.centralwidget)
        self.get_bond_data.setGeometry(QtCore.QRect(230, 470, 75, 23))
        #self.get_bond_data.clicked.connect(self.cancel)
        
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 485, 21))
        self.menubar.setObjectName("menubar")
        
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
    
    def get_xl_income(self):
        filename, filter = QtWidgets.QFileDialog.getOpenFileName(caption='Open file',  filter='xlsx (*.xlsx);;CSV (*.CSV)')

        if filename:
            self.Income_statement_txt.setText(filename)
        
    def get_xl_balance(self):
        filename, filter = QtWidgets.QFileDialog.getOpenFileName(caption='Open file',  filter='xlsx (*.xlsx);;CSV (*.CSV)')

        if filename:
            self.balance_sheet_txt.setText(filename)
        
    def get_xl_cash(self):
        filename, filter = QtWidgets.QFileDialog.getOpenFileName(caption='Open file',  filter='xlsx (*.xlsx);;CSV (*.CSV)')

        if filename:
            self.cash_file_txt.setText(filename)
        
    def get_xl_debt(self):
        filename, filter = QtWidgets.QFileDialog.getOpenFileName(caption='Open file',  filter='xlsx (*.xlsx);;CSV (*.CSV)')

        if filename:
            self.debt_spreadsheet_txt.setText(filename)   
    
    def get_xl_ratios(self):
        filename, filter = QtWidgets.QFileDialog.getOpenFileName(caption='Open file',  filter='xlsx (*.xlsx);;CSV (*.CSV)')

        if filename:
            self.key_ratios_txt.setText(filename)   
    
    def run(self):
        if self.Income_statement_txt.text() == "" or  self.balance_sheet_txt.text() == "" or self.cash_file_txt.text() == "" or self.debt_spreadsheet_txt.text() == "" or self.key_ratios_txt.text() == "" or self.company_ticket_txt.text() == "" or self.mrperp__txt.text() == "" or self.risk_free_rate_txt.text() == "" or self.custom_txt.text() == "":
            msg = QMessageBox()
            msg.setWindowTitle("Notice")
            msg.setIcon(QMessageBox.Information)
            msg.setText("All information must be filled")
            notice = msg.exec()        
            
            
    def reset(self):
        self.Income_statement_txt.setText("")
        self.balance_sheet_txt.setText("")
        self.cash_file_txt.setText("")
        self.debt_spreadsheet_txt.setText("")
        self.key_ratios_txt.setText("")
        self.company_ticket_txt.setText("")
        self.mrperp__txt.setText("")
        self.risk_free_rate_txt.setText("")
        self.custom_txt.setText("")
        
    def cancel(self):
        app.quit()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Autovaluer"))
        self.title.setText(_translate("MainWindow", "Autovaluer"))
        self.income_statement_btn.setText(_translate("MainWindow", "Choose File"))
        self.income_statement_lbl.setText(_translate("MainWindow", "Income Statment"))
        self.balance_sheet_btn.setText(_translate("MainWindow", "Choose File"))
        self.balance_sheet_lbl.setText(_translate("MainWindow", "Balance Sheet"))
        self.cash_file_btn.setText(_translate("MainWindow", "Choose File"))
        self.cash_file_lbl.setText(_translate("MainWindow", "Cash File"))
        self.company_ticket_lbl.setText(_translate("MainWindow", "Company Ticker"))
        self.debt_spreadsheet_lbl.setText(_translate("MainWindow", "Debt Spreadsheet"))
        self.debt_spreadsheet_btn.setText(_translate("MainWindow", "Choose File"))
        self.key_ratios_lbl.setText(_translate("MainWindow", "Key Ratios"))
        self.key_ratios_btn.setText(_translate("MainWindow", "Choose File"))
        self.mrperp_lbl.setText(_translate("MainWindow", "MRP/ERP"))
        self.risk_free_rate_lbl.setText(_translate("MainWindow", "Risk Free Rate"))
        self.growth_rate_lbl.setText(_translate("MainWindow", "Growth Rate:"))
        self.smallest_of_etc_rd.setText(_translate("MainWindow", "Use smallest of IAR, SAR, or ROLC"))
        self.custom_rd.setText(_translate("MainWindow", "Custom"))
        self.run_btn.setText(_translate("MainWindow", "Run"))
        self.reset_btn.setText(_translate("MainWindow", "Reset"))
        self.cancel_btn.setText(_translate("MainWindow", "Cancel"))
        self.get_bond_data.setText(_translate("MainWindow", "Get Bond Data"))
        
class Controller:
    def open_main_window(self):
        self.Gui = QtWidgets.QMainWindow()
        self.ui = Ui_MainWindow()
        self.ui.setup_Ui(self.Gui)
        self.Gui.show()  
        
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Windows')
    Controller = Controller()
    Controller.open_main_window()
    sys.exit(app.exec_())
