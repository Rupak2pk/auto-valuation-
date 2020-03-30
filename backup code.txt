#Python 3
#3/20/2020
#Rupak Kannan and Shyam Kannan
#Auto Valuation

import sys
import openpyxl
import csv
import glob
from xlsxwriter.workbook import Workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import shutil
import os
import os.path
import win32com.client
from openpyxl.utils import get_column_letter
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QMessageBox, QDialogButtonBox, QFileDialog

  
def write_to_target(target, directory, sheet_name):
    book = openpyxl.load_workbook(target)
    ws = book.get_sheet_by_name(sheet_name)
    f = open(directory)
    reader = csv.reader(f)
    for row_index, row in enumerate(reader):
        for column_index, cell in enumerate(row):
            column_letter = get_column_letter((column_index + 1))
            s = cell
            try:
                try:
                    s=float(s.replace(',', ''))
                except:
                    s=float(s)
            except ValueError:
                pass

            ws.cell('%s%s'%(column_letter, (row_index + 1))).value = s
    
    book.save(filename = target)

def write_to_target_xlsx(target, directory, sheet_name):
    wb1 = openpyxl.load_workbook(target)
    ws1 = wb1.get_sheet_by_name(sheet_name)
    
    wb2 = openpyxl.load_workbook(directory)
    ws2 = wb2.active
    mr = ws2.max_row
    mc = ws2.max_column
    
    for i in range(1, mr + 1):
        for j in range(1, mc + 1):
            c = ws2.cell(row = i, column = j)
            ws1.cell(row = i, column = j).value = c.value 
    wb1.save(filename = target)
            
class Ui_MainWindow(object):
    def setup_Ui(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(485, 580)
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
        
        self.company_ticker_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.company_ticker_txt.setGeometry(QtCore.QRect(180, 130, 61, 20))
        self.company_ticker_txt.setObjectName("company_ticker_txt")
        
        self.company_ticker_lbl = QtWidgets.QLabel(self.centralwidget)
        self.company_ticker_lbl.setGeometry(QtCore.QRect(40, 130, 131, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.company_ticker_lbl.setFont(font)
        self.company_ticker_lbl.setObjectName("company_ticket_lbl")
        
        self.company_ticker_btn = QtWidgets.QPushButton(self.centralwidget)
        self.company_ticker_btn.setGeometry(QtCore.QRect(235, 130, 125, 23))
        self.company_ticker_btn.setObjectName("company_ticker_btn")
        self.company_ticker_btn.clicked.connect(self.morningstar_download)
        
        self.balance_sheet_btn = QtWidgets.QPushButton(self.centralwidget)
        self.balance_sheet_btn.setGeometry(QtCore.QRect(380, 160, 75, 23))
        self.balance_sheet_btn.setObjectName("balance_sheet_btn")
        self.balance_sheet_btn.clicked.connect(self.get_xl_balance)
        
        self.balance_sheet_lbl = QtWidgets.QLabel(self.centralwidget)
        self.balance_sheet_lbl.setGeometry(QtCore.QRect(40, 160, 101, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.balance_sheet_lbl.setFont(font)
        self.balance_sheet_lbl.setObjectName("balance_sheet_lbl")
        
        self.balance_sheet_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.balance_sheet_txt.setGeometry(QtCore.QRect(180, 160, 201, 21))
        self.balance_sheet_txt.setObjectName("balance_sheet_txt")
        self.balance_sheet_txt.setReadOnly(True)      
        
        self.cash_flow_btn = QtWidgets.QPushButton(self.centralwidget)
        self.cash_flow_btn.setGeometry(QtCore.QRect(380, 190, 75, 23))
        self.cash_flow_btn.setObjectName("cash_flow_btn")
        self.cash_flow_btn.clicked.connect(self.get_xl_cash)
        
        self.cash_flow_lbl = QtWidgets.QLabel(self.centralwidget)
        self.cash_flow_lbl.setGeometry(QtCore.QRect(40, 190, 101, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.cash_flow_lbl.setFont(font)
        self.cash_flow_lbl.setObjectName("cash_flow_lbl")
        
        self.cash_flow_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.cash_flow_txt.setGeometry(QtCore.QRect(180, 190, 201, 21))
        self.cash_flow_txt.setObjectName("cash_flow_txt")
        self.cash_flow_txt.setReadOnly(True)
        
        self.Income_statement_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.Income_statement_txt.setGeometry(QtCore.QRect(180, 220, 201, 21))
        self.Income_statement_txt.setObjectName("Income_statement_txt")
        self.Income_statement_txt.setReadOnly(True)
        
        self.income_statement_btn = QtWidgets.QPushButton(self.centralwidget)
        self.income_statement_btn.setGeometry(QtCore.QRect(380, 220, 75, 23))
        self.income_statement_btn.setObjectName("income_statement_btn")
        self.income_statement_btn.clicked.connect(self.get_xl_income)
        
        self.income_statement_lbl = QtWidgets.QLabel(self.centralwidget)
        self.income_statement_lbl.setGeometry(QtCore.QRect(40, 220, 151, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.income_statement_lbl.setFont(font)
        self.income_statement_lbl.setObjectName("income_statement_lbl")
        
        self.debt_spreadsheet_lbl = QtWidgets.QLabel(self.centralwidget)
        self.debt_spreadsheet_lbl.setGeometry(QtCore.QRect(40, 280, 151, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.debt_spreadsheet_lbl.setFont(font)
        self.debt_spreadsheet_lbl.setObjectName("debt_spreadsheet_lbl")
        
        self.debt_spreadsheet_btn = QtWidgets.QPushButton(self.centralwidget)
        self.debt_spreadsheet_btn.setGeometry(QtCore.QRect(380, 280, 75, 23))
        self.debt_spreadsheet_btn.setObjectName("debt_spreadsheet_btn")
        self.debt_spreadsheet_btn.clicked.connect(self.get_xl_debt)
        
        self.debt_spreadsheet_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.debt_spreadsheet_txt.setGeometry(QtCore.QRect(180, 280, 201, 21))
        self.debt_spreadsheet_txt.setObjectName("debt_spreadsheet_txt")
        self.debt_spreadsheet_txt.setReadOnly(True)
        
        self.key_ratios_lbl = QtWidgets.QLabel(self.centralwidget)
        self.key_ratios_lbl.setGeometry(QtCore.QRect(40, 250, 101, 20))
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
        self.onlyNumbers = QtGui.QDoubleValidator()
        self.mrperp__txt.setValidator(self.onlyNumbers)     
        
        self.mrperp_lbl = QtWidgets.QLabel(self.centralwidget)
        self.mrperp_lbl.setGeometry(QtCore.QRect(40, 310, 101, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.mrperp_lbl.setFont(font)
        self.mrperp_lbl.setObjectName("mrperp_lbl")
        
        self.risk_free_rate_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.risk_free_rate_txt.setGeometry(QtCore.QRect(180, 340, 61, 20))
        self.risk_free_rate_txt.setText("")
        self.risk_free_rate_txt.setObjectName("risk_free_rate_txt")
        self.onlyNumbers = QtGui.QDoubleValidator()
        self.risk_free_rate_txt.setValidator(self.onlyNumbers)        
        
        self.risk_free_rate_lbl = QtWidgets.QLabel(self.centralwidget)
        self.risk_free_rate_lbl.setGeometry(QtCore.QRect(40, 340, 141, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.risk_free_rate_lbl.setFont(font)
        self.risk_free_rate_lbl.setObjectName("risk_free_rate_lbl")
        
        self.terminal_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.terminal_txt.setGeometry(QtCore.QRect(180, 370, 61, 20))
        self.terminal_txt.setText("")
        self.terminal_txt.setObjectName("terminal_txt")
        self.terminal_txt.setValidator(self.onlyNumbers)
        self.terminal_lbl = QtWidgets.QLabel(self.centralwidget)
        self.terminal_lbl.setGeometry(QtCore.QRect(10, 370, 170, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.terminal_lbl.setFont(font)
        self.terminal_lbl.setObjectName("terminal_lbl")
        
        self.year_growth_txt = QtWidgets.QSpinBox(self.centralwidget)
        self.year_growth_txt.setGeometry(QtCore.QRect(180, 400, 61, 20))
        self.year_growth_txt.setObjectName("risk_free_rate_txt")
        
        self.year_growth_lbl = QtWidgets.QLabel(self.centralwidget)
        self.year_growth_lbl.setGeometry(QtCore.QRect(10, 400, 170, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.year_growth_lbl.setFont(font)
        self.year_growth_lbl.setObjectName("risk_free_rate_lbl")
        self.year_growth_txt.setMinimum(0)
        self.year_growth_txt.setMaximum(10)
        
        self.growth_rate_lbl = QtWidgets.QLabel(self.centralwidget)
        self.growth_rate_lbl.setGeometry(QtCore.QRect(40, 430, 130, 20))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.growth_rate_lbl.setFont(font)
        self.growth_rate_lbl.setObjectName("growth_rate_lbl")
        
        self.smallest_of_etc_rd = QtWidgets.QRadioButton(self.centralwidget)
        self.smallest_of_etc_rd.setGeometry(QtCore.QRect(60, 460, 271, 17))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(11)
        self.smallest_of_etc_rd.setFont(font)
        self.smallest_of_etc_rd.setObjectName("smallest_of_etc_rd")
        
        self.custom_rd = QtWidgets.QRadioButton(self.centralwidget)
        self.custom_rd.setGeometry(QtCore.QRect(60, 490, 231, 17))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(11)
        self.custom_rd.setFont(font)
        self.custom_rd.setObjectName("custom_rd")
        
        self.custom_txt = QtWidgets.QLineEdit(self.centralwidget)
        self.custom_txt.setGeometry(QtCore.QRect(160, 490, 51, 20))
        self.custom_txt.setObjectName("custom_txt")
        
        self.run_btn = QtWidgets.QPushButton(self.centralwidget)
        self.run_btn.setGeometry(QtCore.QRect(50, 520, 75, 23))
        self.run_btn.setObjectName("run_btn")
        self.run_btn.clicked.connect(self.run)
        
        self.reset_btn = QtWidgets.QPushButton(self.centralwidget)
        self.reset_btn.setGeometry(QtCore.QRect(140, 520, 75, 23))
        self.reset_btn.setObjectName("reset_btn")
        self.reset_btn.clicked.connect(self.reset)
        
        self.close_btn = QtWidgets.QPushButton(self.centralwidget)
        self.close_btn.setGeometry(QtCore.QRect(320, 520, 75, 23))
        self.close_btn.setObjectName("close_btn")
        self.close_btn.clicked.connect(self.close)
        
        self.get_bond_data = QtWidgets.QPushButton(self.centralwidget)
        self.get_bond_data.setGeometry(QtCore.QRect(230, 520, 75, 23))
        #self.get_bond_data.clicked.connect(self.close)
        
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 580, 21))
        self.menubar.setObjectName("menubar")
        
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
    
    def morningstar_download(self):
        ticker = self.company_ticker_txt.text()
        location = os.getcwd()
        path = os.path.join(location, ticker)
        try:
            shutil.rmtree(path)
        except:
            pass
        
        try:
            options = webdriver.ChromeOptions()
            current_dir = os.getcwd()
            prefs = {
            "download.default_directory": current_dir + '\\' + ticker,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
            }            
            options.add_experimental_option('prefs', prefs)
            ticker = self.company_ticker_txt.text()   
            driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options = options)
            driver.set_window_size(777, 777)
            driver.get("http://financials.morningstar.com/balance-sheet/bs.html?t="+ticker+"&region=usa&culture=en-US")
            driver.execute_script("javascript:SRT_stocFund.orderControl('desc','Descending')")
            driver.execute_script("javascript:SRT_stocFund.Export();")
            driver.get("http://financials.morningstar.com/income-statement/is.html?t="+ticker+"&region=usa&culture=en-US")
            driver.execute_script("javascript:SRT_stocFund.orderControl('desc','Descending')")
            driver.execute_script("javascript:SRT_stocFund.Export();")
            driver.get("http://financials.morningstar.com/cash-flow/cf.html?t="+ticker+"&region=usa&culture=en-US")
            driver.execute_script("javascript:SRT_stocFund.orderControl('desc','Descending')")
            driver.execute_script("javascript:SRT_stocFund.Export();")
            driver.get("http://financials.morningstar.com/ratios/r.html?t="+ticker+"&region=usa&culture=en-US")
            driver.execute_script("javascript:orderChange('desc','Descending');")
            driver.execute_script("javascript:exportKeyStat2CSV();")
            driver.get("http://www.google.com")
            driver.quit()
            
            #Adds files to the txt statement
            
            self.Income_statement_txt.setText(os.path.realpath(ticker + "\\{} Income Statement.csv").format(ticker) )            
            self.balance_sheet_txt.setText(os.path.realpath(ticker + "\\{} Balance Sheet.csv").format(ticker))           
            self.cash_flow_txt.setText(os.path.realpath(ticker + "\\{} Cash Flow.csv").format(ticker))            
            self.key_ratios_txt.setText(os.path.realpath(ticker + "\\{} Key Ratios.csv").format(ticker))     
             
               
        except:
            try:
                driver.quit()
            except:
                pass
            msg = QMessageBox()
            msg.setWindowTitle("Notice")
            msg.setIcon(QMessageBox.Information)
            msg.setText("An error has occured, there is either: \n* An nonexistent company ticker \n* Error on Morningstar's website  \n* Interruption with the download process \nPlease try again.")
            notice = msg.exec()
    
    def get_xl_income(self):
        global book_income
        income_filename, filter = QtWidgets.QFileDialog.getOpenFileName(caption='Open file',  filter='CSV (*.CSV)')

        if income_filename:
            self.Income_statement_txt.setText(income_filename)
        
    def get_xl_balance(self):
        global book_balance
        balance_filename, filter = QtWidgets.QFileDialog.getOpenFileName(caption='Open file',  filter='CSV (*.CSV)')

        if balance_filename:
            self.balance_sheet_txt.setText(balance_filename)         
        
    def get_xl_cash(self):
        global book_cash
        cash_filename, filter = QtWidgets.QFileDialog.getOpenFileName(caption='Open file',  filter='CSV (*.CSV)')

        if cash_filename:
            self.cash_flow_txt.setText(cash_filename)
        
    def get_xl_debt(self):
        global book_debt
        debt_filename, filter = QtWidgets.QFileDialog.getOpenFileName(caption='Open file',  filter='CSV (*.CSV);;xlsx (*.xlsx)')

        if debt_filename:
            self.debt_spreadsheet_txt.setText(debt_filename)             
            
    
    def get_xl_ratios(self):
        global book_ratios
        ratios_filename, filter = QtWidgets.QFileDialog.getOpenFileName(caption='Open file',  filter='CSV (*.CSV)')

        if ratios_filename:
            self.key_ratios_txt.setText(ratios_filename)  
           
    def run(self):
        '''if self.Income_statement_txt.text() == "" or  self.balance_sheet_txt.text() == "" or self.cash_flow_txt.text() == "" or self.debt_spreadsheet_txt.text() == "" or self.key_ratios_txt.text() == "" or self.company_ticker_txt.text() == "" or self.mrperp__txt.text() == "" or self.risk_free_rate_txt.text() == "":'''
        if self.Income_statement_txt.text() == "" or  self.balance_sheet_txt.text() == "" or self.cash_flow_txt.text() == "" or self.debt_spreadsheet_txt.text() == "" or self.key_ratios_txt.text() == "":
            msg = QMessageBox()
            msg.setWindowTitle("Notice")
            msg.setIcon(QMessageBox.Information)
            msg.setText("All information must be filled")
            notice = msg.exec()        
        
        elif self.custom_rd.isChecked() and self.custom_txt.text() == "":
            msg = QMessageBox()
            msg.setWindowTitle("Notice")
            msg.setIcon(QMessageBox.Information)
            msg.setText("All information must be filled")
            notice = msg.exec()            

        else:
            try:
                ticker = self.company_ticker_txt.text()
                original = 'Template.xlsx'
                target = ticker + '_Valuation.xlsx'
                shutil.copyfile(original, target)
                
                book = openpyxl.load_workbook(target)
                sheet_name = 'Income Statement'
                
                ws = book.get_sheet_by_name('DDM')
                
                terminal_decimal = float(self.terminal_txt.text()) / 100
                ws['B7'].value = float(terminal_decimal)
                ws['B7'].number_format = '0.00%'
                
                risk_free_decimal = float(self.risk_free_rate_txt.text()) / 100
                ws['F4'].value = float(risk_free_decimal)
                ws['F4'].number_format = '0.00%'                
                
                ws['B5'].value = int(self.year_growth_txt.value())
                
                ws = book.get_sheet_by_name('DCF')
                
                MRP_decimal = float(self.mrperp__txt.text()) / 100
                ws['P9'].value = float(MRP_decimal)
                ws['P9'].number_format = '0.00%'
                
                book.save(filename = target)
                
                write_to_target(target, self.Income_statement_txt.text(), 'Income Statement')
                write_to_target(target, self.balance_sheet_txt.text(), 'Balance Sheet (Annual)')
                write_to_target(target, self.cash_flow_txt.text(), 'Cash Flow Statement')
                write_to_target(target, self.key_ratios_txt.text(), 'Key Ratios')
                write_to_target_xlsx(target, self.debt_spreadsheet_txt.text(), 'Debt Template')           
                
                msg = QMessageBox()
                msg.setWindowTitle("Notice")
                msg.setIcon(QMessageBox.Information)
                msg.setText("Workbook Created!")
                notice = msg.exec()
                os.system("start EXCEL.EXE " + target)
                
            except:
                msg = QMessageBox()
                msg.setWindowTitle("Notice")
                msg.setIcon(QMessageBox.Information)
                msg.setText("Something went wrong. Try redownloading the sheets form MorningStar or check the values that have been entered or check if an valuation excel is currently open and close it.")
                notice = msg.exec()           
                return error
            
    def reset(self):
        self.Income_statement_txt.setText("")
        self.balance_sheet_txt.setText("")
        self.cash_flow_txt.setText("")
        self.debt_spreadsheet_txt.setText("")
        self.key_ratios_txt.setText("")
        self.company_ticker_txt.setText("")
        self.mrperp__txt.setText("")
        self.risk_free_rate_txt.setText("")
        self.terminal_txt.setText("")
        self.year_growth_txt.setValue(0)
        self.custom_txt.setText("")
        
    def close(self):
        app.quit()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Autovaluer"))
        self.title.setText(_translate("MainWindow", "Autovaluer"))
        self.income_statement_btn.setText(_translate("MainWindow", "Choose File"))
        self.income_statement_lbl.setText(_translate("MainWindow", "Income Statment"))
        self.balance_sheet_btn.setText(_translate("MainWindow", "Choose File"))
        self.balance_sheet_lbl.setText(_translate("MainWindow", "Balance Sheet"))
        self.cash_flow_btn.setText(_translate("MainWindow", "Choose File"))
        self.cash_flow_lbl.setText(_translate("MainWindow", "Cash Flow"))
        self.company_ticker_lbl.setText(_translate("MainWindow", "Company Ticker"))
        self.company_ticker_btn.setText(_translate("MainWindow", "Get Finacial Data"))
        self.debt_spreadsheet_lbl.setText(_translate("MainWindow", "Debt Spreadsheet"))
        self.debt_spreadsheet_btn.setText(_translate("MainWindow", "Choose File"))
        self.key_ratios_lbl.setText(_translate("MainWindow", "Key Ratios"))
        self.key_ratios_btn.setText(_translate("MainWindow", "Choose File"))
        self.mrperp_lbl.setText(_translate("MainWindow", "MRP/ERP"))
        self.risk_free_rate_lbl.setText(_translate("MainWindow", "Risk Free Rate"))
        self.terminal_lbl.setText(_translate("MainWindow", "Terminal Growth Rate"))
        self.year_growth_lbl.setText(_translate("MainWindow", "Years of High Growth"))
        self.growth_rate_lbl.setText(_translate("MainWindow", "Growth Rate:"))
        self.smallest_of_etc_rd.setText(_translate("MainWindow", "Use smallest of IAR, SAR, or ROLC"))
        self.custom_rd.setText(_translate("MainWindow", "Custom"))
        self.run_btn.setText(_translate("MainWindow", "Run"))
        self.reset_btn.setText(_translate("MainWindow", "Reset"))
        self.close_btn.setText(_translate("MainWindow", "Close"))
        self.get_bond_data.setText(_translate("MainWindow", "Get Bond Data"))
        
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Windows')
    Gui = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setup_Ui(Gui)
    Gui.show()
    msg = QMessageBox()
    msg.setWindowTitle("Disclaimer")
    msg.setIcon(QMessageBox.Information)
    msg.setText("The morningstar website may or may not be unstabled. There have been reports of JP Morgan (JPM) excel sheets being installed at random. Proceed with caution.")
    notice = msg.exec()    
    sys.exit(app.exec_())