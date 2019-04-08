import pyodbc
import sys
import datetime
import openpyxl
from openpyxl.styles.borders import Side
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from PyQt5.QtCore import Qt, QSize, pyqtSlot, QDate, QTime, QEvent
from PyQt5.QtGui import QIcon, QPixmap, QFont, QPalette, QBrush, QColor
from PyQt5.QtWidgets import (QDialog, QApplication, QLineEdit, QPushButton, QWidget, QMainWindow, QCheckBox, QMenu, QItemDelegate,
                             QTableWidget, QMessageBox,
                             QTableWidgetItem, QLabel, QListWidget, QListWidgetItem, QPlainTextEdit, QFileDialog,
                             QComboBox, QCalendarWidget, QAbstractItemView, QDateEdit,
                             QProgressBar, QGroupBox, QGridLayout, QFrame, QTextEdit, QTimeEdit)

class platinumDatabase():
    def __init__(self, server='127.0.0.1', database='Platinum', username='sa', password='plat1007'):
        self.server = server
        self.database = database
        self.username = username
        self.password = password
        self.createCursor()
        # shannon was here
        # Again

    def createCursor(self):
        try:
            self.conn = pyodbc.connect(
                'Driver={};Server={};Database={};UID={};PWD={}'.format('SQL Server', self.server, self.database, self.username, self.password))
            self.cursor = self.conn.cursor()
            return '1'
        except:
            return '0'
    # SQl Date corrector Add time
    def script(self, startDateTIme="2000-01-01 00:00:00", endDateTime="2000-01-01 00:00:00", hoursToAdd="00:00:00"):
        dates = self.cursor.execute("SELECT Line_No, Date_Time FROM Sales_Journal WHERE Date_Time >= ? AND Date_Time <= ?", (startDateTIme, endDateTime))
        dates = list(dates)
        for date in dates:
            tHk = hoursToAdd.split(":")
            newDate = date[1] + datetime.timedelta(hours=int(tHk[0]), minutes=int(tHk[1]), seconds=int(tHk[2]))
            newDate = newDate.strftime("%Y-%m-%d %H:%M:%S")
            print(date[0])
            print(newDate)
            self.cursor.execute("UPDATE Sales_Journal SET Date_Time = ? WHERE Line_No = ?" , (newDate, date[0]))
            self.conn.commit()

    # Sql Date Corrector subtract time
    def script2(self, startDateTIme="2000-01-01 00:00:00", endDateTime="2000-01-01 00:00:00", hoursToAdd="00:00:00"):
        dates = self.cursor.execute("SELECT Line_No, Date_Time FROM Sales_Journal WHERE Date_Time >= ? AND Date_Time <= ?", (startDateTIme, endDateTime))
        dates = list(dates)
        for date in dates:
            tHk = hoursToAdd.split(":")
            newDate = date[1] - datetime.timedelta(hours=int(tHk[0]), minutes=int(tHk[1]), seconds=int(tHk[2]))
            newDate = newDate.strftime("%Y-%m-%d %H:%M:%S")
            print(date[0])
            print(newDate)
            self.cursor.execute("UPDATE Sales_Journal SET Date_Time = ? WHERE Line_No = ?" , (newDate, date[0]))
            self.conn.commit()




class PlatinumCor(QWidget):
    def __init__(self):
        super().__init__()
        self.__initUI__()

    # Function Sets up the main screen
    def __initUI__(self):
        self.productList = {}
        try:
            self.database = platinumDatabase()
            products = self.database.cursor.execute("SELECT Product_Code, Description FROM Products")
            for i in products:
                self.productList[i[0]] = i[1]
            print(self.productList)
        except:
            print("Failed To build productList")

        self.mainMenueScreen = QFrame()
        self.mainMenueScreen.setWindowTitle("Platinum Tool Kit")
        self.mainMenueScreen.setWindowIcon(QIcon("C:\Platinum\Icons\icon_p.ico"))
        self.mainMenueScreen.resize(600, 300)

        addRemoveTimeInvButton = QPushButton("Alter Sales Time")
        addRemoveTimeInvButton.setIcon(QIcon("C:\\Platinum\\Icons\\15minblue.ico"))
        addRemoveTimeInvButton.clicked.connect(self.addRemoveInvoiceTIme)
        # addRemoveTimeInvButton.setDisabled(True)

        alterInvoiceBUtton = QPushButton("Alter Invoices")
        alterInvoiceBUtton.setIcon(QIcon("C:\\Platinum\\Icons\\15min.ico"))
        alterInvoiceBUtton.clicked.connect(self.alterInvoice)
        # alterInvoiceBUtton.setDisabled(True)

        alterCashupsButton = QPushButton("Alter Cashups")
        alterCashupsButton.setIcon(QIcon("C:\Platinum\Icons\\reservations.ico"))
        alterCashupsButton.clicked.connect(self.alterInvoice)
        # alterCashupsButton.setDisabled(True)

        mainGrid = QGridLayout()
        mainGrid.addWidget(addRemoveTimeInvButton, 0, 0)
        mainGrid.addWidget(alterInvoiceBUtton, 0, 1)
        mainGrid.addWidget(alterCashupsButton, 1, 0)
        self.mainMenueScreen.setLayout(mainGrid)
        self.mainMenueScreen.show()

    def alterInvoice(self):
        self.alterInvoceScreen = QFrame()
        self.alterInvoceScreen.setWindowTitle("Alter Invoice")
        self.alterInvoceScreen.setWindowIcon(QIcon("C:\\Platinum\\Icons\\15min.ico"))
        self.alterInvoceScreen.resize(700, 500)

        invoiceNUmber = QLineEdit()
        searchButton = QPushButton("Search")
        searchButton.setIcon(QIcon("C:\\Platinum\\Icons\\Preview.ico"))
        searchButton.clicked.connect(lambda : self.searchInvoice(invoiceNUmber.text()))

        self.alterInvoiceTable = QTableWidget()


        updateButton = QPushButton("Update Changes")
        updateButton.setIcon(QIcon("C:\\Platinum\\Icons\\STH.ico"))
        updateButton.clicked.connect(self.updateSalesJournal)

        maingrid = QGridLayout()
        maingrid.addWidget(QLabel("Invoice Number"), 0, 0)
        maingrid.addWidget(invoiceNUmber, 0, 1)
        maingrid.addWidget(searchButton, 0, 2)
        maingrid.addWidget(self.alterInvoiceTable, 1, 0, 1, 3)
        maingrid.addWidget(updateButton, 2, 0, 1, 3)

        self.alterInvoceScreen.setLayout(maingrid)
        self.alterInvoceScreen.show()

    def recalculateCashupTotals(self, cashup="0"):
        self.confirmation = QMessageBox
        line = cashup
        choice = self.confirmation.question(self, 'Confirmation', "Sales Journal has been updated \n Recalculate Cashup %s ?" % line,
                                            self.confirmation.Yes | self.confirmation.No)
        if choice == self.confirmation.Yes:
            taxablesales = 0.00
            tax = 0.00
            cashTotal = 0.00 # Cash_Sales_Value
            cardTotal = 0.00 # Card_Sales_Value
            chargeTotal = 0.00 # Charge_Sales_Value
            chequeTotal = 0.00 # Cheque_Sales_Value

            # Calculate Cash Total
            try:
                cashData = self.database.cursor.execute("""SELECT Sales_Tax, Line_Total FROM Sales_Journal WHERE Function_Key =? AND Cashup_No=?""", (9.0, int(line)))
                for cash in cashData:
                    cashTotal += cash[1]
                    tax += cash[0]
                    if cash[0] > 0:
                        taxablesales += cash[1]
            except:
                print("Failed TO sun cash")
            # Calculate Card Total
            try:
                cardData = self.database.cursor.execute("""SELECT Sales_Tax, Line_Total FROM Sales_Journal WHERE Function_Key =? AND Cashup_No=?""", (10.0, int(line)))
                for card in cardData:
                    cardTotal += card[1]
                    tax += card[0]
                    if card[0] > 0:
                        taxablesales += card[1]
            except:
                print("Failed TO sum cards")
            # Calculate Charge TOtals
            try:
                chargeData = self.database.cursor.execute("""SELECT Sales_Tax, Line_Total FROM Sales_Journal WHERE Function_Key =? AND Cashup_No=?""", (12.0, int(line)))
                for charge in chargeData:
                    chargeTotal += charge[1]
                    tax += charge[0]
                    if charge[0] > 0:
                        taxablesales += charge[1]
            except:
                print("Failed TO Sum Charge")
            # Calculate Cheque Total
            try:
                checkData = self.database.cursor.execute(
                    """SELECT Sales_Tax, Line_Total FROM Sales_Journal WHERE Function_Key =? AND Cashup_No=?""", (11.0, int(line)))
                for check in checkData:
                    chequeTotal += check[1]
                    tax += check[0]
                    if check[0] > 0:
                        taxablesales += check[1]
            except:
                print("Failed TO Sum Charge")

            # Update Counters Table
            try:
                self.database.cursor.execute("""UPDATE Counters SET Cash_Sales_Value=?, Card_Sales_Value=?, Charge_Sales_Value=?, Cheque_Sales_Value=?, TaxableSales_Value=? WHERE Cashup_No=?""", (cashTotal, cardTotal, chargeTotal, chequeTotal, taxablesales, int(line)))
                self.database.conn.commit()
            except:
                pass
            print(cashTotal, cardTotal, chargeTotal, chequeTotal, taxablesales)



    def updateSalesJournal(self):
        for i in range(self.alterInvoiceTable.rowCount()):
            functionKey = float(self.alterInvoiceTable.item(i, 6).text())
            lineNo = int(self.alterInvoiceTable.item(i, 0).text())
            if functionKey == 7.0:
                producCode = self.alterInvoiceTable.item(i, 1).text()
                qty = float(self.alterInvoiceTable.item(i, 3).text())
                tax = float(self.alterInvoiceTable.item(i, 4).text())
                linetotal = float(self.alterInvoiceTable.item(i, 5).text())
                self.database.cursor.execute("""UPDATE Sales_Journal SET Product_Code=?, Qty=?, Sales_Tax=?, Line_Total=? WHERE Line_No =?""", (producCode, qty, tax, linetotal, lineNo))
                self.database.conn.commit()
            elif functionKey == 9 or functionKey == 10 or functionKey == 11 or functionKey == 12:
                tax = float(self.alterInvoiceTable.item(i, 4).text())
                linetotal = float(self.alterInvoiceTable.item(i, 5).text())
                self.database.cursor.execute(
                    """UPDATE Sales_Journal SET Sales_Tax=?, Function_Key=?, Line_Total=? WHERE Line_No =?""",
                    (tax, functionKey, linetotal, lineNo))
                self.database.conn.commit()
        self.recalculateCashupTotals(self.alterInvoiceTable.item(i, 7).text())


    def searchInvoice(self, invoiceNumber):
        self.alterInvoiceTable.clear()
        horizontalLable = ["Line_No", "Product Code", "Description", "Qty", "Tax", "Line Total", "Function Key", "Cash Up No"]
        self.alterInvoiceTable.setColumnCount(len(horizontalLable))
        self.alterInvoiceTable.setHorizontalHeaderLabels(horizontalLable)
        try:
            invoiceNumber = int(invoiceNumber)
            slip = list(self.database.cursor.execute("SELECT Line_No, Product_Code, Qty, Sales_Tax, Line_Total, Function_Key, Cashup_No FROM Sales_Journal WHERE Invoice_No = ?", (invoiceNumber)))
            self.alterInvoiceTable.setRowCount(len(slip))
            row = 0
            for item in slip:
                lineNo = item[0]
                productCode = item[1]
                functionKey = item[5]
                cashupNO = item[6]
                if functionKey != 14:
                    try:
                        if functionKey == 9:
                            productName = "CASH SALE"
                        elif functionKey == 10:
                            productName = "CARD SALE"
                        elif functionKey == 11:
                            productName = "CHEQ SALE"
                        elif functionKey == 12:
                            productName = "CHARGE SALE"
                        else:
                            productName = self.productList[productCode]

                    except:
                        productName = " "
                    qty = item[2]
                    salesTax = item[3]
                    lineTotal = item[4]

                    self.alterInvoiceTable.setItem(row, 0, QTableWidgetItem(str(lineNo)))
                    self.alterInvoiceTable.setItem(row, 1, QTableWidgetItem(str(productCode)))
                    self.alterInvoiceTable.setItem(row, 2, QTableWidgetItem(str(productName)))
                    self.alterInvoiceTable.setItem(row, 3, QTableWidgetItem(str(qty)))
                    self.alterInvoiceTable.setItem(row, 4, QTableWidgetItem(str(salesTax)))
                    self.alterInvoiceTable.setItem(row, 5, QTableWidgetItem(str(lineTotal)))
                    self.alterInvoiceTable.setItem(row, 6, QTableWidgetItem(str(functionKey)))
                    self.alterInvoiceTable.setItem(row, 7, QTableWidgetItem(str(cashupNO)))
                    row += 1
                    #print(productCode, productName, qty, salesTax, lineTotal, functionKey)
                else:
                    self.alterInvoiceTable.setRowCount(self.alterInvoiceTable.rowCount()-1)
        except:
            self.displaySimple("Failed")

    def alterCashups(self):
        pass

    def addRemoveInvoiceTIme(self):
        self.toolFrame = QFrame()
        self.toolFrame.setWindowTitle("Platinum Sql Date Corrector")
        self.toolFrame.setWindowIcon(QIcon("C:\Platinum\Icons\icon_p.ico"))
        self.toolFrame.resize(600, 300)

        startDate = QDateEdit()
        startDate.setDisplayFormat("yyyy-MM-dd")
        startTime = QTimeEdit()
        startTime.setDisplayFormat("HH:mm:ss")

        endDate = QDateEdit()
        endDate.setDisplayFormat("yyyy-MM-dd")
        endTime = QTimeEdit()
        endTime.setDisplayFormat("HH:mm:ss")

        timeToAdd = QTimeEdit()
        timeToAdd.setDisplayFormat("HH:mm:ss")

        runButton = QPushButton("Add Script")
        runButton.clicked.connect(lambda : self.runScript(startDate.text(), startTime.text(), endDate.text(), endTime.text(), timeToAdd.text()))

        subTractScript = QPushButton("Subtract Script")
        subTractScript.clicked.connect(
            lambda: self.runScript2(startDate.text(), startTime.text(), endDate.text(), endTime.text(),
                                   timeToAdd.text()))

        maingrid = QGridLayout()
        maingrid.addWidget(QLabel("Start Date"), 0, 0)
        maingrid.addWidget(startDate, 0, 1)
        maingrid.addWidget(QLabel("Start Time"), 1, 0)
        maingrid.addWidget(startTime, 1, 1)
        maingrid.addWidget(QLabel("End Date"), 0, 2)
        maingrid.addWidget(endDate, 0, 3)
        maingrid.addWidget(QLabel("End Time"), 1, 2)
        maingrid.addWidget(endTime, 1, 3)
        maingrid.addWidget(QLabel("Time"), 2, 0)
        maingrid.addWidget(timeToAdd, 2, 1)
        maingrid.addWidget(runButton, 3, 0, 1, 4)
        maingrid.addWidget(subTractScript, 4, 0, 1, 4)

        self.toolFrame.setLayout(maingrid)
        self.toolFrame.show()

    # Add invoice TIme script
    def runScript(self, sD, sT, eD, eT, tA):
        try:
            start_Time = sD + " " + sT
            endTime = eD + " " + eT
            time_Add = tA
            runner = platinumDatabase()
            runner.script(start_Time, endTime, time_Add)
            self.displaySimple("Success")
        except:
            self.displaySimple("Failed")

    # Remove invoice TIme Script
    def runScript2(self, sD, sT, eD, eT, tA):
        try:
            start_Time = sD + " " + sT
            endTime = eD + " " + eT
            time_Add = tA
            runner = platinumDatabase()
            runner.script2(start_Time, endTime, time_Add)
            self.displaySimple("Success")
        except:
            self.displaySimple("Failed")

    def displaySimple(self, text="No TExt"):
        self.messageBox = QMessageBox()
        self.messageBox.setWindowTitle("Message")
        self.messageBox.setWindowIcon(QIcon("C:\Platinum\Icons\icon_p.ico"))
        self.messageBox.setText(text)
        self.messageBox.show()

app = QApplication(sys.argv)
p = PlatinumCor()
sys.exit(app.exec_())