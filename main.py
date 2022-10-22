# -*- coding: utf-8 -*-
"""
@author: Mr. Sha2w
sha2w.ir
+989117569002
sha2w@yahoo.com
#about: Ticket Checker
"""
import os
import sys
from os.path import dirname, realpath, join
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5 import QtGui, QtCore
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QTableWidget, QTableWidgetItem
from PyQt5.uic import loadUiType
import pandas as pd
import random, string

scriptDir = dirname(realpath(__file__))
From_Main, _ = loadUiType(join(dirname(__file__), "ticketchecker.ui"))


class MainWindow(QWidget, From_Main):
    def __init__(self):
        super(MainWindow, self).__init__()
        QWidget.__init__(self)
        self.setupUi(self)

        self.uploadbtn.clicked.connect(self.OpenFile)
        self.importbtn.clicked.connect(self.dataHead)
        
        self.customerid.setPlaceholderText("Search...")
        self.customerid.textChanged.connect(self.search)
        count=0
        self.ticketserial.returnPressed.connect(self.checkinbtn.click) 
        self.ticketserial.setPlaceholderText("Search...")
        self.checkinbtn.clicked.connect(self.check_in,count)
        self.exportexclbtn.clicked.connect(self.export_to_excel)
        self.deletedatabasebtn.clicked.connect(self.clear_all)
        self.exitbtn.clicked.connect(self.closeIt)
        self.staticdetail.clicked.connect(self.static_detail)

    def OpenFile(self):
        path = QFileDialog.getOpenFileName(self, 'Open CSV', os.getenv('HOME'), 'CSV(*.csv)')[0]
        if not path:
            self.reports.setText('There is no list!')
            self.reports.setStyleSheet("color:red")
        else:
            self.all_data = pd.read_csv(path)
            if "ورود" not in self.all_data.columns:
                self.all_data["check-in"] = 0
            self.csvdirection.setText(path)
            self.reports.setText('')
            
            
            start=self.tablenumbstart.value() -1
            finish = self.tablenumberfinish.value()
            
            if (start<=0 and finish<=0 or finish<start):
                start= 0 
                finish =len(self.all_data.index)
                self.reports.setText('List limits!') #an option when you have multi gate for entering attenders
                self.reports.setStyleSheet("color:blue")
            else:
                self.all_data=self.all_data.iloc[start:finish,:]
                
            for c in range(0, len(self.all_data.columns)):
                for r in range(start, finish):
                    s = ''
                    i = QTableWidgetItem(s)
                    self.tableWidget.setItem(c, r, i)
                    

    def dataHead(self):
        if not self.all_data.empty:
            NumRows = len(self.all_data.index)
            self.allticketnumb.setText(str(NumRows))
            self.tableWidget.setColumnCount(len(self.all_data.columns))
            self.tableWidget.setRowCount(NumRows)
            self.tableWidget.setHorizontalHeaderLabels(['Name','Code','Location','PersID','Enter'])
            self.uploadtable.setColumnCount(len(self.all_data.columns))
            self.uploadtable.setRowCount(NumRows)
            self.uploadtable.setHorizontalHeaderLabels(['Name','Code','Location','PersID','Enter'])
        
            for i in range(NumRows):
                for j in range(len(self.all_data.columns)):
                    self.uploadtable.setItem(i, j, QTableWidgetItem(str(self.all_data.iat[i, j])))
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(self.all_data.iat[i, j])))
                    if (j==4 and (self.all_data.iat[i, j] == 1 or self.all_data.iat[i, j] == "Entered!")):
                        checkin="Entered"
                        self.tableWidget.item(i,j).setBackground(QtGui.QColor(0,100,0))
                        self.tableWidget.item(i,j).setText(checkin)
                        self.uploadtable.item(i,j).setBackground(QtGui.QColor(0,100,0))
                        self.uploadtable.item(i,j).setText(checkin)
                    elif (j==4 and (self.all_data.iat[i, j] == 0 or self.all_data.iat[i, j] == 'Nan')):
                        checkin="Nan"
                        self.tableWidget.item(i,j).setText(checkin)
                        self.uploadtable.item(i,j).setText(checkin)
            
            self.uploadtable.resizeColumnsToContents()
            self.uploadtable.resizeRowsToContents()
            self.tableWidget.resizeColumnsToContents()
            self.tableWidget.resizeRowsToContents()
            
            self.reports.setText('Attender information have been registered')
            self.reports.setStyleSheet("color:green")
        else:
            self.reports.setText('Initially add the list then select the Next button!')
            self.reports.setStyleSheet("color:red")
    
    def check_in(self, count):
        self.reports.setText('')
        global all_data
        
        # global count
        serial = str(self.ticketserial.text())
        
        if (serial != ''):
            matching_items = self.tableWidget.findItems(
            self.ticketserial.text(), QtCore.Qt.MatchExactly)
            if matching_items:
            # We found something.
                
                i = matching_items[0].row()
                j = matching_items[0].column()
                
                self.atenname.setText(str(self.all_data.iat[i, 0]))
                self.atenticserial.setText(str(self.all_data.iat[i, 1]))
                self.atenposit.setText(str(self.all_data.iat[i, 2]))
                self.ateidcode.setText(str(self.all_data.iat[i, 3]))
                if (self.all_data.iat[i, 4] == 0 or self.all_data.iat[i, 4] == 'Nan'):
                    if (self.all_data.iat[i, 4] == 0):
                        self.all_data.iat[i, 4] += 1
                    if (self.all_data.iat[i, 4] == 'Nan'):
                        self.all_data.iat[i, 4] = "Entered"
                            
                    checkin="Ticket is ok. Entered!."
                    
                    checkin_cell="Entered"
                    self.atencheckin.setStyleSheet("background-color: green; border: 1px solid black;")
                    self.atencheckin.setText(checkin)
                    self.uploadtable.setItem(i, 4, QTableWidgetItem(str(self.all_data.iat[i, 4])))
                    self.tableWidget.setItem(i, 4, QTableWidgetItem(str(self.all_data.iat[i, 4])))
                    self.tableWidget.item(i,4).setBackground(QtGui.QColor(0,100,0))
                    self.tableWidget.item(i,4).setText(checkin_cell)
                    self.uploadtable.item(i,4).setBackground(QtGui.QColor(0,100,0))
                    self.uploadtable.item(i,4).setText(checkin_cell)
                    
                elif (self.all_data.iat[i, 4] == 1 or self.all_data.iat[i, 4] == "ورود کرد"):
                    checkin="The ticket is expired! \n Ticker reuse!!"
                    self.atencheckin.setStyleSheet("background-color: red;border: 1px solid black;")
                    self.atencheckin.setText(checkin)
                
            else:
                self.atenname.setText("")
                self.atenticserial.setText("")
                self.atenposit.setText("")
                self.ateidcode.setText("")
                self.atencheckin.setText("")
                self.atencheckin.setStyleSheet("background-color: ; border: ;")
                
        else:
            return(-1)
                
    def static_detail(self):
        if (self.tableWidget.rowCount() > 0):
            checkinnumb = len(self.all_data[self.all_data['Enter']=='Entered'])
            self.checkinnumb.setText(str(checkinnumb))
            self.notchechkinnumb.setText(str(len(self.all_data.index)-checkinnumb))
        else:
            self.reports.setText('There is no information')
            self.reports.setStyleSheet("color:red") 
    def search(self, s):
        self.reports.setText('')
        self.tableWidget.setCurrentItem(None)
        if not s:
            # Empty string, don't search.
            return
        matching_items = self.tableWidget.findItems(s, Qt.MatchContains)
        if matching_items:
            # We have found something.
            item = matching_items[0]  # Take the first.
            self.tableWidget.setCurrentItem(item)
            rowItem= self.tableWidget.currentRow()
            self.searchreportlab.setText('line {} with ticket code {}'.format(rowItem+1,self.all_data.iat[rowItem, 1]))
            self.searchreportlab.setStyleSheet("color:green")
        else:    
            self.searchreportlab.setText('No ticket founded!')
            self.searchreportlab.setStyleSheet("color:red")
        
    
    def export_to_excel(self):
        self.reports.setText('')
        if (self.tableWidget.rowCount() > 0):
            self.reports.setText('Export the list')
            self.reports.setStyleSheet("color:green")  
            
            name = QFileDialog.getSaveFileName(self, 'Save File', filter='*.csv')
            if(name[0] == ''):
                pass
            else:
                columnHeaders = []
    
                # create column header list
                for j in range(self.tableWidget.model().columnCount()):
                    columnHeaders.append(self.tableWidget.horizontalHeaderItem(j).text())
        
                df = pd.DataFrame(columns=columnHeaders)
                # create dataframe object recordset
                for row in range(self.tableWidget.rowCount()):
                    for col in range(self.tableWidget.columnCount()):
                        df.at[row, columnHeaders[col]] = self.tableWidget.item(row, col).text()
        
                df.to_csv(name[0], index = False)
                print('Excel file exported')
        else:
            self.reports.setText('There is no list to view!')
            self.reports.setStyleSheet("color:red")
            
    def clear_all(self):
        self.uploadtable.clear()
        self.uploadtable.setRowCount(0);
        self.tableWidget.clear()
        self.tableWidget.setRowCount(0);
        self.reports.setText('The list has been removed correctly')
        self.reports.setStyleSheet("color:green")          
        
    def closeIt(self): 
         self.close()  


app = QApplication(sys.argv)
sheet = MainWindow()
sheet.show()
sys.exit(app.exec_())

