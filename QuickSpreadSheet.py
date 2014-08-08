#!/usr/bin/python

from PyQt4.QtCore import *
from PyQt4.QtGui import *

import sys
import re

import XLSProc
import ui_utils


#tableHeaders = [ XLSProc.XlsHeader(i) for i in range(0, sheet.col_count())]
#w.setHorizontalHeaderLabels(tableHeaders)


     #   except:
      #      print "Error in %d %d" % (row,col)

class SpreadSheetModel(QAbstractTableModel):
    defaultRowCount = 50
    defaultColCount = 24
    
    def __init__(self,sheet,parent=None):
        super(QAbstractTableModel, self).__init__(parent)
        self.sheet = sheet
        self.row_count = max(self.defaultRowCount,sheet.row_count())
        self.col_count = max(self.defaultColCount,sheet.col_count())
        self._data = {}
        self._init_data = {}

        for row in range(0, sheet.row_count()):
            for col in range(0, sheet.col_count()):
                key="%d,%d" % (row,col)
                self._data[key] = sheet.cell_value(row,col)
                self._init_data[key] = sheet.cell_value(row,col)
                
    def rowCount(self,parent = None):
        return self.row_count
    
    def columnCount(self,parent = None):
        return self.col_count
    
    def headerData(self,section,orientation,role):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return XLSProc.XlsHeader(section)
            else:
                return "%d" % section
        return QVariant()
            
    def data(self,index,role = Qt.DisplayRole):
        if not index.isValid():
            return QVariant()
        
        if role == Qt.TextAlignmentRole:
            return int(Qt.AlignLeft | Qt.AlignVCenter)
        
        elif role == Qt.DisplayRole or role == Qt.EditRole:
            key="%d,%d" % (index.row(),index.column())
            if self._data.has_key(key):
                return self._data[key]

        return QVariant()

    def setData(self,index,value,role = Qt.DisplayRole):
        if role == Qt.DisplayRole or role == Qt.EditRole:
            key="%d,%d" % (index.row(),index.column())
            self._data[key] = value
            if not self._init_data.has_key(key):
                self._init_data[key] = ''
            self.dataChanged.emit(index,index)
            return True
        return False
    
    def flags(self,index):
        return Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled
    
    def dump(self):
        return self._data
        

class MainUI(ui_utils.QXSingleDocMainWindow):


    
    def __init__(self,parent=None):
        super(ui_utils.QXSingleDocMainWindow, self).__init__(parent)
        self.initUI()
        self.initDefaultUI()

    def initUI(self):
        self.mainWidget = QTabWidget()
        self.mainWidget.setTabPosition(QTabWidget.South)
        self.setCentralWidget(self.mainWidget)
        
        for x in range(3):
            w = QTableView()
            m = SpreadSheetModel(XLSProc.SpreadSheetQuickSheet())
            w.setModel(m)
            self.mainWidget.addTab(w,"Sheet%d" % (x + 1))
            
        self.updateStatusBarMessage('Ready')
            
            
    def createSlotOnTableCellChange(self,sheet):
        def slotOnTableCellChange(self,xrow,xcol):
            sheet.set_cell_value()
            
    def onFileOpen(self):
        fname = self.fileName()
        self.mainWidget.clear()
        self.workbook = XLSProc.SpreadSheetQuick(fname)
        
        for sheet_name in self.workbook.worksheets():
            w = QTableView()
            sheet = self.workbook.worksheet(sheet_name)
            m = SpreadSheetModel(sheet)
            w.setModel(m)        
            self.mainWidget.addTab(w,sheet_name)
        self.updateStatusBarMessage('Ready')
        self.loadFinished()
        
    def onFileSaveAs(self, fileName):
        xworkbook = []
        for i in range(self.mainWidget.count()):
            w = self.mainWidget.widget(i)
            xworkbook.append({
                'name' : self.mainWidget.tabText(i),
                'data' : w.model().dump()
            })
        XLSProc.createWorkBook(xworkbook, fileName)
    
        
if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    fname = sys.argv[1]
    w = MainUI()
    w.show()
    w.ActionFileOpen(fname)
    exit(app.exec_())
