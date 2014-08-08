#!/usr/bin/python

from PyQt4.QtCore import *
from PyQt4.QtGui import *
from ui_utils import *

import sys
import re

import XLSProc
import ui_utils


class SpreadSheetModel(QAbstractTableModel):
    defaultRowCount = 50
    defaultColCount = 24
    
    def __init__(self,sheet,parent=None):
        QAbstractTableModel.__init__(self,parent)
        
        self.sheet = sheet
        self.row_count = max(self.defaultRowCount,sheet.row_count())
        self.col_count = max(self.defaultColCount,sheet.col_count())
        self._data = {}
        self._init_data = {}

        for row in range(0, sheet.row_count()):
            for col in range(0, sheet.col_count()):
                key="%d,%d" % (row,col)
                self._data[key] = str(sheet.cell_value(row,col))
                self._init_data[key] = str(sheet.cell_value(row,col))
                
    def rowCount(self,parent = None):
        return self.row_count
    
    def columnCount(self,parent = None):
        return self.col_count
    
    def headerData(self,section,orientation,role):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return XLSProc.XlsHeader(section)
            else:
                return "%d" % (section + 1)
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
            self._data[key] = str(value.toString())
            if not self._init_data.has_key(key):
                self._init_data[key] = ''
            self.dataChanged.emit(index,index)
            return True
        return False
    
    def flags(self,index):
        return Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled
    
    def dump(self):
        return self._data
    
    def diff(self):
        rtn = {}
        for key in self._data:
            val = re.sub(r'^\s+','',self._data[key])
            val = re.sub(r'\s+$','',val)
            
            val_init = re.sub(r'^\s+','',self._init_data[key])
            val_init = re.sub(r'\s+$','',val_init)
            if val != val_init:
                rtn[key] = val
        return rtn
    
    def sync(self):
        self._init_data = self._data
                
        

class MainUI(QXSingleDocMainWindow):

    def __init__(self,parent=None):
        QXSingleDocMainWindow.__init__(self,parent)
        self.initUI()
       
    def initUI(self):

        self.setFileSaveAsSuffix("All Support (*.xlsx *.xls *.tsv *.csv);;Excel WorkBook (*.xlsx);;Excel 1997 - 2003 WorkBook (*.txt);;Tab Seperated Value (*.tsv);;Comma Seperated Value (*.csv)")
        
        self.mainWidget = QTabWidget()
        self.mainWidget.setTabPosition(QTabWidget.South)
        self.setCentralWidget(self.mainWidget)
        
        for x in range(3):
            w = QTableView()
            m = SpreadSheetModel(XLSProc.SpreadSheetQuickSheet(),self)
            w.setModel(m)
            self.mainWidget.addTab(w,"Sheet%d" % (x + 1))
            
        self.updateStatusBarMessage('Ready')
            
            
    def createSlotOnTableCellChange(self,sheet):
        def slotOnTableCellChange(self,xrow,xcol):
            sheet.set_cell_value()
            
    def onFileLoad(self):
        fname = str(self.fileName())
        
        self.workbook = XLSProc.SpreadSheetQuick(fname,self)
        
        if self.workbook.fmt == '':
            self.loadFinished(False)
            self.updateStatusBarMessage("Error: Cannot open %s" % fname)
            return
        
        self.mainWidget.clear()
            
        if re.match(r'.*\.xls$',fname.lower()) and not self.fileCreateByMe():
            self.setFileReadOnly(True)
        else:
            self.setFileReadOnly(False)
            
        for sheet_name in self.workbook.worksheets():
            w = QTableView()
            sheet = self.workbook.worksheet(sheet_name)
            m = SpreadSheetModel(sheet,self)
            w.setModel(m)        
            self.mainWidget.addTab(w,sheet_name)

        if self.fileCreateByMe():
            self.mainWidget.setCurrentIndex(self.activeTab)
            
        self.updateStatusBarMessage('Ready')
        self.loadFinished(True)

    def getXWorkBook(self):
        xworkbook = []
        for i in range(self.mainWidget.count()):
            w = self.mainWidget.widget(i)
            xworkbook.append({
                'name' : str(self.mainWidget.tabText(i)),
                'data' : w.model().dump(),
                'diff' : w.model().diff()
            })
        return xworkbook
    
    def modelSync(self):
        for i in range(self.mainWidget.count()):
            w = self.mainWidget.widget(i)
            w.model().sync()
        
    def onFileSaveAs(self, fileName):
        self.activeTab = self.mainWidget.currentIndex()
        if XLSProc.SpreadSheetQuick.create(self.getXWorkBook(), str(fileName)):
            self.modelSync()
            return True
    
    def onFileSave(self, fileName):
        self.activeTab = self.mainWidget.currentIndex()
                   
        if XLSProc.SpreadSheetQuick.save(self.getXWorkBook(),self.workbook,str(fileName)):
            self.modelSync()
            self.updateStatusBarMessage("Saved %s at %s" % (self.fileName() , QDateTime.currentDateTime().toString()))
            return True            
    
        
if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    fname = sys.argv[1]
    w = MainUI()
    w.show()
    w.ActionFileLoad(fname)
    exit(app.exec_())
