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
        return QString()
            
    def data(self,index,role = Qt.DisplayRole):
        if not index.isValid():
            return QVariant()
        
        if role == Qt.TextAlignmentRole:
            return int(Qt.AlignLeft | Qt.AlignVCenter)
        
        elif role == Qt.DisplayRole or role == Qt.EditRole:
            key="%d,%d" % (index.row(),index.column())
            if self._data.has_key(key):
                return self._data[key]

        return QString()

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
                
class CellEditor(QWidget):
    
    requestData = pyqtSignal(int,QModelIndex)
    setData = pyqtSignal(int,int,int,str)
    
    def __init__(self,parent=None):
        QWidget.__init__(self,parent)
        
        self.cbxCellAddress = QComboBox()
        self.cbxCellAddress.setMinimumWidth(100)
        self.cbxCellAddress.setMaximumWidth(100)
        
        self.edtCellEditor = QComboBox()
        self.edtCellEditor.setEditable(False)
        self.edtCellEditor.lock = True
        self.edtCellEditor.editTextChanged.connect(self.onCellEditorContentChanged)
        
        layout = QHBoxLayout()
        layout.addWidget(self.cbxCellAddress)
        layout.addWidget(self.edtCellEditor)
        layout.setContentsMargins(4,4,4,4)
        layout.setSpacing(4)        
        self.setLayout(layout)

        
    def setCellAddress(self,s):
        self.cbxCellAddress.insertItem(0,s)
        self.cbxCellAddress.setCurrentIndex(0)
        
    def setCellEditorContent(self,s):
        self.edtCellEditor.lock = True
        self.edtCellEditor.setEditable(True)
        self.edtCellEditor.setEditText(s)

        self.edtCellEditor.lock = False

    def onCellEditorContentChanged(self,s):
        if not self.edtCellEditor.lock:
            self.setData.emit(self.currentSheetId,self.currentRow,self.currentCol,s)
            
    def onSheetIdChange(self,sheetId):
        self.edtCellEditor.lock = True
        self.edtCellEditor.setEditable(False)        
    
    def cloOnCurrentIndexChange(self,sheetId):
        def onCurrentIndexChange(currentIndex,prevIndex):
            self.currentSheetId = sheetId
            self.currentRow = currentIndex.row()
            self.currentCol = currentIndex.column()
            self.setCellAddress(XLSProc.XlsHeader(self.currentCol,self.currentRow))
            self.requestData.emit(sheetId,currentIndex)
        return onCurrentIndexChange
    
        

class MainUI(QXSingleDocMainWindow):

    def __init__(self,parent=None):
        QXSingleDocMainWindow.__init__(self,parent)
        self.initUI()
       
    def initUI(self):

        self.setFileSaveAsSuffix("All Support (*.xlsx *.xls *.tsv *.csv);;Excel WorkBook (*.xlsx);;Excel 1997 - 2003 WorkBook (*.txt);;Tab Seperated Value (*.tsv);;Comma Seperated Value (*.csv)")

        self.cellEditor = CellEditor()
        self.cellEditor.requestData.connect(self.onGetSheetCellDataToCellEditor)
        self.cellEditor.setData.connect(self.onCellEditorRequestDataChange)
        
        self.tabFrame = QTabWidget()
        self.tabFrame.setTabPosition(QTabWidget.South)
        self.tabFrame.currentChanged.connect(self.cellEditor.onSheetIdChange)
        

        
        layMain = QVBoxLayout()
        layMain.addWidget(self.cellEditor)
        layMain.addWidget(self.tabFrame)

        layMain.setContentsMargins(0,0,0,0)
        layMain.setSpacing(0)
        
        self.mainWidget = QWidget()
        
        
        self.mainWidget.setLayout(layMain)
        self.setCentralWidget(self.mainWidget)
        
        for x in range(3):
            w = QTableView()
            m = SpreadSheetModel(XLSProc.SpreadSheetQuickSheet(),self)
            w.setModel(m)
            self.tabFrame.addTab(w,"Sheet%d" % (x + 1))
            
        self.updateStatusBarMessage('Ready')
            
            
    def createSlotOnTableCellChange(self,sheet):
        def slotOnTableCellChange(self,xrow,xcol):
            sheet.set_cell_value()
            
    def onFileLoad(self):

        self.setCursor(Qt.BusyCursor)
        
        fname = str(self.fileName())
        
        self.workbook = XLSProc.SpreadSheetQuick(fname,self)
        
        if self.workbook.fmt == '':
            self.loadFinished(False)
            self.updateStatusBarMessage("Error: Cannot open %s" % fname)
            return
        
        self.tabFrame.clear()
            
        if re.match(r'.*\.xls$',fname.lower()) and not self.fileCreateByMe():
            self.setFileReadOnly(True)
        else:
            self.setFileReadOnly(False)
            
        for sheet_id,sheet_name in enumerate(self.workbook.worksheets()):
            w = QTableView()
            sheet = self.workbook.worksheet(sheet_name)
            m = SpreadSheetModel(sheet,self)
            w.setModel(m)
            m.dataChanged.connect(self.onModelDataChanged)
            self.tabFrame.addTab(w,sheet_name)
            w.selectionModel().currentChanged.connect(self.cellEditor.cloOnCurrentIndexChange(sheet_id))

        if not self.fileCreateByMe():
            self.activeTab = 0
            
        self.tabFrame.setCurrentIndex(self.activeTab)       
        self.tabFrame.setFocus()
        self.updateStatusBarMessage('Ready')
        self.loadFinished(True)
        self.unsetCursor()
        
    def onGetSheetCellDataToCellEditor(self,sheetid,index):
        w = self.tabFrame.widget(sheetid)
        self.cellEditor.setCellEditorContent(w.model().data(index))

    def onCellEditorRequestDataChange(self,sheetid,row,col,value):
        w = self.tabFrame.widget(sheetid)
        m = w.model()
        i = m.index(row,col,QModelIndex())
        m.setData(i,QVariant(value))      

    def getXWorkBook(self):
        xworkbook = []
        for i in range(self.tabFrame.count()):
            w = self.tabFrame.widget(i)
            xworkbook.append({
                'name' : str(self.tabFrame.tabText(i)),
                'data' : w.model().dump(),
                'diff' : w.model().diff()
            })
        return xworkbook
    
    def modelSync(self):
        for i in range(self.tabFrame.count()):
            w = self.tabFrame.widget(i)
            w.model().sync()
        
    def onModelDataChanged(self,index1,index2):
        self.setWindowModified(True)
        
    def onFileSaveAs(self, fileName):
        
        self.activeTab = self.tabFrame.currentIndex()
        if XLSProc.SpreadSheetQuick.create(self.getXWorkBook(), str(fileName)):
            self.setCursor(Qt.BusyCursor)
            self.modelSync()
            self.setWindowModified(False)
            self.unsetCursor()
            return True
    
    def onFileSave(self, fileName):
        self.activeTab = self.tabFrame.currentIndex()
                   
        if XLSProc.SpreadSheetQuick.save(self.getXWorkBook(),self.workbook,str(fileName)):
            self.setCursor(Qt.BusyCursor)
            self.modelSync()
            self.updateStatusBarMessage("Saved %s at %s" % (self.fileName(),QDateTime.currentDateTime().toString()))
            self.setWindowModified(False)
            self.unsetCursor()
            return True            
    
        
if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    w = MainUI()
    w.show()
    try:
        fname = sys.argv[1]
        w.ActionFileLoad(fname)
    except:
        pass
    app.exec_()
