#!/usr/bin/python

from PyQt4.QtCore import *
from PyQt4.QtGui import *
from ui_utils import *


import sys
import re

import XLSProc
import ui_utils

_print_enabled = True

def xprint(s):
    if _print_enabled:
        print s


class SpreadSheetModel(QAbstractTableModel):
    defaultRowCount = XLSProc.SpreadSheetQuickSheet.rowCount
    defaultColCount = XLSProc.SpreadSheetQuickSheet.colCount
    
    logData = pyqtSignal(str,str,str)
    
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
            return self.rawData(key)
        
        return QString()

    def setData(self,index,value,role = Qt.DisplayRole,noLog = False):
        if role == Qt.DisplayRole or role == Qt.EditRole:
            key="%d,%d" % (index.row(),index.column())
            prev = self.rawData(key)
            try:
                value = str(value.toString())
            except:
                value = str(value)
            self.setRawData(key,value)
            self.dataChanged.emit(index,index)
            if not noLog:
                self.logData.emit(key,prev,value)
            return True
        return False
    
    def setRawData(self,key,value,emitSignal = False):
        self._data[key] = value
        if not self._init_data.has_key(key):
            self._init_data[key] = '' 
    
    def setValue(self,key,value,noLog = True):
        a = str(key).split(',')
        index = self.index(int(a[0]),int(a[1]),QModelIndex())
        self.setData(index, QVariant(str(value)),Qt.DisplayRole,noLog)
            
    def rawData(self,key):
        if self._data.has_key(key):
            return self._data[key]
        else:
            return ''
        
    def flags(self,index):
        return Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled
    
    def dump(self):
        return self._data
    
    def diff(self):
        rtn = {}
        for key in self._data:
            key = str(key)
            val = re.sub(r'^\s+','',str(self._data[key]))
            val = re.sub(r'\s+$','',val)
            
            val_init = re.sub(r'^\s+','',str(self._init_data[key]))
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
    


class LogManager(QObject):
    def __init__(self,name,model,parent= None):
        QObject.__init__(self,parent)
        self.name = name
        self.dataSource = model
        self.dataSource.logData.connect(self.onDataChange)
        self.lastKey = ''
        self.lastValue = ''
        self.lastPreValue = ''
        
        self.localStack = {}
        self.localStackCurPtr = -1
        self.localStackEndPtr = -1

    def onDataChange(self,key,prevalue,value):
        if not key == self.lastKey:
            self.flush()
            self.lastKey = key
            self.lastPreValue = prevalue
            
        self.lastValue = value        
    def flush(self,*kargs):
        if self.lastKey == '':
            return
        
        xprint("set_cell_value '%s' %s # %s" % (self.lastKey,self.lastValue,self.lastPreValue))
        
        self.localStackCurPtr += 1
        self.localStack[self.localStackCurPtr] = {
            'dataSource' : self.dataSource,
            'key' : self.lastKey,
            'prev' : self.lastPreValue,
            'curv' : self.lastValue
        }
        self.localStackEndPtr = self.localStackCurPtr
                
        self.lastKey = ''

    def undo(self):
        self.flush()
        if self.canUndo():           
            item = self.localStack[self.localStackCurPtr]
            xprint("set_cell_value '%s' %s # undo %d" % (item['key'],item['prev'],self.localStackCurPtr))
            item['dataSource'].setValue(item['key'],item['prev'],True)
            self.localStackCurPtr -= 1
        
    def canUndo(self):
        return self.localStackCurPtr >= 0
        
    def redo(self):
        self.flush()
        if self.canRedo():
            self.localStackCurPtr += 1
            item = self.localStack[self.localStackCurPtr]
            xprint("set_cell_value '%s' %s # redo %d" % (item['key'],item['curv'],self.localStackCurPtr))
            item['dataSource'].setValue(item['key'],item['curv'],True)
                        

    def canRedo(self):
        return self.localStackCurPtr < self.localStackEndPtr        
            

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
        self.tabFrame.setDocumentMode(True)
        self.tabFrame.currentChanged.connect(self.onCurrentTabChanged)
        
        layMain = QVBoxLayout()
        layMain.addWidget(self.cellEditor)
        layMain.addWidget(self.tabFrame)

        layMain.setContentsMargins(0,0,0,0)
        layMain.setSpacing(0)
        
        self.mainWidget = QWidget()
        
        
        self.mainWidget.setLayout(layMain)
        self.setCentralWidget(self.mainWidget)
        
        self.onFileLoad(True)
        
        self.logManagers = []       
        self.updateStatusBarMessage('Ready')
    
    def onCurrentTabChanged(self,sheetid):
        xprint("sheet_active %d # %s" % (sheetid,self.tabFrame.tabText(sheetid)))
        self.cellEditor.onSheetIdChange(sheetid)
            
    def createSlotOnTableCellChange(self,sheet):
        def slotOnTableCellChange(self,xrow,xcol):
            sheet.set_cell_value()
            
    def onFileLoad(self,isNew = False):

        

        if not isNew:        
            fname = str(self.fileName())
            xprint("open '%s'" % fname)
            self.workbook = XLSProc.SpreadSheetQuick(fname,self)     
            if self.workbook.fmt == '':
                self.loadFinished(False)
                self.updateStatusBarMessage("Error: Cannot open %s" % fname)
                return
        
        else:
            self.workbook = XLSProc.SpreadSheetQuick(None,self)
            fname = ''
            xprint('new')
        self.setCursor(Qt.BusyCursor)
        self.tabFrame.clear()
        self.logManagers = []
            
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
            
            l = LogManager(sheet_name,m,self)
            self.logManagers.append(l)
            w.selectionModel().currentChanged.connect(l.flush)
            self.tabFrame.currentChanged.connect(l.flush)

        if not self.fileCreateByMe():
            self.activeTab = 0
            
        self.tabFrame.setCurrentIndex(self.activeTab)
        self.onCurrentTabChanged(self.activeTab)
        self.tabFrame.setFocus()
        self.updateStatusBarMessage('Ready')
        if not isNew:
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
    
    def onEditUndo(self):
        self.logManagers[self.tabFrame.currentIndex()].undo()
    def onEditRedo(self):
        self.logManagers[self.tabFrame.currentIndex()].redo()
    
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
