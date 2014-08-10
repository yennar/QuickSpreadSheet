#!/usr/bin/python

from PyQt4.QtCore import *
from PyQt4.QtGui import *
import platform
import sys
import re
import ui_utils_res
import json

class QXInputDialog(QInputDialog):
    
    @staticmethod 
    def getMulti(parent,title,label,items,defaultVal = None,flags=0):
        dlg = QDialog(parent)
        dlg.setWindowTitle(title)
        if flags !=0:
            dlg.setWindowFlags(flags)
        
        mainL = QVBoxLayout()
        mainL.addWidget(QLabel(label))
        
        formL = QFormLayout()
        edits = {}
        for item in items:
            edit = QLineEdit()
            
            if item[-1] == '*':
                item = item[0:len(item) - 1]
                edit.setEchoMode(QLineEdit.Password)
            
            edits[item] = edit    
            formL.addRow(item,edit)
                
            
            try:
                edit.setText(defaultVal[item])
            except:
                pass
            
        mainL.addLayout(formL)
        buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttonBox.accepted.connect(dlg.accept)
        buttonBox.rejected.connect(dlg.reject)
        mainL.addWidget(buttonBox)
        dlg.setLayout(mainL)
        r = dlg.exec_()
        if r == QDialog.Accepted:
            result = {}
            for item in items:
                if item[-1] == '*':
                    item = item[0:len(item) - 1]
                result[item] = str(edits[item].text())
            return result
        else:
            return None          

class QXAction(QAction):

   
    def __init__(self,*kargs,**kwargs):
        QAction.__init__(self,*kargs,**kwargs)
        kq = QKeySequence.mnemonic(self.text())
        if kq.isEmpty():
            if self.text().contains('&'):
                t = self.text().replace('&','')
                try:
                    kq = QKeySequence(eval('QKeySequence.%s' % t))
                except:
                    pass
        self.setShortcut(kq)
        
        if not kq.isEmpty():
            self.setToolTip("%s (%s)" % (self.text().replace('&',''),kq.toString(QKeySequence.NativeText)))
        
        image_name = "%s.png" % self.text().replace('&','').toLower()
        dir = QDir(QDir.currentPath());
        if dir.exists(image_name):
            self.setIcon(QIcon(image_name))
        else:
            image_name = ":" + image_name
            if dir.exists(image_name):
                self.setIcon(QIcon(image_name))            

_mime_type = 'application/x-qt-windows-mime;value="qspreadsheet_blocks"'

class QXTableView(QTableView):

    beginRemove = pyqtSignal(bool)
    endRemove = pyqtSignal(bool)
    dataChangeMulti = pyqtSignal(list)
    
    def __init__(self,*kargs):
        QTableView.__init__(self,*kargs)
    
       
    def copySelection(self,doCut = False,mode = QClipboard.Clipboard):
        if self.model() is None:
            return
        selction = self.selectionModel()
        indexes = selction.selectedIndexes()
        
        logData = []
        values = []
        clipData = []
        for index in indexes:
            value = self.model().data(index)
            key = "%d,%d" % (index.row(),index.column())
            logData.append({
                'dataSource' : self.model(),
                'key' : key,
                'prev' : value,
                'curv' : ''
            })
            values.append(value)
            clipData.append({
                'r' : index.row(),
                'c' : index.column(),
                'v' : value
            })

        clipBoard = QApplication.clipboard()
        
        mimeData = QMimeData()
        mimeData.setText(",".join(values))
        mimeData.setData(_mime_type,json.dumps(clipData))
        clipBoard.setMimeData(mimeData,mode)
        
        if doCut:
            self.beginRemove.emit(False)
            for index in indexes:
                self.model().setData(index,"")
                self.endRemove.emit(True)
            self.dataChangeMulti.emit(logData)
            
    def cutSelection(self):
        self.copySelection(True)
        
    def pasteSelection(self,mode = QClipboard.Clipboard):
        if self.model() is None:
            return
        selction = self.selectionModel()
        indexes = selction.selectedIndexes()
        
        #start point
        d = -1
        p_r = -1
        p_c = -1
        
        for index in indexes:
            this_d = index.row() + index.column()
            if d == -1 or this_d < d:
                d = this_d
                p_r = index.row()
                p_c = index.column()
            selction.select(index,QItemSelectionModel.Clear)
        
        if d == -1:
            index = selction.currentIndex()
            if not index is None:
                p_r = index.row()
                p_c = index.column()
                d = p_r + p_c
                selction.select(selction.currentIndex(),QItemSelectionModel.Clear)
        
        if d == -1:
            p_r = 0
            p_c = 0
            d = 0
        
        clipBoard = QApplication.clipboard()
        mimeData = QMimeData()
        clipData = []
        
        if mode == QClipboard.Clipboard:
            firstTryMode = QClipboard.Clipboard
            secondTryMode = QClipboard.Selection
        elif mode == QClipboard.Selection:
            firstTryMode = QClipboard.Selection
            secondTryMode = QClipboard.Clipboard            
            
        mimeData = clipBoard.mimeData(firstTryMode)
        if (not mimeData is None) and mimeData.hasFormat(_mime_type):
            clipData = json.loads(str(mimeData.data(_mime_type)))
        else:
            mimeData = clipBoard.mimeData(secondTryMode)
            if (not mimeData is None) and mimeData.hasFormat(_mime_type):
                clipData = json.loads(str(mimeData.data(_mime_type)))
            else:
                mimeData = clipBoard.mimeData(firstTryMode)
                if (not mimeData is None) and mimeData.hasText():
                    clipData = [{
                        'r' : 0,
                        'c' : 0,
                        'v' : mimeData.text()
                    }]
                else:
                    mimeData = clipBoard.mimeData(secondTryMode)
                    if (not mimeData is None) and mimeData.hasText():
                        clipData = [{
                            'r' : 0,
                            'c' : 0,
                            'v' : mimeData.text()
                        }]        
                    else:
                        clipData = []
                    
        

        m_r = -1
        m_c = -1
        
        for item in clipData:
            if m_r == -1 or item['r'] < m_r:
                m_r = item['r']
            if m_c == -1 or item['c'] < m_c:
                m_c = item['c']                
        
        dif_r = p_r - m_r
        dif_c = p_c - m_c
        
        logData = []
        self.beginRemove.emit(False)
        
        
        
        for item in clipData:
            index = self.model().index(item['r'] + dif_r,item['c'] + dif_c,QModelIndex())
            selction.select(index,QItemSelectionModel.Select)
            key = "%d,%d" % (index.row(),index.column())
            logData.append({
                'dataSource' : self.model(),
                'key' : key,
                'prev' : self.model().data(index),
                'curv' : item['v']
            })            
            self.model().setData(index,item['v'])
            self.endRemove.emit(True)
        self.dataChangeMulti.emit(logData)   
        
    def mousePressEvent(self,e):
        if e.button() == Qt.MiddleButton:
            self.pasteSelection(QClipboard.Selection)
        QTableView.mousePressEvent(self,e)
    
    def mouseReleaseEvent(self,e):
        QTableView.mouseReleaseEvent(self,e)
        if e.button() == Qt.LeftButton:
            self.copySelection(False,QClipboard.Selection)
        
    
class QXSingleDocMainWindow(QMainWindow):
    
    def __init__(self,parent=None):
        QMainWindow.__init__(self,parent)
        self.initDefaultUI()
        
        

    def initDefaultUI(self,hasToolBar = True,hasMenuBar = False):
        d = QApplication.desktop()
        screenWidget = d.screen(d.primaryScreen())
        w = screenWidget.size().height()
        h = screenWidget.size().height() * 0.6
        self.resize(w,h)
        import random,os
        random.seed(os.urandom(128))
        self.move(QPoint(
            random.randint(0,int((screenWidget.size().width()  - w) / 2)) ,
            random.randint(0,int((screenWidget.size().height() - h) / 2)))
            )
        
        self.appName = re.sub(r'^.*\/','',sys.argv[0])
        self.appName = re.sub(r'\..*$','',self.appName)
        
        self.setWindowTitle("Untitled[*] - %s" % self.appName)   
        self.setFileSaveAsSuffix("All Files (*.*)")
        self.setFileReadOnly(False)
        self._fileName = None
        self.setFileCreateByMe(False)

        self.actionFileNew = QXAction('&New',self,triggered=self.ActionFileNew)
        self.actionFileOpen = QXAction('&Open',self,triggered=self.ActionFileOpen)
        self.actionFileSave = QXAction('&Save',self,triggered=self.ActionFileSave)
        self.actionFileSaveAs = QXAction('Save &As',self,triggered=self.ActionFileSaveAs)
        
        self.actionEditUndo = QXAction('&Undo',self,triggered=self.onEditUndo)
        self.actionEditRedo = QXAction('&Redo',self,triggered=self.onEditRedo)
        self.actionEditCut = QXAction('Cu&t',self,triggered=self.onEditCut)
        self.actionEditCopy = QXAction('&Copy',self,triggered=self.onEditCopy)
        self.actionEditPaste = QXAction('&Paste',self,triggered=self.onEditPaste)
        
        #toolbar
        
        if hasToolBar:
            self.tbrMain = self.addToolBar("Main")
            
            self.tbrMain.addAction(self.actionFileNew)
            self.tbrMain.addAction(self.actionFileOpen)
            self.tbrMain.addAction(self.actionFileSave)
            
            self.tbrMain.addSeparator()
            
            self.tbrMain.addAction(self.actionEditUndo)
            self.tbrMain.addAction(self.actionEditRedo)            
            self.tbrMain.addSeparator()
            
            self.tbrMain.addAction(self.actionEditCut)
            self.tbrMain.addAction(self.actionEditCopy)
            self.tbrMain.addAction(self.actionEditPaste)
            
            
        self.setUnifiedTitleAndToolBarOnMac(True)        
        
    def setFileSaveAsSuffix(self,s):
        self._fileSaveAsSuffix = s
        
    def setFileCreateByMe(self,t):
        self._fileCreateByMe = t
        
    def fileCreateByMe(self):
        return self._fileCreateByMe
        
    def setFileName(self,f):
        self._fileName = f;
    
    def setFileReadOnly(self,t):
        self._fileReadOnly = t
        
    def loadFinished(self,success = False):
        
        if not success:
            self.setFileName(None)
            self.setWindowTitle("Untitled[*] - %s" % self.appName)
            return
        
        if not self._fileReadOnly:
            self.setWindowTitle("%s[*] - %s" % (self._fileName,self.appName))
        else:
            self.setWindowTitle("%s (Read Only) - %s" % (self._fileName,self.appName))
            
    def fileName(self):
        return self._fileName
    
    def getAppExecutable(self):
        appExec = QCoreApplication.applicationFilePath()
        appFile = sys.argv[0]
        #print platform.system(),appExec,appFile
        if platform.system() == 'Darwin':
            appExecDir = QDir(appExec)
            if appExecDir.dirName().toLower() == 'python':
                # The script is executed by python
                return "\"%s\" \"%s\"" % (appExec,appFile)
            else:
                return "\"%s\"" %appExec
        elif platform.system() == 'Windows':
            appExecDir = QDir(appExec)
            if appExecDir.dirName().toLower() == 'python.exe' or appExecDir.dirName().toLower() == 'pythonw.exe':
                # The script is executed by python
                return "\"%s\" \"%s\"" % (appExec,appFile)
            else:
                return "\"%s\"" %appExec    
    
        
    def ActionFileNew(self):
        execStr = self.getAppExecutable()
        QProcess.startDetached(execStr)

    def ActionFileOpen(self):
        fileName = QFileDialog.getOpenFileName(self,"Open",QDir.currentPath(),self._fileSaveAsSuffix)
        if not fileName is None and fileName != '':
            self.setFileCreateByMe(False)
            self.ActionFileLoad(fileName)
    
    def ActionFileLoad(self,f):
        self.updateStatusBarMessage("Loading %s" % f)
        self.setFileName(f)
        self.t = QTimer()
        self.t.setSingleShot(True)
        self.t.timeout.connect(self.onFileLoad)
        self.t.start(100)

    def ActionFileSave(self):
        if self._fileName is None or self._fileReadOnly:
            self.setFileCreateByMe(True)
            self.ActionFileSaveAs()
        else:
            self.onFileSave(self.fileName())
                
            
    def ActionFileSaveAs(self):
        fileName = QFileDialog.getSaveFileName(self,"Save As",QDir.currentPath(),self._fileSaveAsSuffix)
        if not fileName is None and fileName != '':
            if self.onFileSaveAs(fileName):            
                self.ActionFileLoad(fileName)

    
        

    def onFileLoad(self):
        pass

    def onFileSaveAs(self,fileName):
        return False
    
    def onFileSave(self,fileName):
        pass

    def setEditUndoRedoStatus(self,canUndo,canRedo):
        self.actionEditUndo.setEnabled(canUndo)
        self.actionEditRedo.setEnabled(canRedo)

    def onEditUndo(self):
        pass
    
    def onEditRedo(self):
        pass
    
    def onEditCut(self):
        pass
    
    def onEditCopy(self):
        pass
    
    def onEditPaste(self):
        pass
    
    def updateStatusBarMessage(self,s):
        self.statusBar().showMessage(s)
    
    def closeEvent(self,e):
        if (self.isWindowModified()):
            msgBox = QMessageBox(self)
            msgBox.setIcon(QMessageBox.Question)
            msgBox.setWindowTitle(self.appName)
            msgBox.setText("The document %s has been modified." % self.fileName())
            msgBox.setInformativeText("Do you want to save your changes?")
            msgBox.setStandardButtons(QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
            msgBox.setDefaultButton(QMessageBox.Save)
            ret = msgBox.exec_()
            
            if ret == QMessageBox.Cancel:
                e.ignore()
                return
            
            if ret == QMessageBox.Save:
                self.ActionFileSave()
                e.accept()
                return            
        e.accept()
        return                  
            
            
if __name__ == '__main__':
    print dir(QKeySequence)
                