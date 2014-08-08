#!/usr/bin/python

from PyQt4.QtCore import *
from PyQt4.QtGui import *

import sys
import re

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

class QXSingleDocMainWindow(QMainWindow):
    
    def __init__(self,parent=None):
        super(QMainWindow, self).__init__(parent)
        
        

    def initDefaultUI(self,hasToolBar = True,hasMenuBar = False):
        d = QApplication.desktop()
        screenWidget = d.screen(d.primaryScreen())
        w = screenWidget.size().height()
        h = screenWidget.size().height() * 0.6
        self.resize(w,h)
        self.move(QPoint((screenWidget.size().width() - w) / 2 , (screenWidget.size().height() - h) / 2 ))
        
        self.appName = re.sub(r'^.*\/','',sys.argv[0])
        self.appName = re.sub(r'\..*$','',self.appName)
        
        self.setWindowTitle("Untitled[*] - %s" % self.appName)   
        self.setFileSaveAsSuffix("All Files (*.*)")
        self.setFileReadOnly(False)
        self._fileName = None
        
        #toolbar
        
        if hasToolBar:
            self.tbrMain = self.addToolBar("Main")
            
        
        self.setUnifiedTitleAndToolBarOnMac(True)
        
        
    def setFileSaveAsSuffix(s):
        self._fileSaveAsSuffix = s
        
    def setFileName(self,f):
        self._fileName = f;
    
    def setFileReadOnly(self,t):
        self._fileReadOnly = t
        
    def loadFinished(self):
        if self._fileReadOnly:
            self.setWindowTitle("%s[*] - %s" % (self._fileName,self.appName))
        else:
            self.setWindowTitle("%s (Read Only) - %s" % (self._fileName,self.appName))
            
    def fileName(self):
        return self._fileName
    
    def ActionFileOpen(self,f):
        self.updateStatusBarMessage("Loading %s" % f)
        self.setFileName(f)
        self.t = QTimer()
        self.t.setSingleShot(True)
        self.t.timeout.connect(self.onFileOpen)
        self.t.start(100)

    def ActionFileSave(self):
        if self._fileName is None:
            self.ActionFileSaveAs()
            
    def ActionFileSaveAs(self):
        fileName = QFileDialog.getSaveFileName(self,"Save As",QDir.currentPath(),self._fileSaveAsSuffix)
        if not fileName is None:
            self.onFileSaveAs(fileName)

    def onFileOpen(self):
        pass

    def onFileSaveAs(self,fileName):
        pass
    
    
    def updateStatusBarMessage(self,s):
        self.statusBar().showMessage(s)
    


if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    x = QXInputDialog.getMulti(None,"multi","Enter Survey",["Name","Gender","Opt*"])
    print x 
                