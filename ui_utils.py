#!/usr/bin/python

from PyQt4.QtCore import *
from PyQt4.QtGui import *
import platform
import sys
import re
import ui_utils_res

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
        
class QXSingleDocMainWindow(QMainWindow):
    
    def __init__(self,parent=None):
        super(QMainWindow, self).__init__(parent)
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
        
        self.actionEditUndoDefault = 

        #toolbar
        
        if hasToolBar:
            self.tbrMain = self.addToolBar("Main")
            
            self.tbrMain.addAction(self.actionFileNew)
            self.tbrMain.addAction(self.actionFileOpen)
            self.tbrMain.addAction(self.actionFileSave)

            
            
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
    
    def updateStatusBarMessage(self,s):
        self.statusBar().showMessage(s)
    
    def closeEvent(self,e):
        if (self.isWindowModified()):
            msgBox = QMessageBox(self)
            msgBox.setIcon(QMessageBox.Question)
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
                