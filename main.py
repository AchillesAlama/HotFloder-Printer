from asyncio.windows_events import NULL
from asyncore import loop
import win32api, win32print, os, sys, time, shutil, re, json, threading
from datetime import datetime
from mainGUI import Ui_MainWindow
from PyQt5.QtWidgets import QDialog, QMainWindow, QFileDialog
from PyQt5.QtGui import QIcon

def loop(self):
    while True:
        files = self.checkForValidFiles()
        if len(files) > 0:
            time.sleep(3)
            files = self.checkForValidFiles()
            for f in files:
                try:
                    if self._printerSelected == NULL:
                        self._printerSelected = win32print.GetDefaultPrinter()
                    win32api.ShellExecute(0,"print", os.path.join(self._hotPath,f), None,  ".",  0)
                    shutil.copy(os.path.join(self._hotPath,f),os.path.join(self._archivePath,f))
                    self.deleteFile(f)
                    self.updateGUI(True,f)
                except Exception as e:
                    # print(e)
                    self.updateGUI(False,f)
                    shutil.copy(os.path.join(self._hotPath,f),os.path.join(self._errorPath,f))
                    self.deleteFile(f)


class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowIcon(QIcon("icon.png"))
        self.show()
        self.ui.textBrowser.append(self.getActualTime() + " hi")
        with open('settings.json') as json_file:
            settings = json.load(json_file)
            self._hotPath = settings['HotFolder']
            self._archivePath = settings['Archive']
            self._errorPath = settings['Error']
            self._printerSelected = settings['Printer']
        self.ui.comboBox.highlighted.connect(self.onPrinterChange)
        self.ui.pushButton_2.clicked.connect(self.openArchiveSelector)
        self.ui.pushButton_4.clicked.connect(self.openErrorSelector)
        self.ui.pushButton_3.clicked.connect(self.openHotSelector)
        for item in self.printerlist():
            self.ui.comboBox.addItem(str(item[2]))
        
        # self.loop()
        threading.Thread(target=loop, args=(self,)).start()

    def autosave(self):
        settings = {
            "HotFolder": self._hotPath,
            "Archive": self._archivePath,
            "Error": self._errorPath,
            "Printer": self._printerSelected
        }
        with open('settings.json', 'w+') as json_file:
            json.dump(settings, json_file)
    def openHotSelector(self):
        self._hotPath=QFileDialog.getExistingDirectory(self,"Choose HotFolder Directory","E:\\")
        self.autosave()
    def openArchiveSelector(self):
        self._archivePath=QFileDialog.getExistingDirectory(self,"Choose Archive Directory","E:\\")
        self.autosave()
    def openErrorSelector(self):
        self._errorPath=QFileDialog.getExistingDirectory(self,"Choose Error Directory","E:\\")
        self.autosave()
    def onPrinterChange(self, value):
        l = self.printerlist()
        self._printerSelected = l[value]
        self.autosave()
    def printerlist(self):
        return win32print.EnumPrinters(2)  

    def checkForValidFiles(self):
        files = []
        filesInFolder = os.listdir(self._hotPath)
        for f in filesInFolder:
            if f[-3:] == "pdf":
                files.append(f)
        return files

    def getActualTime(self):
        now = datetime.now()
        t = now.strftime("%d.%m.%Y, %H:%M:%S")
        return t

    def deleteFile(self,f):
        while f in os.listdir(self._hotPath):
            try:              
                time.sleep(2)
                os.remove(os.path.join(self._hotPath,f))
            except Exception as e:
                time.sleep(2)
        return None

    def updateGUI(self,succ,f):
        if succ == True:
            self.ui.textBrowser.append(self.getActualTime()+" File "+f+" printed on " + self._printerSelected[2] + " and moved to archive")
        if succ == False:
            self.ui.textBrowser.append(self.getActualTime()+" File "+f+" faced an error when printing on " + self._printerSelected[2] + " and moved to error folder")