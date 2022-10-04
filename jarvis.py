from cProfile import label
from re import S
from PyQt5 import QtCore , QtWidgets , QtGui
from PyQt5.QtGui import QMovie
from PyQt5.QtGui import *
from PyQt5.QtWidgets import*
from PyQt5.QtCore import *
from PyQt5.QtCore import Qt , QTimer , QTime , QDate
from jarvisUi import Ui_MainWindow
import sys
import Main

class MainThread(QThread):
    
    def __init__(self):
        super(MainThread, self).__init__()
        
    def run(self):
        self.Task_Gui()
    
    
    def Task_Gui(self):
        Main.TaskExe()
        
startFunctions = MainThread()

class Gui_Start(QMainWindow):
    def __init__(self):
        
        super().__init__()
        
        self.jarvis_ui = Ui_MainWindow()
        self.jarvis_ui.setupUi(self)
        self.jarvis_ui.pushButton.clicked.connect(self.startFunc)
        self.jarvis_ui.pushButton_2.clicked.connect(self.close)
         
         
    def startFunc(self):
        self.jarvis_ui.movies = QtGui.QMovie("GUI Material/B.G/Iron_Template_1.gif")
        self.jarvis_ui.Gif.setMovie(self.jarvis_ui.movies)
        self.jarvis_ui.movies.start()
        
        
        self.jarvis_ui.movies_2 = QtGui.QMovie("GUI Material/VoiceReg/jarvis_jj.gif")
        self.jarvis_ui.Gif_2.setMovie(self.jarvis_ui.movies_2)
        self.jarvis_ui.movies_2.start()
        
        
        self.jarvis_ui.movies_3 = QtGui.QMovie("GUI Material/ExtraGui/initial.gif")
        self.jarvis_ui.Gif_3.setMovie(self.jarvis_ui.movies_2)
        self.jarvis_ui.movies_3.start()
        
        
        timer = QTimer(self)
        timer.timeout.connect(self.showtime)
        timer.start(1000)
        
        startFunctions.start()
    
    def showtime(self):
        
        current_time =QTime.currentTime()
        label_time = current_time.toString("hh:mm:ss")
        label = "Time is Now : " + label_time
        self.jarvis_ui.textBrowser.selfText(label)
       
Gui_App = QApplication(sys.argv)
Gui_Jarvis = Gui_Start()

Gui_Jarvis.show()
exit(Gui_App.exec_())
        