# -*- coding: utf-8 -*-
"""
Created on Fri Nov 23 10:16:42 2018

@author: Lukasz
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Nov 22 22:13:58 2018

@author: Lukasz
"""

import sys
from pyqtgraph.Qt import QtGui, QtCore
from PyQt5.QtWidgets import QDialog, QApplication, QPushButton, QVBoxLayout
from PyQt5 import uic

qtCreatorFile = "CreateEmployeeGUI.ui"

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

class CreateEmployee(QtGui.QMainWindow, Ui_MainWindow):
    
    
    def __init__(self):
        QtGui.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        
        'Setup UI'
        self.setupUi(self)
        
        'Connect buttons to functions'
        self.cancelButton.clicked.connect(self.cancel)
        
        
    def cancel(self):
        self.close()
        
        
        