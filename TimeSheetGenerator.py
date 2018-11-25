# -*- coding: utf-8 -*-
"""
Created on Wed Nov 21 17:17:30 2018

@author: Lukasz

TimeSheet Generator

"""



from openpyxl import Workbook
import openpyxl


from TimeSheetGUI import TimeSheetApp

from pyqtgraph.Qt import QtGui
import sys


'''Excel related -----------------------------------------------------------'''
timeSheet = Workbook('TimeSheetApp.xlsx')
timeSheet = openpyxl.load_workbook('TimeSheetApp.xlsx')
'''Excel END ------------------------------5---------------------------------'''

if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    window = TimeSheetApp(timeSheet)
    window.show()
    
    
    sys.exit(app.exec_())
