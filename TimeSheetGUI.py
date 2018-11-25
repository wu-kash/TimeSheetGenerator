# -*- coding: utf-8 -*-
"""
Created on Thu Nov 22 22:13:58 2018

@author: Lukasz
"""
import pickle
import os
import time
import sys
from pyqtgraph.Qt import QtGui, QtCore
from PyQt5 import uic
from calendar import monthrange
from openpyxl import Workbook
import datetime


from TimeSheetPatchBorders import patchSheetBorder, fillCellColour

timeSheetGUI = "TimeSheetGUI.ui"

Ui_MainWindow, QtBaseClass = uic.loadUiType(timeSheetGUI)

class TimeSheetApp(QtGui.QMainWindow, Ui_MainWindow):
    
    
    pathLocation = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
    
    employeeList = []
    printList = []
    
    monthDictionary = {'January': 1, 'February': 2, 'March': 3, 'April': 4,
                    'May': 5, 'June': 6, 'July': 7, 'August': 8, 
                    'September': 9, 'October': 10, 'November': 11, 'December': 12 }
    
    dayDictionary = {0 : 'Monday', 1 : 'Tuesday', 2 : 'Wednesday', 3 : 'Thursday',
                     4 : 'Friday', 5 : 'Saturday', 6 : 'Sunday'}
    
    monthSet = ""
    yearSet = 0
    monthDays = 0
    monthStartDayIndex = 0
    monthStartDay = ""
    updateSheet = True
    employee = [] 
    
    sheet = ""
    workBook = ""
    currentDate = []
        
    def __init__(self, excelWorkBook):
        QtGui.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        
        'Setup UI'
        self.setupUi(self)
        
        self.workBook = excelWorkBook
        self.sheet = excelWorkBook.active
        patchSheetBorder(self.sheet)
        
        'Connect buttons to functions'
        self.addButton.clicked.connect(self.addPerson)
        self.editButton.clicked.connect(self.editPerson)
        self.deleteButton.clicked.connect(self.deletePerson)
        self.printButton.clicked.connect(self.printSheets)
        self.quitButton.clicked.connect(self.quitApp)

        'Populate month combo box with months'
        for month, index in self.monthDictionary.items():
            self.monthComboBox.addItem(month)
        
        'Populate listwidget with all employees in saved text file "Employees.txt" '
        self.loadInformation()
        
        'Update dates to current month'
        self.getCurrentDate()
        self.monthComboBox.setCurrentIndex(self.currentDate[0])
        self.yearSpinBox.setValue(self.currentDate[1])
        
        
    def getCurrentDate(self):
        month = datetime.datetime.today().month - 1
        year = datetime.datetime.today().year

        self.currentDate = [ month, year ]
        
    def updateExcel(self):
        
        
        'Get dates set'
        self.monthSet = self.monthComboBox.currentText()
        self.yearSet = self.yearSpinBox.value()
        self.monthStartDayIndex, self.monthDays = monthrange(self.yearSet, int(self.monthDictionary[self.monthSet]))
        
        'Index for week days'
        count = self.monthStartDayIndex
        
        for index, day in self.dayDictionary.items():
            if int(self.monthStartDayIndex) == index:
                self.monthStartDay = day
        
        'Change month and year on sheet'
        self.sheet['C5'] = str(self.monthSet) + " " + str(self.yearSet)
        
        'Clear previous dates, day number and day name'
        for i in range(1, 32):
            self.sheet['B'+str(i+8)] = ""
            self.sheet['C'+ str(i+8)] = ""
        
        'Clear grey filled cells'
        for i in range(1, 32):
            fillCellColour('B'+ str(i+8), 'S'+ str(i+8), 'whiteFill', self.sheet)
        
        'Fill new day numbers and names and fill weekend days grey'
        for day in range(1, self.monthDays + 1):
            self.sheet['B'+ str(day+8)] = str(day)
            self.sheet['C'+ str(day+8)] = self.dayDictionary[count]
            
            if self.dayDictionary[count] == 'Sunday' or self.dayDictionary[count] == 'Saturday':
                fillCellColour('B'+ str(day+8), 'S'+ str(day+8), 'greyFill', self.sheet)
            
            count = count + 1  #Loop to go from first day of month to Sunday, and then reset loop
            if count == 7:
                count = 0
        
        'Clear information cells'
        self.sheet['C2'] = self.employee[4] #Company
        self.sheet['F3'] = self.employee[0] #Name
        self.sheet['F4'] = self.employee[1] #Surname
        self.sheet['N3'] = self.employee[2] #ID
        self.sheet['N4'] = self.employee[3] #Contact
        
        self.workBook.save('TimeSheetApp.xlsx')
        print("Updated workbook")
        
        
    def addPerson(self):
        newEmployee = CreateEmployee()
        if newEmployee.exec_():
            self.employeeList.append(newEmployee.employeeDetails)
            self.updateEmployeeList() #Update employee list with new person
            
    def editPerson(self):
        print(self.listWidget.currentItem().text())
        
        
        employeeName = self.listWidget.currentItem().text()
        employeeInfo = [] 
        'Find selected employee by name in employee list'
        for i in self.employeeList:
            if i[0] == employeeName.split()[0]:
                employeeInfo = i
                print(i)
        
        editEmployee = CreateEmployee()
        editEmployee.changeButton()
        editEmployee.setTextBox(employeeInfo)
        if editEmployee.exec_():
            self.employeeList.remove(employeeInfo)
            self.employeeList.append(editEmployee.employeeDetails)
            self.updateEmployeeList()
            
        print("Edit Person")
        
        self.updateEmployeeList()
    
    def deletePerson(self):
        print(self.listWidget.currentItem().text())
        
        employeeName = self.listWidget.currentItem().text()
        employeeInfo = [] 
        'Find selected employee by name in employee list'
        for i in self.employeeList:
            if i[0] == employeeName.split()[0]:
                employeeInfo = i
                print(i)
        

        self.employeeList.remove(employeeInfo)
        self.updateEmployeeList()
    
    def printSheets(self):     
        self.printList = []
        for i in range(self.listWidget.count()):
                
            check_box = self.listWidget.item(i)
            state = check_box.checkState()
            
            if int(state) == 2:
                self.employeeList[i][5] = True
                self.printList.append(self.employeeList[i])
            else:
                self.employeeList[i][5] = False
        
        
        print("Printing List")
        print(self.printList)
        print("----------------------")
        
        copies = self.copyCountSpinBox.value()
        
        for j in range(1, copies + 1):
            print("Copy " + str(j) + " of " + str(copies))
            for i in self.printList:
                print("Printing for:", i[0], i[1])
                self.employee = i
                self.updateExcel()
                print("Printing... ")
                os.startfile( str(self.pathLocation) + '\\TimeSheetApp.xlsx', 'print')
            
            
                time.sleep(8)
            
            print("Print " + str(j) + " of " + str(copies) + " complete")
   

     
    def loadInformation(self):
    
        'try load the pickled data stored in EmployeeList.txt and store in list emplooyeeList'
        with open(os.path.join(self.pathLocation, 'EmployeeList.txt'), 'rb') as picklefile:
            try:
                self.employeeList = pickle.load(picklefile)
                self.updateEmployeeList()
            except EOFError:
                print("EOF")
                self.employeeList = []
        
    
    def updateEmployeeList(self):
        
        'clear listwidget with the employees'
        self.listWidget.clear()
        
        're add all the employees with new added information'
        for i in self.employeeList:
            item = QtGui.QListWidgetItem()
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
            item.setCheckState(QtCore.Qt.Unchecked)
            info = str(i[0]) + " " + str(i[1])
            item.setText(info)
            self.listWidget.addItem(item)  
        
        print("Employee List Information:")
        for i in self.employeeList:
            print(i)
                    
        
    def saveInformation(self):
    
        'Reset all print states back to False'
        for i in self.employeeList:
            i[5] = False
        
        print("Saving..")
        with open(os.path.join(self.pathLocation, 'EmployeeList.txt'), 'wb') as picklefile:
            pickle.dump(self.employeeList, picklefile)
        
    def quitApp(self):
        self.saveInformation()
        sys.exit()
    

createEmployeeGUI = "CreateEmployeeGUI.ui"

Ui_MainWindow, QtBaseClass = uic.loadUiType(createEmployeeGUI)

class CreateEmployee(QtGui.QDialog, Ui_MainWindow):
    
    employeeDetails = []
    name = ""
    surname = ""
    IDNumber = ""
    contact = ""
    company = ""
    status = False

    def __init__(self):
        QtGui.QDialog.__init__(self)
        Ui_MainWindow.__init__(self)
        
        'Setup UI'
        self.setupUi(self)
        
        'Connect buttons to functions'
        self.createButton.clicked.connect(self.accept)
        self.createButton.clicked.connect(self.getInformation)
        
        self.cancelButton.clicked.connect(self.reject)
        
        self.employeeDetails = []
        self.name = ""
        self.surname = ""
        self.IDNumber = ""
        self.contact = ""
        self.company = ""
        self.status = False
        
    def getInformation(self):
        self.name = self.nameTextBox.toPlainText()
        self.surname = self.surnameTextBox.toPlainText()
        self.IDNumber = self.IDTextBox.toPlainText()
        self.contact = self.contactTextBox.toPlainText()
        self.company = self.companyTextBox.toPlainText()
        
        self.employeeDetails.append(self.name)
        self.employeeDetails.append(self.surname)
        self.employeeDetails.append(self.IDNumber)
        self.employeeDetails.append(self.contact)
        self.employeeDetails.append(self.company)
        self.employeeDetails.append(self.status)    #Status for whether employee is checked on list
        print("New Employee:")
        print(self.employeeDetails)
        self.close()
        
    def setTextBox(self, employee):
        self.nameTextBox.insertPlainText(employee[0])
        self.surnameTextBox.insertPlainText(employee[1])
        self.IDTextBox.insertPlainText(employee[2])
        self.contactTextBox.insertPlainText(employee[3])
        self.companyTextBox.insertPlainText(employee[4])
        
    def changeButton(self):
        self.createButton.setText("Edit")
        