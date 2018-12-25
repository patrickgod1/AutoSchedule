#! python3
import os, sys, time, inspect, datetime
#import PyQt5
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QApplication, QWidget, QPushButton, QMessageBox
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
# import docx
# from docx.enum.section import WD_ORIENT
# from docx.shared import Pt, Inches
# from pptx import Presentation
# from pptx.dml.color import RGBColor

if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)

if hasattr(QtCore.Qt, 'AA+_UseHighDpiPixmaps'):
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)   

# Global variables and flags
current_folder = os.path.realpath(os.path.abspath(os.path.split(inspect.getfile(inspect.currentframe() ))[0]))
# ABSWTemplate = os.path.join(current_folder, "Template-ABSW.docx")
# GBCTemplate = os.path.join(current_folder, "Template-GBC.docx")
# GBCPptTemplate = os.path.join(current_folder, "Template-GBC.pptx")
genReport = False
startDate = str(QtCore.QDate.currentDate().toPyDate())
endDate = str(QtCore.QDate.currentDate().toPyDate())
location = 'ABSW'
# createSigns = False
useExistingReport = False
saveReportToPath = ''
# existingReportPath, _ = ('','')
# classroomSignsOutput = False
# dailyScheduleOutput = False
# powerpointOutput = False
ABSWScheduleOutput = False
BelmontScheduleOutput = False
GBCScheduleOutput = False
SFCScheduleOutput = False
# saveSignsDirectory = ''
center = {'ABSW': {'campus': 'Berkeley - CA0001', 'building': 'UC Berkeley Extension American Baptist Seminary of the West, 2515 Hillegass Ave. - '},
          'Belmont': {'campus': 'Belmont - CA0004', 'building': 'UC Berkeley Extension Belmont Center, 1301 Shoreway Rd., Ste. 400 - BEL'},
          'Golden Bear Center': {'campus': 'Berkeley - CA0001', 'building': 'UC Berkeley Extension Golden Bear Center, 1995 University Ave. - GBC'},
          'San Francisco Center': {'campus': 'San Francisco - CA0003', 'building': 'San Francisco Campus, 160 Spear St. - SFCAMPUS'}
            }

centerReverse = {'ABSW - UC Berkeley Extension American Baptist Seminary of the West, 2515 Hillegass Ave.': {'name':'ABSW'},
                 'BEL - UC Berkeley Extension Belmont Center, 1301 Shoreway Rd., Ste. 400': {'name':'BLM'},
                 'GBC - UC Berkeley Extension Golden Bear Center, 1995 University Ave.': {'name': 'GBC'},
                 'SFCAMPUS - San Francisco Campus, 160 Spear St.': {'name': 'SFC'}}

# Main Window for GUI
class Ui_mainWindow(object):
    def setupUi(self, mainWindow):
        mainWindow.setObjectName("mainWindow")
        mainWindow.setWindowModality(QtCore.Qt.NonModal)
        mainWindow.setEnabled(True)
        mainWindow.resize(547, 222)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(mainWindow.sizePolicy().hasHeightForWidth())
        mainWindow.setSizePolicy(sizePolicy)
        mainWindow.setBaseSize(QtCore.QSize(430, 400))
        self.mainWindowLayout = QtWidgets.QVBoxLayout(mainWindow)
        self.mainWindowLayout.setObjectName("mainWindowLayout")
        self.genReportBox = QtWidgets.QGroupBox(mainWindow)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.MinimumExpanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.genReportBox.sizePolicy().hasHeightForWidth())
        self.genReportBox.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.genReportBox.setFont(font)
        self.genReportBox.setCheckable(False)
        self.genReportBox.setChecked(False)
        self.genReportBox.setObjectName("genReportBox")
        self.genReportLayout = QtWidgets.QVBoxLayout(self.genReportBox)
        self.genReportLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.genReportLayout.setObjectName("genReportLayout")
        self.dateLayout = QtWidgets.QHBoxLayout()
        self.dateLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.dateLayout.setObjectName("dateLayout")
        self.startDateLabel = QtWidgets.QLabel(self.genReportBox)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.startDateLabel.sizePolicy().hasHeightForWidth())
        self.startDateLabel.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.startDateLabel.setFont(font)
        self.startDateLabel.setObjectName("startDateLabel")
        self.dateLayout.addWidget(self.startDateLabel)
        self.selectStartDate = QtWidgets.QDateEdit(self.genReportBox)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.selectStartDate.sizePolicy().hasHeightForWidth())
        self.selectStartDate.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.selectStartDate.setFont(font)
        self.selectStartDate.setFrame(True)
        self.selectStartDate.setReadOnly(False)
        self.selectStartDate.setProperty("showGroupSeparator", False)
        self.selectStartDate.setCalendarPopup(True)
        self.selectStartDate.setDate(QtCore.QDate.currentDate())
        startDate = str(QtCore.QDate.currentDate().toPyDate())
        self.selectStartDate.setObjectName("selectStartDate")
        self.dateLayout.addWidget(self.selectStartDate)
        self.endDateLabel = QtWidgets.QLabel(self.genReportBox)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.endDateLabel.sizePolicy().hasHeightForWidth())
        self.endDateLabel.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.endDateLabel.setFont(font)
        self.endDateLabel.setObjectName("endDateLabel")
        self.dateLayout.addWidget(self.endDateLabel)
        self.selectEndDate = QtWidgets.QDateEdit(self.genReportBox)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.selectEndDate.sizePolicy().hasHeightForWidth())
        self.selectEndDate.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.selectEndDate.setFont(font)
        self.selectEndDate.setFocusPolicy(QtCore.Qt.WheelFocus)
        self.selectEndDate.setReadOnly(False)
        self.selectEndDate.setCalendarPopup(True)
        self.selectEndDate.setDate(QtCore.QDate.currentDate())
        endDate = str(QtCore.QDate.currentDate().toPyDate())
        self.selectEndDate.setObjectName("selectEndDate")
        self.dateLayout.addWidget(self.selectEndDate)
        self.genReportLayout.addLayout(self.dateLayout)
        self.saveReportPathLayout = QtWidgets.QHBoxLayout()
        self.saveReportPathLayout.setObjectName("saveReportPathLayout")
        self.saveReportPathLabel = QtWidgets.QLabel(self.genReportBox)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.saveReportPathLabel.sizePolicy().hasHeightForWidth())
        self.saveReportPathLabel.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.saveReportPathLabel.setFont(font)
        self.saveReportPathLabel.setObjectName("saveReportPathLabel")
        self.saveReportPathLayout.addWidget(self.saveReportPathLabel)
        self.selectSaveReportPath = QtWidgets.QLineEdit(self.genReportBox)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.selectSaveReportPath.sizePolicy().hasHeightForWidth())
        self.selectSaveReportPath.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.selectSaveReportPath.setFont(font)
        self.selectSaveReportPath.setReadOnly(True)
        self.selectSaveReportPath.setObjectName("selectSaveReportPath")
        self.saveReportPathLayout.addWidget(self.selectSaveReportPath)
        self.browseSaveReportButton = QtWidgets.QToolButton(self.genReportBox)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.browseSaveReportButton.sizePolicy().hasHeightForWidth())
        self.browseSaveReportButton.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        self.browseSaveReportButton.setFont(font)
        self.browseSaveReportButton.setCheckable(False)
        self.browseSaveReportButton.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        self.browseSaveReportButton.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
        self.browseSaveReportButton.setObjectName("browseSaveReportButton")
        self.saveReportPathLayout.addWidget(self.browseSaveReportButton)
        self.genReportLayout.addLayout(self.saveReportPathLayout)
        self.locationLayout = QtWidgets.QHBoxLayout()
        self.locationLayout.setObjectName("locationLayout")
        self.ABSWcheckBox = QtWidgets.QCheckBox(self.genReportBox)
        font = QtGui.QFont()
        font.setUnderline(False)
        self.ABSWcheckBox.setFont(font)
        self.ABSWcheckBox.setObjectName("ABSWcheckBox")
        self.locationLayout.addWidget(self.ABSWcheckBox)
        self.BelmontcheckBox = QtWidgets.QCheckBox(self.genReportBox)
        font = QtGui.QFont()
        font.setUnderline(False)
        self.BelmontcheckBox.setFont(font)
        self.BelmontcheckBox.setObjectName("BelmontcheckBox")
        self.locationLayout.addWidget(self.BelmontcheckBox)
        self.GBCcheckBox = QtWidgets.QCheckBox(self.genReportBox)
        font = QtGui.QFont()
        font.setUnderline(False)
        self.GBCcheckBox.setFont(font)
        self.GBCcheckBox.setObjectName("GBCcheckBox")
        self.locationLayout.addWidget(self.GBCcheckBox)
        self.SFCcheckBox = QtWidgets.QCheckBox(self.genReportBox)
        font = QtGui.QFont()
        font.setUnderline(False)
        self.SFCcheckBox.setFont(font)
        self.SFCcheckBox.setObjectName("SFCcheckBox")
        self.locationLayout.addWidget(self.SFCcheckBox)
        self.genReportLayout.addLayout(self.locationLayout)
        self.mainWindowLayout.addWidget(self.genReportBox)
        self.startExitLayout = QtWidgets.QHBoxLayout()
        self.startExitLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.startExitLayout.setObjectName("startExitLayout")
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.startExitLayout.addItem(spacerItem)
        self.StartButton = QtWidgets.QPushButton(mainWindow)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.StartButton.sizePolicy().hasHeightForWidth())
        self.StartButton.setSizePolicy(sizePolicy)
        self.StartButton.setMinimumSize(QtCore.QSize(175, 50))
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setWeight(75)
        self.StartButton.setFont(font)
        self.StartButton.setObjectName("StartButton")
        self.startExitLayout.addWidget(self.StartButton)
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Minimum)
        self.startExitLayout.addItem(spacerItem1)
        self.exitButton = QtWidgets.QPushButton(mainWindow)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.exitButton.sizePolicy().hasHeightForWidth())
        self.exitButton.setSizePolicy(sizePolicy)
        self.exitButton.setMinimumSize(QtCore.QSize(175, 50))
        font = QtGui.QFont()
        font.setFamily("Tahoma")
        font.setPointSize(14)
        font.setBold(True)
        font.setItalic(False)
        font.setWeight(75)
        self.exitButton.setFont(font)
        self.exitButton.setObjectName("exitButton")
        self.startExitLayout.addWidget(self.exitButton)
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.startExitLayout.addItem(spacerItem2)
        self.mainWindowLayout.addLayout(self.startExitLayout)

        self.retranslateUi(mainWindow)
        QtCore.QMetaObject.connectSlotsByName(mainWindow)

    def retranslateUi(self, mainWindow):
        _translate = QtCore.QCoreApplication.translate
        mainWindow.setWindowTitle(_translate("mainWindow", "Auto Schedule"))
        self.genReportBox.setTitle(_translate("mainWindow", "Generate Destiny Report"))
        self.startDateLabel.setText(_translate("mainWindow", "Start Date:"))
        self.endDateLabel.setText(_translate("mainWindow", "End Date:"))
        self.saveReportPathLabel.setText(_translate("mainWindow", "Save Path:"))
        self.selectSaveReportPath.setText(_translate("mainWindow", ""))
        self.browseSaveReportButton.setText(_translate("mainWindow", "Browse"))
        self.ABSWcheckBox.setText(_translate("mainWindow", "ABSW"))
        self.BelmontcheckBox.setText(_translate("mainWindow", "Belmont"))
        self.GBCcheckBox.setText(_translate("mainWindow", "GBC"))
        self.SFCcheckBox.setText(_translate("mainWindow", "SFC"))
        self.StartButton.setText(_translate("mainWindow", "Start"))
        self.exitButton.setText(_translate("mainWindow", "Exit"))


        #self.genReportBox.toggled.connect(self.genReportState)
        self.selectStartDate.dateChanged.connect(self.startDateChanged)
        self.selectEndDate.dateChanged.connect(self.endDateChanged)
        # self.selectLocation.currentIndexChanged.connect(self.locationChanged)
        self.browseSaveReportButton.clicked.connect(self.saveReportDirectory)
        # self.createSignsBox.toggled.connect(self.createSignsState)
        # self.useExistingReportBox.toggled.connect(self.useExistingReportState)
        # self.browseExistingReportButton.clicked.connect(self.existingReportPath)
        self.ABSWcheckBox.toggled.connect(self.ABSWcheckBoxState)
        self.BelmontcheckBox.toggled.connect(self.BelmontcheckBoxState)
        self.GBCcheckBox.toggled.connect(self.GBCcheckBoxState)
        self.SFCcheckBox.toggled.connect(self.SFCcheckBoxState)
        # self.browseSaveSignsButton.clicked.connect(self.saveSignsPath)
        self.exitButton.clicked.connect(self.exitApp)
        self.StartButton.clicked.connect(self.startApp)

    # def genReportState(self):
    #     global genReport, useExistingReport
    #     if self.genReportBox.isChecked():
    #         genReport = True
    #         useExistingReport = False
    #         self.useExistingReportBox.setChecked(False)
    #     else:   
    #         genReport = False
    #         useExistingReport = True
    #         self.useExistingReportBox.setEnabled(True)
    #         self.useExistingReportBox.setChecked(True)

    def startDateChanged(self):
        global startDate
        startDate = str(self.selectStartDate.date().toPyDate())

    def endDateChanged(self):
        global endDate
        endDate = str(self.selectEndDate.date().toPyDate())

    def locationChanged(self, i):
        global location
        location = self.selectLocation.currentText()

    def saveReportDirectory(self):
        global saveReportToPath
        saveReportToPath = QFileDialog.getExistingDirectory(None, 'Save Destiny Report to')
        self.selectSaveReportPath.setText(saveReportToPath)

    def ABSWcheckBoxState(self):
        global ABSWScheduleOutput
        if self.ABSWcheckBox.isChecked():
            ABSWScheduleOutput = True
        else:   
            ABSWScheduleOutput = False

    def BelmontcheckBoxState(self):
        global BelmontScheduleOutput
        if self.BelmontcheckBox.isChecked():
            BelmontScheduleOutput = True
        else:   
            BelmontScheduleOutput = False

    def GBCcheckBoxState(self):
        global GBCScheduleOutput
        if self.GBCcheckBox.isChecked():
            GBCScheduleOutput = True
        else:   
            GBCScheduleOutput = False

    def SFCcheckBoxState(self):
        global SFCScheduleOutput
        if self.SFCcheckBox.isChecked():
            SFCScheduleOutput = True
        else:   
            SFCScheduleOutput = False       

    def exitApp(self):        
        reply = QMessageBox.question(None, 'Exit', "Are you sure you want to exit?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)                                      
        if reply == QMessageBox.Yes:
            sys.exit()
        else:
            pass

    def startApp(self):
        # Invalid input checks
        result = 1
        if endDate < startDate:
            QMessageBox.warning(None, 'Invalid date range', "Please select a valid date range.")
            return
        elif saveReportToPath == '':
            QMessageBox.warning(None, 'Save location error', "Please select where you want to save the report to.")
            return
        elif not (ABSWScheduleOutput or BelmontScheduleOutput or GBCScheduleOutput or SFCScheduleOutput):
            QMessageBox.warning(None, 'Output Error', "Please a location.")
            return
        else:
            if os.path.isdir(saveReportToPath):
                result = genReportFunction()
            else:  
                QMessageBox.warning(None, 'Save location error', "The directory you've selected does not exist. Please select where you want to save the report to.")
                return
        if result == 0:
            QMessageBox.warning(None, 'Error!!!', "An unexpected error occured.")
        else:
            QMessageBox.warning(None, 'Done', "Done creating signs.")
        
                
def genReportFunction():
    # Set Chrome defaults to automate download
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": saveReportToPath,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.endabled": True
        })

    locationList = []
    if ABSWScheduleOutput:
        locationList.append('ABSW')
    if BelmontScheduleOutput:
        locationList.append('Belmont')
    if GBCScheduleOutput:
        locationList.append('Golden Bear Center')
    if SFCScheduleOutput:
        locationList.append('San Francisco Center')
        
    # Delete old report if it exists
    if os.path.exists(f"{saveReportToPath}/SectionScheduleDailySummary.xls"):
        os.remove(f"{saveReportToPath}/SectionScheduleDailySummary.xls")
    if os.path.exists(f"{saveReportToPath}/SectionScheduleDailySummary (1).xls"):
        os.remove(f"{saveReportToPath}/SectionScheduleDailySummary (1).xls") 
    if os.path.exists(f"{saveReportToPath}/SectionScheduleDailySummary (2).xls"):
        os.remove(f"{saveReportToPath}/SectionScheduleDailySummary (2).xls")
    if os.path.exists(f"{saveReportToPath}/SectionScheduleDailySummary (3).xls"):
        os.remove(f"{saveReportToPath}/SectionScheduleDailySummary (3).xls")   

    reportPath = []    
    chromedriver = os.path.join(current_folder,"chromedriver.exe")
    browser = webdriver.Chrome(executable_path = chromedriver, chrome_options=chrome_options)
    browser.get('https://berkeleysv.destinysolutions.com')
    WebDriverWait(browser,3600).until(EC.presence_of_element_located((By.ID,"main-area-body")))
    for i in range(len(locationList)):
    # Download Destiny Report
        browser.get('https://berkeleysv.destinysolutions.com/srs/reporting/sectionScheduleDailySummary.do?method=load')
        startDateElm = browser.find_element_by_id('startDateRecordString')
        startDateElm.send_keys(startDate)
        endDateElm = browser.find_element_by_id('endDateRecordString')
        endDateElm.send_keys(endDate)
        campusElm = browser.find_element_by_name('scheduleBlock.campusId')
        campusElm.send_keys(center[locationList[i]]['campus'])
        buildingElm = browser.find_element_by_name('scheduleBlock.buildingId')
        buildingElm.send_keys(center[locationList[i]]['building'])
        outputTypeElm = browser.find_element_by_name('outputType')
        outputTypeElm.send_keys("Output to XLS (Export)")
        generateReportElm = browser.find_element_by_id('processReport')
        generateReportElm.click()
        if i == 0:
            while not os.path.exists(f"{saveReportToPath}/SectionScheduleDailySummary.xls"):
                time.sleep(1)
            reportPath.append(f"{saveReportToPath}/SectionScheduleDailySummary.xls")
        elif i == 1:
            while not os.path.exists(f"{saveReportToPath}/SectionScheduleDailySummary (1).xls"):
                time.sleep(1)
            reportPath.append(f"{saveReportToPath}/SectionScheduleDailySummary (1).xls")
        elif i == 2:
            while not os.path.exists(f"{saveReportToPath}/SectionScheduleDailySummary (2).xls"):
                time.sleep(1)
            reportPath.append(f"{saveReportToPath}/SectionScheduleDailySummary (2).xls")
        else:
            while not os.path.exists(f"{saveReportToPath}/SectionScheduleDailySummary (3).xls"):
                time.sleep(1)
            reportPath.append(f"{saveReportToPath}/SectionScheduleDailySummary (3).xls")           
    browser.quit()

    for rp in reportPath:
        createSchedule(rp)
    return 1

def createSchedule(reportPath):
    # Read in courses from Excel
    # 1     B   Date
    # 3     D   Type
    # 4     E   Start Time
    # 6     G   End Time
    # 9     J   Section Number
    # 11    L   Section Title
    # 12    M   Instructor
    # 13    N   Building
    # 15    P   Room
    # 16    Q   Configuration
    # 17    R   Technology
    # 18    S   Section Size
    # 20    U   Notes
    # 22    W   Final Approval
    
    # Read into Pandas dataframe for relevant columns
    schedule = pd.read_excel(reportPath, header=6, skipfooter=1, usecols=[1,15,18,4,6,11,12,9,17,20,13], parse_dates=['Start Time', 'End Time'])
    #print(schedule.head())
    # Determine if the Destiny report does not have any classes
    if schedule.empty:
        print(f"No classes in {reportPath[idx]}")
    # If report is not empty, determine which location report is for and which template to use
    else:
        location = centerReverse[schedule.iloc[0][6]]['name']
        if location == 'SFC':
            SFCSchedule(schedule, location)
        else:
            GBCSchedule(schedule, location)
        

def GBCSchedule(schedule, location):
    sortedSchedule = schedule.sort_values(by=['Date','Start Time', 'Room'])
    sortedSchedule['Start Time'] = sortedSchedule['Start Time'].dt.strftime('%I:%M %p')
    sortedSchedule['End Time'] = sortedSchedule['End Time'].dt.strftime('%I:%M %p')
    sortedSchedule = sortedSchedule.fillna('')
    dateList = sortedSchedule['Date'].astype(datetime.datetime).unique()

    writer = pd.ExcelWriter(f"{saveReportToPath}/{location} Schedule {dateList[0].strftime('%Y-%m-%d')} {dateList[0].strftime('%A')}.xlsx", engine='xlsxwriter')
    workbook = writer.book

    worksheet = workbook.add_worksheet(location)
    worksheet.set_default_row(hide_unused_rows=True)
    worksheet.set_column('K:XFD', None, None, {'hidden': True})
    worksheet.freeze_panes(2, 0)
    worksheet.set_landscape()       # Page orientation as landscape.
    worksheet.hide_gridlines(0)     # Don’t hide gridlines.
    worksheet.fit_to_pages(1, 0)    # Fit to 1x1 pages.
    worksheet.center_horizontally()
    worksheet.center_vertically()
    worksheet.set_paper(1)          # Set paper size to 8.5" x 11"
    worksheet.set_margins(left=0.25, right=0.25, top=0.25, bottom=0.25)
    worksheet.set_header('', {'margin': 0})
    worksheet.set_footer('', {'margin': 0})

    worksheet.set_column('A:A', 29)   # Column A (Date) width
    worksheet.set_column('B:B', 7.29)   # Column B (Room) width
    worksheet.set_column('C:C', 3.86)   # Column C (Size) width
    worksheet.set_column('D:D', 9.29)   # Column D (Start Date) width
    worksheet.set_column('E:E', 8.43)   # Column E (End Date) width
    worksheet.set_column('F:F', 46.29)   # Column F (Section Title) width
    worksheet.set_column('G:G', 12.57)   # Column G (Instructor) width
    worksheet.set_column('H:H', 16)   # Column H (Section Number) width
    worksheet.set_column('I:I', 18.57)   # Column H Technology) width
    worksheet.set_column('J:J', 64)   # Column H (Notes) width

    locationFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        'font_color': '#000000',
        'bg_color': '#FFC000'
        })

    genFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        'font_color': '#C00000',
        'bg_color': '#FDE9D9'
        })

    titleFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        'font_color': '#000000'
        })

    headerFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        #'valign': 'vcenter',
        'bottom': 2,
        'bottom_color': '#000000'
        })

    bodyFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': False,
        'valign': 'top',
        'text_wrap': False,
        'font_color': '#000000'
        })

    daySeparator = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        'font_color': '#000000',
        'bg_color': '#00B050'
        })

    roomFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        'align': 'right',
        'font_color': '#C00000'
        })

    roomFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        'font_color': '#C00000'
        })

    laptopReadyFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'italic': True,
        'text_wrap': False,
        'font_color': '#0070C0'
        })

    instructorFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': False,
        'text_wrap': False,
        'font_color': '#C00000'
        })

    # AM
    amFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': False,
        'text_wrap': False,
        'font_color': '#FF0000',
        'bg_color': '#DAEEF3'
        })

    # Computer Lab
    labFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        'font_color': '#FF0000',
        'bg_color': '#FFFF00'
        })

    worksheet.write(0, 0, f"{sortedSchedule.iloc[0][6]}", locationFormat)
    worksheet.write(0, 1, "", locationFormat)
    worksheet.write(0, 2, "", locationFormat)
    worksheet.write(0, 3, "", locationFormat)
    worksheet.write(0, 4, "", locationFormat)
    
    worksheet.write(0, 5, f"Report generated as of {dateList[0].strftime('%A')}, {dateList[0].strftime('%B %d, %Y').replace(' 0', ' ')}", genFormat)
    for col_num, value in enumerate(['Date', 'Room', 'Size', 'Start Time','End Time','Section Title', 'Instructor', 'Section Number','Technology','Notes']):
        worksheet.write(1, col_num, value, headerFormat)
    excelRow = 2
    # loop through each day
    for i in range(0,len(dateList)):
        singleDaySched = sortedSchedule.loc[sortedSchedule['Date'] == dateList[i], : ]
        morningBlock = singleDaySched.loc[singleDaySched['Start Time'].astype('datetime64') < '12:00:00', : ]
        afternoonBlock = singleDaySched.loc[(singleDaySched['Start Time'].astype('datetime64') >= '12:00:00') & (singleDaySched['Start Time'].astype('datetime64') < '17:00:00'), : ]
        eveningBlock = singleDaySched.loc[singleDaySched['Start Time'].astype('datetime64') >= '17:00:00', : ]
        
        #excelRow +=1
        worksheet.write(excelRow, 0, f"{dateList[i].strftime('%A')}, {dateList[i].strftime('%B %d, %Y').replace(' 0', ' ')}", daySeparator)
        worksheet.write_number(excelRow, 3, len(singleDaySched.index), amFormat)
        excelRow +=1
        if not morningBlock.empty:
            # worksheet.write(excelRow, 0, 'Morning Classes', titleFormat)
            for idx, row in morningBlock.iterrows():
                worksheet.write(excelRow, 0, f"{dateList[i].strftime('%A')}, {dateList[i].strftime('%B %d, %Y').replace(' 0', ' ')}", bodyFormat)
                if location == "GBC" and row['Room'] == "Classroom 201":
                # if row['Room'].replace("Classroom ", "") in ["201", "502", "510", "514", "515"]:
                    worksheet.write_number(excelRow, 1, int(row['Room'].replace("Classroom ", "")), labFormat)
                else:
                    if "Conference Room" in row['Room'] or "Conference Room" in row['Room']:
                        worksheet.write(excelRow, 1, row['Room'].replace("Conference Room ", "CR"), roomFormat)
                    elif row['Room'].replace("Classroom ", "").isdigit():
                        worksheet.write_number(excelRow, 1, int(row['Room'].replace("Classroom ", "")), roomFormat)
                    else:    
                        worksheet.write(excelRow, 1, row['Room'].replace("Classroom ", "").lstrip('0'), roomFormat)
                worksheet.write(excelRow, 2, row['Section Size'], bodyFormat)
                if "AM" in row['Start Time']:
                    worksheet.write(excelRow, 3, row['Start Time'].lstrip('0'), amFormat)
                else:
                    worksheet.write(excelRow, 3, row['Start Time'].lstrip('0'), bodyFormat)
                if "AM" in row['End Time']:
                    worksheet.write(excelRow, 4, row['End Time'].lstrip('0'), amFormat)
                else:
                    worksheet.write(excelRow, 4, row['End Time'].lstrip('0'), bodyFormat)
                if "Boot Camp" in row['Section Title']:
                    worksheet.write(excelRow, 5, row['Section Title'], laptopReadyFormat)  
                else:
                    worksheet.write(excelRow, 5, row['Section Title'], bodyFormat)
                if row["Instructor"] == "Instructor To Be Announced":
                    worksheet.write(excelRow, 6, 'TBA', instructorFormat)
                elif not pd.isnull(row["Instructor"]):
                    worksheet.write(excelRow, 6, row['Instructor'], instructorFormat)
                worksheet.write(excelRow, 7, row['Section Number'], bodyFormat)
                worksheet.write(excelRow, 8, row['Technology'], bodyFormat)
                worksheet.write(excelRow, 9, row['Notes'], bodyFormat)
                excelRow +=1

        if not afternoonBlock.empty:
            # excelRow += 1
            # worksheet.write(excelRow, 0, 'Afternoon Classes', titleFormat)
            # excelRow += 1
            for idx, row in afternoonBlock.iterrows():
                worksheet.write(excelRow, 0, f"{dateList[i].strftime('%A')}, {dateList[i].strftime('%B %d, %Y').replace(' 0', ' ')}", bodyFormat)
                if location == "GBC" and row['Room'] == "Classroom 201":
                # if row['Room'].replace("Classroom ", "") in ["201", "502", "510", "514", "515"]:
                    worksheet.write_number(excelRow, 1, int(row['Room'].replace("Classroom ", "")), labFormat)
                else:
                    if "Conference Room" in row['Room'] or "Conference Room" in row['Room']:
                        worksheet.write(excelRow, 1, row['Room'].replace("Conference Room ", "CR"), roomFormat)
                    elif row['Room'].replace("Classroom ", "").isdigit():
                        worksheet.write_number(excelRow, 1, int(row['Room'].replace("Classroom ", "")), roomFormat)
                    else:    
                        worksheet.write(excelRow, 1, row['Room'].replace("Classroom ", "").lstrip('0'), roomFormat)
                worksheet.write(excelRow, 2, row['Section Size'], bodyFormat)
                worksheet.write(excelRow, 3, row['Start Time'].lstrip('0'), bodyFormat)
                worksheet.write(excelRow, 4, row['End Time'].lstrip('0'), bodyFormat)
                if "Boot Camp" in row['Section Title']:
                    worksheet.write(excelRow, 5, row['Section Title'], laptopReadyFormat)  
                else:
                    worksheet.write(excelRow, 5, row['Section Title'], bodyFormat)
                if row["Instructor"] == "Instructor To Be Announced":
                    worksheet.write(excelRow, 6, 'TBA', instructorFormat)
                elif not pd.isnull(row["Instructor"]):
                    worksheet.write(excelRow, 6, row['Instructor'], instructorFormat)
                worksheet.write(excelRow, 7, row['Section Number'], bodyFormat)
                worksheet.write(excelRow, 8, row['Technology'], bodyFormat)
                worksheet.write(excelRow, 9, row['Notes'], bodyFormat)
                excelRow +=1

        if not eveningBlock.empty:
            # excelRow += 1
            # worksheet.write(excelRow, 0, 'Evening Classes', titleFormat)
            # excelRow += 1
            for idx, row in eveningBlock.iterrows():
                worksheet.write(excelRow, 0, f"{dateList[i].strftime('%A')}, {dateList[i].strftime('%B %d, %Y').replace(' 0', ' ')}", bodyFormat)
                if location == "GBC" and row['Room'] == "Classroom 201":
                # if row['Room'].replace("Classroom ", "") in ["201", "502", "510", "514", "515"]:
                    worksheet.write_number(excelRow, 1, int(row['Room'].replace("Classroom ", "")), labFormat)
                else:
                    if "Conference Room" in row['Room'] or "Conference Room" in row['Room']:
                        worksheet.write(excelRow, 1, row['Room'].replace("Conference Room ", "CR"), roomFormat)
                    elif row['Room'].replace("Classroom ", "").isdigit():
                        worksheet.write_number(excelRow, 1, int(row['Room'].replace("Classroom ", "")), roomFormat)
                    else:    
                        worksheet.write(excelRow, 1, row['Room'].replace("Classroom ", "").lstrip('0'), roomFormat)
                worksheet.write(excelRow, 2, row['Section Size'], bodyFormat)
                worksheet.write(excelRow, 3, row['Start Time'].lstrip('0'), bodyFormat)
                worksheet.write(excelRow, 4, row['End Time'].lstrip('0'), bodyFormat)
                if "Boot Camp" in row['Section Title']:
                    worksheet.write(excelRow, 5, row['Section Title'], laptopReadyFormat)  
                else:
                    worksheet.write(excelRow, 5, row['Section Title'], bodyFormat)
                if row["Instructor"] == "Instructor To Be Announced":
                    worksheet.write(excelRow, 6, 'TBA', instructorFormat)
                elif not pd.isnull(row["Instructor"]):
                    worksheet.write(excelRow, 6, row['Instructor'], instructorFormat)
                worksheet.write(excelRow, 7, row['Section Number'], bodyFormat)
                worksheet.write(excelRow, 8, row['Technology'], bodyFormat)
                worksheet.write(excelRow, 9, row['Notes'], bodyFormat)
                excelRow +=1

    workbook.close()            
    return 1

def SFCSchedule(schedule, location):
    sortedSchedule = schedule.sort_values(by=['Date', 'Room', 'Start Time'])
    sortedSchedule['Start Time'] = sortedSchedule['Start Time'].dt.strftime('%I:%M %p')
    sortedSchedule['End Time'] = sortedSchedule['End Time'].dt.strftime('%I:%M %p')
    sortedSchedule = sortedSchedule.fillna('')
    dateList = sortedSchedule['Date'].astype(datetime.datetime).unique()

    writer = pd.ExcelWriter(f"{saveReportToPath}/{location} Schedule {dateList[0].strftime('%Y-%m-%d')} {dateList[0].strftime('%A')}.xlsx", engine='xlsxwriter')
    workbook = writer.book

    worksheet = workbook.add_worksheet(location)
    worksheet.set_default_row(hide_unused_rows=True)
    worksheet.set_column('K:XFD', None, None, {'hidden': True})
    worksheet.freeze_panes(2, 0)
    worksheet.set_landscape()       # Page orientation as landscape.
    worksheet.hide_gridlines(0)     # Don’t hide gridlines.
    worksheet.fit_to_pages(1, 0)    # Fit to 1x1 pages.
    worksheet.center_horizontally()
    worksheet.center_vertically()
    worksheet.set_paper(1)          # Set paper size to 8.5" x 11"
    worksheet.set_margins(left=0.25, right=0.25, top=0.25, bottom=0.25)
    worksheet.set_header('', {'margin': 0})
    worksheet.set_footer('', {'margin': 0})

    worksheet.set_column('A:A', 29)   # Column A (Date) width
    worksheet.set_column('B:B', 7.29)   # Column B (Room) width
    worksheet.set_column('C:C', 3.86)   # Column C (Size) width
    worksheet.set_column('D:D', 9.29)   # Column D (Start Date) width
    worksheet.set_column('E:E', 8.43)   # Column E (End Date) width
    worksheet.set_column('F:F', 46.29)   # Column F (Section Title) width
    worksheet.set_column('G:G', 12.57)   # Column G (Instructor) width
    worksheet.set_column('H:H', 16)   # Column H (Section Number) width
    worksheet.set_column('I:I', 18.57)   # Column H Technology) width
    worksheet.set_column('J:J', 64)   # Column H (Notes) width

    locationFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        'font_color': '#000000',
        'bg_color': '#FFC000'
        })

    genFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        'font_color': '#C00000',
        'bg_color': '#FDE9D9'
        })

    titleFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        'font_color': '#000000'
        })

    headerFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        #'valign': 'vcenter',
        'bottom': 2,
        'bottom_color': '#000000'
        })

    bodyFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': False,
        'valign': 'top',
        'text_wrap': False,
        'font_color': '#000000'
        })

    daySeparator = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        'font_color': '#000000',
        'bg_color': '#00B050'
        })

    roomFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        'align': 'right',
        'font_color': '#C00000'
        })

    roomFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        'font_color': '#C00000'
        })

    laptopReadyFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'italic': True,
        'text_wrap': False,
        'font_color': '#0070C0'
        })

    instructorFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': False,
        'text_wrap': False,
        'font_color': '#C00000'
        })

    # AM
    amFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': False,
        'text_wrap': False,
        'font_color': '#FF0000',
        'bg_color': '#DAEEF3'
        })

    # Computer Lab
    labFormat = workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'bold': True,
        'text_wrap': False,
        'font_color': '#FF0000',
        'bg_color': '#FFFF00'
        })

    worksheet.write(0, 0, f"{sortedSchedule.iloc[0][6]}", locationFormat)
    worksheet.write(0, 1, "", locationFormat)
    worksheet.write(0, 2, "", locationFormat)
    worksheet.write(0, 3, "", locationFormat)
    worksheet.write(0, 4, "", locationFormat)
    
    worksheet.write(0, 5, f"Report generated as of {dateList[0].strftime('%A')}, {dateList[0].strftime('%B %d, %Y').replace(' 0', ' ')}", genFormat)
    for col_num, value in enumerate(['Date', 'Room', 'Size', 'Start Time','End Time','Section Title', 'Instructor', 'Section Number','Technology','Notes']):
        worksheet.write(1, col_num, value, headerFormat)
    excelRow = 2
    # loop through each day
    for i in range(0,len(dateList)):
        singleDaySched = sortedSchedule.loc[sortedSchedule['Date'] == dateList[i], : ]
        #morningBlock = singleDaySched.loc[singleDaySched['Start Time'].astype('datetime64') < '12:00:00', : ]
        daytimeBlock = singleDaySched.loc[singleDaySched['Start Time'].astype('datetime64') < '17:00:00', : ]
        #afternoonBlock = singleDaySched.loc[(singleDaySched['Start Time'].astype('datetime64') >= '12:00:00') & (singleDaySched['Start Time'].astype('datetime64') < '17:00:00'), : ]
        eveningBlock = singleDaySched.loc[singleDaySched['Start Time'].astype('datetime64') >= '17:00:00', : ]
        
        #excelRow +=1
        worksheet.write(excelRow, 0, f"{dateList[i].strftime('%A')}, {dateList[i].strftime('%B %d, %Y').replace(' 0', ' ')}", daySeparator)
        worksheet.write_number(excelRow, 3, len(singleDaySched.index), amFormat)
        excelRow +=1
        if not daytimeBlock.empty:
            # worksheet.write(excelRow, 0, 'Morning Classes', titleFormat)
            for idx, row in daytimeBlock.iterrows():
                worksheet.write(excelRow, 0, f"{dateList[i].strftime('%A')}, {dateList[i].strftime('%B %d, %Y').replace(' 0', ' ')}", bodyFormat)
                #if location == "GBC" and row['Room'] == "Classroom 201":
                if row['Room'].replace("Classroom ", "") in ["502", "510", "514", "515"]:
                    worksheet.write_number(excelRow, 1, int(row['Room'].replace("Classroom ", "")), labFormat)
                else:
                    if "Conference Room" in row['Room'] or "Conference Room" in row['Room']:
                        worksheet.write(excelRow, 1, row['Room'].replace("Conference Room ", "CR"), roomFormat)
                    elif row['Room'].replace("Classroom ", "").isdigit():
                        worksheet.write_number(excelRow, 1, int(row['Room'].replace("Classroom ", "")), roomFormat)
                    else:    
                        worksheet.write(excelRow, 1, row['Room'].replace("Classroom ", "").lstrip('0'), roomFormat)
                worksheet.write(excelRow, 2, row['Section Size'], bodyFormat)
                if "AM" in row['Start Time']:
                    worksheet.write(excelRow, 3, row['Start Time'].lstrip('0'), amFormat)
                else:
                    worksheet.write(excelRow, 3, row['Start Time'].lstrip('0'), bodyFormat)
                if "AM" in row['End Time']:
                    worksheet.write(excelRow, 4, row['End Time'].lstrip('0'), amFormat)
                else:
                    worksheet.write(excelRow, 4, row['End Time'].lstrip('0'), bodyFormat)
                if "Boot Camp" in row['Section Title']:
                    worksheet.write(excelRow, 5, row['Section Title'], laptopReadyFormat)  
                else:
                    worksheet.write(excelRow, 5, row['Section Title'], bodyFormat)
                if row["Instructor"] == "Instructor To Be Announced":
                    worksheet.write(excelRow, 6, 'TBA', instructorFormat)
                elif not pd.isnull(row["Instructor"]):
                    worksheet.write(excelRow, 6, row['Instructor'], instructorFormat)
                worksheet.write(excelRow, 7, row['Section Number'], bodyFormat)
                worksheet.write(excelRow, 8, row['Technology'], bodyFormat)
                worksheet.write(excelRow, 9, row['Notes'], bodyFormat)
                excelRow +=1


        if not eveningBlock.empty:
            # excelRow += 1
            # worksheet.write(excelRow, 0, 'Evening Classes', titleFormat)
            # excelRow += 1
            for idx, row in eveningBlock.iterrows():
                worksheet.write(excelRow, 0, f"{dateList[i].strftime('%A')}, {dateList[i].strftime('%B %d, %Y').replace(' 0', ' ')}", bodyFormat)
                #if location == "GBC" and row['Room'] == "Classroom 201":
                if row['Room'].replace("Classroom ", "") in ["502", "510", "514", "515"]:
                    worksheet.write_number(excelRow, 1, int(row['Room'].replace("Classroom ", "")), labFormat)
                else:
                    if "Conference Room" in row['Room'] or "Conference Room" in row['Room']:
                        worksheet.write(excelRow, 1, row['Room'].replace("Conference Room ", "CR"), roomFormat)
                    elif row['Room'].replace("Classroom ", "").isdigit():
                        worksheet.write_number(excelRow, 1, int(row['Room'].replace("Classroom ", "")), roomFormat)
                    else:    
                        worksheet.write(excelRow, 1, row['Room'].replace("Classroom ", "").lstrip('0'), roomFormat)
                worksheet.write(excelRow, 2, row['Section Size'], bodyFormat)
                worksheet.write(excelRow, 3, row['Start Time'].lstrip('0'), bodyFormat)
                worksheet.write(excelRow, 4, row['End Time'].lstrip('0'), bodyFormat)
                if "Boot Camp" in row['Section Title']:
                    worksheet.write(excelRow, 5, row['Section Title'], laptopReadyFormat)  
                else:
                    worksheet.write(excelRow, 5, row['Section Title'], bodyFormat)
                if row["Instructor"] == "Instructor To Be Announced":
                    worksheet.write(excelRow, 6, 'TBA', instructorFormat)
                elif not pd.isnull(row["Instructor"]):
                    worksheet.write(excelRow, 6, row['Instructor'], instructorFormat)
                worksheet.write(excelRow, 7, row['Section Number'], bodyFormat)
                worksheet.write(excelRow, 8, row['Technology'], bodyFormat)
                worksheet.write(excelRow, 9, row['Notes'], bodyFormat)
                excelRow +=1

    workbook.close()            
    return 1
     

if __name__ == "__main__":
    #os.environ["QT_AUTO_SCREEN_FACTOR"] = "1"
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = QtWidgets.QWidget()
    ui = Ui_mainWindow()
    ui.setupUi(mainWindow)
    mainWindow.show()
    sys.exit(app.exec_())
