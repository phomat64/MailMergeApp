import sys
import os
import platform
import subprocess
import pyodbc
import json
import re
from datetime import datetime
from datetime import timedelta
from PyQt5.QtWidgets import (QWidget, 
                            QLabel, 
                            QLineEdit, 
                            QTextEdit, 
                            QVBoxLayout,
                            QHBoxLayout,
                            QGridLayout, 
                            QApplication,
                            QComboBox,
                            QListWidget,
                            QPushButton,
                            QFileDialog,
                            QAction,
                            qApp,
                            QMessageBox,
                            QTabWidget,
                            QTableWidget,
                            QTableWidgetItem,
                            QTableView,
                            QAbstractItemView,
                            QAbstractScrollArea, # remove this
                            QMenu,
                            QListWidgetItem,
                            QStyle,
                            QCheckBox,
                            QFormLayout)
from PyQt5.QtGui import QIcon, QStandardItemModel, QStandardItem
from PyQt5.QtCore import pyqtSlot, pyqtSignal, Qt, QTimer
from mailmerge import MailMerge

# TODO:
# - Add checkbox to show template files in subdirectories in current directory
#   Update: maybe just disable viewing subdirectories completely just in case a user selects the C: drive, the you're app will freeze
# - Remove unneccessary instance variable and use the widgets to hold data instead.
#   Ex:  onTemplateListBoxItemClicked assigns a variable. don't need to. just reference the qlistwidgetitem current item value.
# - Need to reimplement onTmplSearchFilterTyped() since it is horribly inefficient. 
# - Enable saving root path directory to config file
# - Fix layout
# - Make "searchKeyVal" into a instance variable since it is being called twice in different methods.
# - When loading .dot files (which docx-mailmerge cannot read), convert the .dot to .docx and let the user know.
# - Separate save out of save mapping. Just have it save in memory and then have a save button on another tab save configurations.
# - Need to reimplement createFileDictionary() to make it more efficient. Figure out how to use os.walk to also get top-level files and directory list
# - Need to concatenate config file loading errors into one alert box instead of 3 alert boxes popping up separately.
# - Need to research what Qt.UserRole is used for when setting/retrieving data for a QListWidgetItem
# - Add validation to config to make sure the business search keys match the mappings.
# - Add saving of window size and dimensions
# - Need to check if file is already open before generating or close the open file, otherwise we get an error:
#     PermissionError: [Errno 13] Permission denied: 'output\\Catering BW 7-15.docx'
# - Add business headers table fields to config
# - Add label in "Setings" tab letting the user know what configuration they have loaded
class LetterTemplates(QWidget):

    CONFIG_PATH = "./application-default.config.json"
    DB_REGEX = "db{[a-zA-z]+}"
    DATE_REGEX = "date{[\.\(\)_A-Za-z0-9\|\,\/\-\s]+}"
    
    def __init__(self):
        super().__init__()
        self.appConfigPath = None
        self.appConfig = {}
        self.appConfigBackup = {}
        self.dbMappingPattern = re.compile(LetterTemplates.DB_REGEX) # pattern format db{<database-column>}. for ex: db{columnName}
        self.includeSubDir = False
        self.outputDir = "output"
        self.fileNameToFilePathMap = {}
        self.mergeFieldToDataMap = {}
        self.loadConfig(LetterTemplates.CONFIG_PATH)
        self.initUI()

    def initUI(self):
        self.setupUiComponentLayout() # setup layout of components
        self.wireUiComponentsTogether() # setup signals and events for components
        self.initData() # prepopulate the app with necessary data it needs to operate
        self.show()

    def reloadUI(self):
        print("reloading ui...")
        self.initData()
    
    def setupUiComponentLayout(self):
        self.globalLayout = QVBoxLayout()
        # Initialize tab screen
        self.tabs = QTabWidget() # holds all the tabs
        self.mainTab = QWidget()
        self.mappingTab = QWidget()
        self.settingsTab = QWidget()
        
        # Add tabs to parent tab widget
        self.tabs.addTab(self.mainTab, "Main")
        self.tabs.addTab(self.mappingTab, "Mapping")
        self.tabs.addTab(self.settingsTab, "Settings")
        
        # Create Main tab widgets ===========================================
        # business search
        self.businessSearchKeyComboBox = QComboBox(self)

        # business search key label
        # business search key textfield
        self.businessSearchEdit = QLineEdit(self)

        # business search button
        self.openBusinessSearchDialogBtn = QPushButton("...")
        self.openBusinessSearchDialogBtn.setToolTip('Open dialogue to search for business')  

        # root path label
        self.rootDirLabel = QLabel("Root Path:")
        # root path textfield
        self.rootDirEdit = QLineEdit()
        self.rootDirEdit.setReadOnly(True)

        # root path button
        self.chooseRootDirBtn = QPushButton("...")
        self.chooseRootDirBtn.setToolTip('Open dialogue to choose root directory')  

        # template directory dropdown label
        self.templateDirLabel = QLabel('Template Directory:')
        # template directory dropdown
        self.templateDirComboBox = QComboBox(self)

        # checkbox to enable listing templates in subdirectories as well
        self.listSubDirCheckbox = QCheckBox("Include files in subdirectories", self)

        # template search filter
        self.templateSearchFilterEdit = QLineEdit()
        self.templateSearchFilterEdit.setPlaceholderText("Type to search template")

        # template document listbox label
        self.templateListLabel = QLabel('Template List:')
        # template document listbox
        self.templateListBox = QListWidget()

        # Create Letter button
        self.createLetterBtn = QPushButton("Create Letter")

        # create layout and add the main tab widgets
        mainTabLayout = QGridLayout(self)
        mainTabLayout.setSpacing(10)

        mainTabLayout.addWidget(self.rootDirLabel, 0, 0, Qt.AlignRight)
        mainTabLayout.addWidget(self.rootDirEdit, 0, 1)
        mainTabLayout.addWidget(self.chooseRootDirBtn, 0, 2)

        mainTabLayout.addWidget(self.templateDirLabel, 1, 0, Qt.AlignRight)
        mainTabLayout.addWidget(self.templateDirComboBox, 1, 1, 1, 2)

        mainTabLayout.addWidget(self.listSubDirCheckbox, 2, 1)
        mainTabLayout.addWidget(self.templateSearchFilterEdit, 2, 2)

        mainTabLayout.addWidget(self.templateListLabel, 3, 0, Qt.AlignRight | Qt.AlignTop)
        mainTabLayout.addWidget(self.templateListBox, 3, 1, 1, 2)

        mainTabLayout.addWidget(self.businessSearchKeyComboBox, 4, 0, Qt.AlignRight)
        mainTabLayout.addWidget(self.businessSearchEdit, 4, 1)
        mainTabLayout.addWidget(self.openBusinessSearchDialogBtn, 4, 2)

        mainTabLayout.addWidget(self.createLetterBtn, 5, 2)

        # currently we are using gridlayout with 3 columns
        mainTabLayout.setColumnStretch(0, 1)
        mainTabLayout.setColumnStretch(1, 2) # make sure the middle column gets most of the space
        mainTabLayout.setColumnStretch(2, 1)

        # add layout containing the added widgets to the main tab
        self.mainTab.setLayout(mainTabLayout)
        # ===========================================================

        # Create Mapping tab widgets ========================================
        mappingTabLayout = QGridLayout(self) # create layout
        self.mappingTable = QTableWidget()

        self.addNewMappingRowBtn = QPushButton("Add New Row")
        self.revertMappingBtn = QPushButton("Revert")
        self.saveMappingRowBtn = QPushButton("Save")

        # add your newly created widgets to the main layout
        mappingTabLayout.addWidget(self.mappingTable, 0, 0, 4, 4)
        mappingTabLayout.addWidget(self.addNewMappingRowBtn, 0, 4)
        mappingTabLayout.addWidget(self.revertMappingBtn, 1, 4)
        mappingTabLayout.addWidget(self.saveMappingRowBtn, 2, 4)

        # add layout containing the added widgets to the mapping tab
        self.mappingTab.setLayout(mappingTabLayout)
        # ===========================================================

        # Create Settings tab widget
        settingsTabLayout = QFormLayout(self)
        # Config label
        self.configNameLabel = QLabel(self.appConfigPath)
        # Choose configuration file button
        self.chooseConfigFileBtn = QPushButton("Choose Configuration")
        self.chooseConfigFileBtn.setToolTip('Open dialogue to choose a configuration file')  
        # Save configuration button
        self.saveConfigBtn = QPushButton("Save Configuration")
        
        settingsTabLayout.addRow(QLabel("Current Configuration:"), self.configNameLabel)
        settingsTabLayout.addRow(self.chooseConfigFileBtn)
        settingsTabLayout.addRow(self.saveConfigBtn)

        self.settingsTab.setLayout(settingsTabLayout)

        # Add the widgets to the global layout
        # add tab widgets
        self.globalLayout.addWidget(self.tabs)
        
        # add logging widget
        self.logsTextAreaLabel = QLabel("Logs:")
        self.logsTextArea = QTextEdit()
        self.logsTextArea.setReadOnly(True)
        self.logsTextArea.setMaximumHeight(100)
        self.globalLayout.addWidget(self.logsTextAreaLabel)
        self.globalLayout.addWidget(self.logsTextArea)

        # set the global layout, which contains all your widgets, into the parent layout
        self.setLayout(self.globalLayout)

        self.setGeometry(300, 200, 600, 650)
        self.setWindowTitle('Letter Template Generator')

        # setup dialog/windows
        self.businessSearchWindow = BusinessSearchWindow()

    def wireUiComponentsTogether(self):
        # Main tab widgets ===========================================================
        self.chooseRootDirBtn.clicked.connect(self.onChooseRootDirBtnClicked)
        self.openBusinessSearchDialogBtn.clicked.connect(self.onOpenBusinessSearchBtnClicked)

        # template directory dropdown method listener (listens for value changes)
        self.templateDirComboBox.activated[str].connect(self.onTemplateDirComboChanged)
        self.listSubDirCheckbox.stateChanged.connect(self.includeSubDirCheckboxClicked)
        self.templateSearchFilterEdit.textChanged.connect(self.onTmplSearchFilterTyped)

        # template list box item click listener
        self.templateListBox.itemDoubleClicked.connect(self.onTemplateListBoxItemDoubleClicked)

        # allows for right-clicking in the template list box to select menu options
        self.templateListBox.setContextMenuPolicy(Qt.CustomContextMenu)
        self.templateListBox.customContextMenuRequested.connect(self.onTemplateItemContextMenuOpen)

        # business search key type combo box
        # self.businessSearchKeyComboBox.activated[str].connect(self.onTemplateDirComboChanged)
        self.createLetterBtn.clicked.connect(self.onCreateLetterClicked)

        # Mappings tab widgets =======================================================
        self.addNewMappingRowBtn.clicked.connect(self.addNewMappingRowClicked)
        self.revertMappingBtn.clicked.connect(self.revertMappingBtnClicked)
        self.saveMappingRowBtn.clicked.connect(self.saveMappingRowBtnClicked)

        # Settings tab widgets =======================================================
        self.chooseConfigFileBtn.clicked.connect(self.onChooseConfigFileBtnClicked)
        self.saveConfigBtn.clicked.connect(self.onSaveConfigBtnClicked)

        # Dialog/Window widgets ======================================================
        self.businessSearchWindow.businessKeySelectSignal.connect(self.setBusinessKeyText)

    def initData(self):
        initialRootDir = self.appConfig["rootPath"].strip()
        self.initRootDir(initialRootDir)
        self.populateTemplDirComboBox()
        self.populateTemplateListBox()

        self.populateBusinessSearchKeyComboBox()
        self.populateMappingTable()

    def includeSubDirCheckboxClicked(self, state):
        if state == Qt.Checked:
            self.includeSubDir = True
            self.populateTemplateListBox()
        else:
            self.includeSubDir = False
            self.populateTemplateListBox()

    def onSaveConfigBtnClicked(self):
        self.saveConfig(self.appConfig)

    def onChooseConfigFileBtnClicked(self):
        chosenConfigFile = QFileDialog.getOpenFileName(
            self, 
            "Choose a configuration file to load",
            filter=('JSON configuration file (*.json)'))
        if chosenConfigFile:
            # the file object is a tuple, containing two value pairs, so select the first value in the pair which is the string file path to the config
            # ex: ('.../application-default.config.json', 'JSON configuration file (*.json)')
            self.loadConfig(chosenConfigFile[0])
            # update config name label
            self.configNameLabel.setText(chosenConfigFile[0])
            self.reloadUI()

    # TODO: horribly inefficient. this is called on every character that a user types.
    # I just added this to have a quick prototype in the mean time
    def onTmplSearchFilterTyped(self):
        self.populateTemplateListBox()

    def onTemplateItemContextMenuOpen(self, point):
        currentItem = self.templateListBox.currentItem()
        templateFilePath = currentItem.data(Qt.UserRole)
        templateFileDir = os.path.dirname(templateFilePath)

        contextMenu = QMenu(self)
        open_file_action = contextMenu.addAction("Open File")
        nav_to_file_dir_action = contextMenu.addAction("Go to Directory")
        # use self.sender().mapToGlobal() and not self.mapToGlobal() or it the context menu will pop up in the wrong position
        action = contextMenu.exec_(self.sender().mapToGlobal(point)) 
        if action == open_file_action:
            print("opening file: " + templateFilePath)
            self.launchTargetPath(templateFilePath)
        elif action == nav_to_file_dir_action:
            print("opening directory: " + templateFileDir)
            self.launchTargetPath(templateFileDir)

    # TODO: method is a little bloated. need to streamline
    def initRootDir(self, newRootDir):
        # check if a root path was configured in the config file
        try:
            if newRootDir:
                self.rootDirEdit.setText(newRootDir)
            else:
                self.rootDirEdit.setText(os.path.abspath(".")) # default directory
        except KeyError:
            print("Optional root directory path was not found in the configuration. Setting to current directory.")
            self.rootDirEdit.setText(os.path.abspath(".")) # default directory

    def populateMappingTable(self):
        # clear table to make a blank slate
        self.mappingTable.setRowCount(0)

        # get mappings from config file
        fieldMappings = self.appConfig["fieldMappings"]

        # setup table column headers
        self.mappingColumnHeaders = ["Merge Field", "Value", ""]
        self.mappingTable.setColumnCount(len(self.mappingColumnHeaders))
        # set a wider column width for the "Merge Field" column
        self.mappingTable.setColumnWidth(0, 175)
        # set a wider column width for the "Value" column
        self.mappingTable.setColumnWidth(1, 180)
        # set a wider column width for the Delete buttons column
        self.mappingTable.setColumnWidth(2, 40)

        self.mappingTable.setHorizontalHeaderLabels(self.mappingColumnHeaders)

        # add mapping data to table widget
        for i, self.item in enumerate(fieldMappings):
            field = str(self.item["field"])
            value = str(self.item["value"])
            self.addNewRowToMappingTable(field, value)

    def addNewRowToMappingTable(self, newField, newValue):
        # insert a new blank row before you add the data
        newRowIndex = self.mappingTable.rowCount()
        self.mappingTable.insertRow(newRowIndex) 

        # create delete button for mapping
        deleteButton = QPushButton()
        deleteButton.setIcon(self.style().standardIcon(QStyle.SP_TitleBarCloseButton))
        deleteButton.setToolTip("Delete this field mapping row.")
        # deleteButton.setStyleSheet("background-color: transparent")
        deleteButton.clicked.connect(self.deleteMappingRowClicked)

        self.mappingTable.setItem(newRowIndex, 0, QTableWidgetItem(newField))
        self.mappingTable.setItem(newRowIndex, 1, QTableWidgetItem(newValue))
        self.mappingTable.setCellWidget(newRowIndex, 2, deleteButton)

    def setBusinessKeyText(self, businessKey):
        self.businessSearchEdit.setText(businessKey)

    @pyqtSlot() # not sure if i need that annotation there, since it works without it
    def deleteMappingRowClicked(self):
        # get a reference to the "Delete" button that called this method
        clickedButton = self.sender()
        if clickedButton:
            clickedButtonRowIndex = clickedButton.pos()
            rowToRemove = self.mappingTable.indexAt(clickedButtonRowIndex).row()
            self.mappingTable.removeRow(rowToRemove)

    def revertMappingBtnClicked(self):
        revertMappingBtnReply = QMessageBox.question(self, 'Revert Mappings', "Are you sure you want to overwrite the current mappings?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if revertMappingBtnReply == QMessageBox.Yes:
            # restore mappings from backup
            self.appConfig["fieldMappings"] = self.appConfigBackup["fieldMappings"].copy()
            self.populateMappingTable() # repopulate table with reloaded config
        else:
            print('Save Mapping process canceled.')

    def saveMappingRowBtnClicked(self):
        saveMappingBtnReply = QMessageBox.question(self, 'Save Mappings', "Are you sure you want to overwrite the current mappings?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if saveMappingBtnReply == QMessageBox.Yes:
            self.saveMappingTableDataToConfig()
        else:
            print('Save Mapping process canceled.')

    # saves mapping to config object in memory, not the external file
    def saveMappingTableDataToConfig(self):
        # build new field mappings array from mapping table
        newFieldMappings = []
        rowCount = self.mappingTable.rowCount()
        for row in range(rowCount):
            fieldItem = self.mappingTable.item(row, 0)
            valueItem = self.mappingTable.item(row, 1)
            newMappingPair = {
                "field" : fieldItem.text(),
                "value" : valueItem.text()
            }
            newFieldMappings.append(newMappingPair)
        # replace the old field mappings with the new mappings in the app config variable
        self.appConfig["fieldMappings"] = newFieldMappings

    def loadConfig(self, configFilePath):
        try:
            with open(configFilePath) as json_config_file:  
                self.appConfig = json.load(json_config_file)
                self.validateConfig(self.appConfig)
                # if config file is valid, save the file path to the config
                self.appConfigPath = configFilePath
                # make a backup copy just in case the user wants to revert some changes
                self.appConfigBackup = self.appConfig.copy()
                
        except FileNotFoundError:
            print("User canceled load config process.")
    
    # check if certain properties are present in the config file, otherwise throw error message
    def validateConfig(self, config):
        try:
            config["database"]["connectionString"]
            config["database"]["sqlQuery"]   
        except KeyError:
            print("Database configurations are missing.")
            QMessageBox.about(self, "Configuration Error", "Database configurations are missing from the config file. Please ensure the 'connectionString' and 'sqlQuery' properties are in the config file.")
        
        try:
            config["business"]["search_key_map"]
        except KeyError:
            print("Business key configurations are missing.")
            QMessageBox.about(self, "Configuration Error", "Business keys are missing from the configuration file.")
        
        try:
            config["fieldMappings"]
        except KeyError:
            print("Field mappings are missing.")
            QMessageBox.about(self, "Configuration Error", "Field mappings are missing from the configuration file.")

    def saveConfig(self, configObj):
        # Save root path
        self.appConfig["rootPath"] = self.rootDirEdit.text()

        # Before saving the mappings to the config file, lets clean it up by removing any blank values
        validFieldMappings = []
        for fieldMap in self.appConfig["fieldMappings"]:
            if fieldMap["field"] and fieldMap["value"]:
                validFieldMappings.append(fieldMap)
            else:
                print("Blank field map was found: " + str(fieldMap))
        self.appConfig["fieldMappings"] = validFieldMappings

        try:
            with open(self.appConfigPath, 'w') as updated_config_file:  
                json.dump(configObj, updated_config_file, indent = 4, sort_keys = False)
            QMessageBox.about(self, "Configuration Saved", "Saved to configuration file: \"" + self.appConfigPath + "\"")
        except FileNotFoundError:
            QMessageBox.about(self, "Unable to save configurations", "Target cConfiguration file was not found: \"" + self.appConfigPath + "\"")
    
    def populateTemplDirComboBox(self):
        self.templateDirComboBox.clear()
        subDirectoryList = filter(self.isdir, os.listdir(self.rootDirEdit.text()))
        self.templateDirComboBox.addItem("") # default value
        for subDir in subDirectoryList:
            self.templateDirComboBox.addItem(subDir)

    def isdir(self, path):
        fullPath = os.path.join(self.rootDirEdit.text(), path)
        return os.path.isdir(fullPath)

    def populateBusinessSearchKeyComboBox(self):
        self.businessSearchKeyComboBox.clear()
        for searchKey in self.appConfig["business"]["search_key_map"]:
            self.businessSearchKeyComboBox.addItem(searchKey)

    def populateTemplateListBox(self):
        # clear the list every time to make room for the new list of file templates
        self.templateListBox.clear()

        # combine root directory with template directory name
        fullPathToTemplates = os.path.join(self.rootDirEdit.text(), str(self.templateDirComboBox.currentText()))

        # before you add the template items, add an item that allows the user to go up one level when double clicking
        goUpOneLevelItem = QListWidgetItem()
        templateDirParentPath = os.path.abspath(os.path.join(fullPathToTemplates, os.pardir))
        goUpOneLevelItem.setText("Go to parent directory")
        goUpOneLevelItem.setData(Qt.UserRole, templateDirParentPath)
        goUpOneLevelItem.setToolTip(templateDirParentPath)
        goUpOneLevelItem.setIcon(self.style().standardIcon(QStyle.SP_FileDialogToParent))
        self.templateListBox.addItem(goUpOneLevelItem)
         
        # get list of file templates from the selected directory
        self.fileNameToFilePathMap = self.createFileDictionary(fullPathToTemplates, self.includeSubDir)
        
        # finally, display any file templates that were found
        mainItemList = self.convertFileMapToWidgetItemList(self.fileNameToFilePathMap)  
        for item in mainItemList:
            self.templateListBox.addItem(item)

    def convertFileMapToWidgetItemList(self, fileNamePathMap):
        # we are going to organize our items into two lists; one for directories and one for files.
        # this is purely for presentation; separate items into two lists, sort them by name, then
        # add both lists to the main list. 
        directoryItems = []
        fileItems = []
        mainItemList = []

        for fileName in fileNamePathMap:
            fileFullPathName = fileNamePathMap[fileName]
            searchFilterText = self.templateSearchFilterEdit.text().strip()
            if re.search(searchFilterText, fileName, re.IGNORECASE):
                newItem = QListWidgetItem()
                
                # assign core info for item
                newItem.setText(fileName)
                newItem.setData(Qt.UserRole, fileFullPathName)
                newItem.setToolTip(fileFullPathName)

                # assign appropriate icon and add to its appropriate list
                if os.path.isdir(fileFullPathName):
                    newItem.setIcon(self.style().standardIcon(QStyle.SP_DirHomeIcon))
                    directoryItems.append(newItem)
                else:
                    newItem.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))   
                    fileItems.append(newItem)
        
        # sort each item list by their display name
        directoryItems.sort(key=lambda x: x.text().lower())
        fileItems.sort(key=lambda x: x.text().lower())

        # add items to main list
        mainItemList.extend(directoryItems)
        mainItemList.extend(fileItems)
        return mainItemList

    # scans the specified directories/subdirectories and creates a map of filenames and their filepaths
    # TODO: make this more efficient. 
    def createFileDictionary(self, dirPath, includeSubDirFiles=False):
        fileToPathMap = {}
    
        if includeSubDirFiles:
            for path, subdirs, files in os.walk(dirPath):
                for fileName in files:
                    if (self.validFile(fileName)):
                        filePath = os.path.join(path, fileName)
                        fileToPathMap[fileName] = filePath
        else:
            # just list files and directories in the top-level of the current path
            for fileName in os.listdir(dirPath):
                fullFilePath = os.path.join(dirPath, fileName)
                if (self.validFile(fullFilePath)):   
                    fileToPathMap[fileName] = fullFilePath
        return fileToPathMap

    # we only want regular file names. 
    # For example, we don't want "hidden" files like ".DS_Store"
    # and other files we may add in the future
    def validFile(self, fileName):
        return (os.path.isdir(fileName) or 
                fileName.endswith(".docx") or 
                fileName.endswith(".doc") or 
                fileName.endswith(".dot"))

    def onTemplateDirComboChanged(self, newTemplateDirVal):
        print("Template directory changed to: " + newTemplateDirVal)
        self.populateTemplateListBox()

    def onTemplateListBoxItemDoubleClicked(self, newTemplateListBoxItem):
        templateFilePath = newTemplateListBoxItem.data(Qt.UserRole)
        if os.path.isdir(templateFilePath):
            self.initRootDir(templateFilePath)
            self.populateTemplDirComboBox()
            self.populateTemplateListBox()
        else:
            self.launchTargetPath(templateFilePath)

    def onChooseRootDirBtnClicked(self):
        chosenRootDir = QFileDialog.getExistingDirectory(
            self, 
            "Choose the root directory of the letter templates")
        if chosenRootDir:
            self.rootDirEdit.setText(chosenRootDir) # update root textfield of new root path
            self.populateTemplDirComboBox()
            # populate template list with templates from the chosen root directory
            self.populateTemplateListBox() 
        else:
            print("Chosen directory was empty. User must've canceled.")

    def addNewMappingRowClicked(self):
        self.addNewRowToMappingTable("", "")

    def onOpenBusinessSearchBtnClicked(self):
        businessSearchWindowParams = {
            "database": self.appConfig["database"],
            "search_table_column_mapping": self.appConfig["business"]["search_table_column_mapping"]
        }
        self.businessSearchWindow.setBusinessData(businessSearchWindowParams)
        self.businessSearchWindow.show()

    def getBaseSqlQuery(self):
        return self.appConfig["database"]["sqlQuery"]

    def getCurrentTemplateListItem(self):
        return self.templateListBox.currentItem()

    # TODO: instead of getting file path from map, lets get it from the
    # current item in the QListWidget template list. currentItem.data(Qt.UserRole)
    def onCreateLetterClicked(self):
        print("Creating Letter...")
        if self.getCurrentTemplateListItem():
            currentTemplateItemName = self.getCurrentTemplateListItem().text()
            businessSearchKeyVal = self.businessSearchEdit.text()
            if businessSearchKeyVal:
                templateFilePath = self.fileNameToFilePathMap[currentTemplateItemName]
                self.mailMergeDocument(templateFilePath)
            else:
                QMessageBox.about(self, "Invalid business search key", "Please enter a business search key value for \"" + str(self.businessSearchKeyComboBox.currentText()) + "\"")
        else:
            QMessageBox.about(self, "No template selected", "Please select a template from the list.")
    
    def mailMergeDocument(self, templateFilePath):
        with MailMerge(templateFilePath) as template:
            # print("Current merge fields in document: \n" + str(template.get_merge_fields()))
            if len(template.get_merge_fields()) > 0:
                fieldMappings = self.createMergeFieldMap()
                
                if len(fieldMappings) > 0:
                    # merge data into template
                    template.merge_pages([fieldMappings])

                    # if output directory does not exist, then create it
                    if not os.path.exists(os.path.join(os.path.abspath("."), self.outputDir)):
                        os.mkdir(self.outputDir)
        
                    filledTemplateFilePath = os.path.join(self.outputDir, self.getCurrentTemplateListItem().text())
                    try:
                        template.write(filledTemplateFilePath)
                        self.launchTargetPath(filledTemplateFilePath)
                    except PermissionError:
                        QMessageBox.about(self, "Unable to generate template", "Cannot write to output file \"" + filledTemplateFilePath + "\". \nPlease any previously generated templates that are still open.")
                else:
                    print("Merge field map was empty. No field mappings were evaluated from the configuration.")
            else:
                QMessageBox.about(self, "Unable to Create Letter", "The template you selected does not contain any merge fields. \n\n" + templateFilePath)

    def createMergeFieldMap(self):
        # get a map of all the named queries and their resultsets
        namedQueryDatasetMap = createNamedQueryDatasetMap()
        fieldMappings = {}
        if namedQueryDatasetMap:
            for mapping in self.appConfig["fieldMappings"]:
                newValue = ""
                field = mapping["field"]
                value = mapping["value"]
                namedQuery = mapping["namedQuery"]

                # TODO: get namedQuery result map
                
                # compile any expressions into values
                # Hello db{columnName}, how are you, db{columnName}
                dataset = namedQueryDatasetMap[namedQuery]
                if dataset:
                    newValue = self.evaluateValueExpression(dataset, value)
                    fieldMappings[field] = newValue
                else:
                    QMessageBox.about(self, "Named Query Not Found", "Could not find named query, \"" + namedQuery + "\". Please make sure your SQL configuration is setup correctly.")
        else:
            QMessageBox.about(self, "Empty Result Set", "No data was returned from the query. Please make sure your SQL configuration is setup correctly.")
        return fieldMappings
    
    def createNamedQueryDatasetMap(self):
        namedQueryDatasetMap = {}

        searchKeyType = str(self.businessSearchKeyComboBox.currentText())
        businessSearchObj = self.appConfig["business"]["search_key_map"][searchKeyType]
        params = [
            { 
                "column" : businessSearchObj["column"], 
                "compareOperator" : businessSearchObj["compareOperator"], 
                "value" : self.businessSearchEdit.text() 
            }
        ]

        for namedQueryKey in self.appConfig["database"]["namedQueries"]:
            dataQuery = {
                "baseSql" : self.appConfig["database"]["namedQueries"][namedQueryKey]["query"],
                "params" : params
            }
            namedQueryDatasetMap[namedQueryKey] = self.retrieveDataSet(dataQuery)

        return namedQueryDatasetMap

    def createUnnamed1Query(self):
        searchKeyType = str(self.businessSearchKeyComboBox.currentText())
        businessSearchObj = self.appConfig["business"]["search_key_map"][searchKeyType]
        dataQuery = {
            "baseSql" : self.getBaseSqlQuery(),
            "params" : [
                { 
                  "column" : businessSearchObj["column"], 
                  "compareOperator" : businessSearchObj["compareOperator"], 
                  "value" : self.businessSearchEdit.text() 
                }
            ]
        }
        return dataQuery


    def evaluateValueExpression(self, dataSet, value,):
        evaluatedExpressionsDict = {}
        
        # check if there is a filter expression, "<main-expression> || <filter-expression>"
        # ex: db{column_name} || date : mmmm dd, yyyy
        # split string only on the first occurrence, other wise it will interfere with the filter params
        expressionParts = value.split("||")

        mainExpression = expressionParts[0].strip()

        # evaluate db expressions
        for dbExpr in re.findall(LetterTemplates.DB_REGEX, mainExpression):
            evaluatedDbExpr = self.evaluateDatabaseExpr(dataSet, dbExpr)
            # add evaluated db expression to dictionary
            evaluatedExpressionsDict.update({dbExpr : evaluatedDbExpr})
        
        # evaluate date expressions
        for dateExpr in re.findall(LetterTemplates.DATE_REGEX, mainExpression):
            evaluatedDateExpr = self.evaluateDateExpr(dateExpr)
            # add evaluated date expression to dictionary
            evaluatedExpressionsDict.update({dateExpr : evaluatedDateExpr})

        # replace expression strings with the newly evaluated ones
        for exprKey in evaluatedExpressionsDict.keys():
            mainExpression = mainExpression.replace(exprKey, evaluatedExpressionsDict[exprKey])
        
        # apply any filters to the evaluated expression
        mainExpression = self.applyFilters(expressionParts, mainExpression)

        return mainExpression.strip()

    def applyFilters(self, expressionParts, mainExpression):
        if (len(expressionParts) > 1):
            filterExpr = expressionParts[1]
            filterExprParts = filterExpr.split(":")
            if (len(filterExprParts) > 1):
                filterCommand = filterExprParts[0].strip()
                filterCommandParams = filterExprParts[1].strip() 
                if (filterCommand == "date"):
                    # "yyyymmdd" "mmmm dd, yyyy"
                    mainExpression = self.applyDateFilter(filterCommandParams, mainExpression)
            else:
                print("Filter expression is not valid: " + filterExpr)
        return mainExpression

    def applyDateFilter(self, filterCommandParams, mainExpression):
        dataFilterParamParts = re.findall("(?:\".*?\"|\S)+", filterCommandParams)
        fromDateFormat = dataFilterParamParts[0].replace("\"", "") # remove quotes
        fromDateFormat = self.convertUserDateFormatStrToNativeStr(fromDateFormat) 
        toDateFormat = dataFilterParamParts[1].replace("\"", "") # remove quotes
        toDateFormat = self.convertUserDateFormatStrToNativeStr(toDateFormat)

        # convert evaluated main expression date to the other format
        datetimeObjForOriginal = datetime.strptime(mainExpression,fromDateFormat)
        mainExpression = datetimeObjForOriginal.strftime(toDateFormat)
        return mainExpression

    def evaluateDatabaseExpr(self, dataSet, dbExpr):
        newValue = ""
        # extract column name from "db{<column-name>}", which means remove the "db{" and the "}" leaving you the <column-name>
        db_column = re.search('db{(.+?)}', dbExpr).group(1)
        # get the value from the database result row using the column name as key
        try:
            # if multiple data rows returned, then concatenate each value into a single string separated by newline characters "\n".
            if len(dataSet) > 1:
                for row in dataSet:
                    newValue += "\n" + getattr(row, db_column)
            else:
                newValue = getattr(dataSet[0], db_column)
        except AttributeError:
            print("Database column, \""+ db_column +"\" not found in result set. Cannot evaluate db expression: " + dbExpr)
            newValue = "<UNKNOWN VALUE>"
        return newValue
    
    def evaluateDateExpr(self, dateExpr):
        # date{current.add(90) | mmmm/dd/yyyy}

        evaluatedDateExpr = None
        dateFormatStr = "mmmm dd, yyyy" # default date format "July 13, 2019"

        # extract inner date expression inside date{ <inner-expression-here> }
        # be aware that an optional date format string could be provided as the second parameter
        # ex: "current.add(90)" or "current.add(90) | mmmm/dd/yyyy"
        innerDateExpr = re.search('date{(.+?)}', dateExpr).group(1)

        # split the inner date expression using a comma to see if the user supplied a second parameter (date format string)
        # after splitting ex: parts[0] = "current.add(90)" parts[1] = "mmmm/dd/yyyy"
        innerDateExprParts = innerDateExpr.split("|")

        # check if current.add command
        if re.match("current.add\([0-9]+\)", innerDateExprParts[0].strip()):
            # extract day parameter
            numOfDaysToAdd = re.search('current.add\((.+?)\)', innerDateExprParts[0]).group(1)
            # create current date with specified days added
            evaluatedDateExpr = datetime.now() + timedelta(days=int(numOfDaysToAdd))

        # if date format string argument was found, then use that format
        if len(innerDateExprParts) > 1:
            # convert user-specified date format string to python's date format string then format the date. Ex: "yyyy" => "%Y"
            dateFormatStr = innerDateExprParts[1].strip()

        dateFormatStr = self.convertUserDateFormatStrToNativeStr(dateFormatStr)
        evaluatedDateExpr = evaluatedDateExpr.strftime(dateFormatStr)

        return evaluatedDateExpr

    def convertUserDateFormatStrToNativeStr(self, originalDateFormatStr):
        dateFormatStr = originalDateFormatStr
        if "mmmm" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("mmmm", "%B")
        if "mmm" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("mmm", "%b")
        if "mm" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("mm", "%m")
        if "dddd" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("dddd", "%A")
        if "ddd" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("ddd", "%a")
        if "dd" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("dd", "%d")
        if "yyyy" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("yyyy", "%Y")
        if "yy" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("yy", "%y")
        return dateFormatStr

    def convertNativeStrToUserDateFormatStr(self, originalDateFormatStr):
        dateFormatStr = originalDateFormatStr
        if "%B" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("%B", "mmmm")
        if "%b" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("%b", "mmm")
        if "%m" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("%m", "mm")
        if "%A" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("%A", "dddd")
        if "%a" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("%a", "ddd")
        if "%d" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("%d", "dd")
        if "%Y" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("%Y", "yyyy")
        if "%y" in dateFormatStr:
            dateFormatStr = dateFormatStr.replace("%y", "yy")
        return dateFormatStr

    def retrieveDataSet(self, dataQuery):
        sqlQuery = dataQuery["baseSql"]
        paramTupleArr = ()
        connStr = self.appConfig["database"]["connectionString"]
        conn = pyodbc.connect(connStr)
        cursor = conn.cursor()
        if dataQuery["params"]: 
            sqlQuery = sqlQuery + " where "
            firstConditionalAdded = False
            for queryCondition in dataQuery["params"]:
                if firstConditionalAdded:
                    sqlQuery = sqlQuery + " and "
                sqlQuery = sqlQuery + queryCondition["column"] + " " + queryCondition["compareOperator"] + " ?"
                firstConditionalAdded = True
                paramTupleArr = paramTupleArr + (queryCondition["value"],)
        
        cursor.execute(sqlQuery, paramTupleArr)
        return cursor.fetchall()

    # opens a file if path points to a file.
    # opens a directory if path points to a directory.
    def launchTargetPath(self, docFilePath):
        # start file. assuming we are on windows. if on linux system, it will throw error.
        # but, we will catch error and launch it using the linux way instead
        try:
            if platform.system() == 'Windows':
                os.startfile(docFilePath)
            else:
                subprocess.call(['open', docFilePath])
        except AttributeError:
            print("An error occurred while attempting to launch the file: " + docFilePath)




class BusinessSearchWindow(QWidget):
    businessKeySelectSignal = pyqtSignal(str)

    def __init__(self):
        super(BusinessSearchWindow, self).__init__()

        self.dbConnStr = None
        self.sqlQuery = None
        self.pageNum = 0
        self.pageSize = 5
        self.totalPages = 0
        self.dataCount = 0
        self.oldSearchText = None

        # add components to layout
        businessSearchLayout = QVBoxLayout()

        # business search filter
        self.businessSearchFilterEdit = QLineEdit()
        self.businessSearchFilterEdit.setPlaceholderText("Type to search table...")
        # business search filter combobox
        self.businessSearchComboBox = QComboBox(self)
        # create horizontal layout for search filter components
        searchFilterLayout = QHBoxLayout()
        searchFilterLayout.addWidget(self.businessSearchFilterEdit)
        searchFilterLayout.addWidget(self.businessSearchComboBox)

        self.pageButtonsLayout = QHBoxLayout()
        self.pageButtonsLayout.addStretch(1)

        self.pageLeftBtn = QPushButton()
        self.pageLeftBtn.setIcon(self.style().standardIcon(QStyle.SP_ArrowLeft))
        self.pageButtonsLayout.addWidget(self.pageLeftBtn)

        self.pageInfoLabel = QLabel("0 of 0")
        self.pageButtonsLayout.addWidget(self.pageInfoLabel)

        self.pageRightBtn = QPushButton()
        self.pageRightBtn.setIcon(self.style().standardIcon(QStyle.SP_ArrowRight))
        self.pageButtonsLayout.addWidget(self.pageRightBtn)

        self.pageSizeLabel = QLabel("Page Size:")
        self.pageButtonsLayout.addWidget(self.pageSizeLabel)

        self.pageSizeComboBox = QComboBox(self)
        self.pageSizeComboBox.addItem("5")
        self.pageSizeComboBox.addItem("10")
        self.pageSizeComboBox.addItem("15")
        self.pageSizeComboBox.addItem("20")

        pageComponentsLayout = QHBoxLayout()
        pageComponentsLayout.addLayout(self.pageButtonsLayout)
        pageComponentsLayout.addWidget(self.pageSizeComboBox)

        self.businessTable = QTableView()
        self.model = QStandardItemModel(self)
        self.businessTable.setModel(self.model) 
        # disable cell editing
        self.businessTable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        # set selection behavior
        # self.businessTable.setSelectionBehavior(QAbstractItemView.SingleSelection)

        self.select_button = QPushButton("Select Value")
        self.close_button = QPushButton("Close")
        businessSearchLayout.addLayout(searchFilterLayout)
        businessSearchLayout.addWidget(self.businessTable)
        businessSearchLayout.addLayout(pageComponentsLayout)
        businessSearchLayout.addWidget(self.select_button)
        businessSearchLayout.addWidget(self.close_button)
        self.setLayout(businessSearchLayout)
        self.setWindowTitle("Business Search")
        self.setMinimumWidth(650)
        self.setMinimumHeight(550)

        # wire up components
        self.businessSearchFilterEdit.textChanged.connect(self.onBusSearchFilterTyped)
        self.pageSizeComboBox.activated[str].connect(self.onNumRowsComboChanged)
        self.pageLeftBtn.clicked.connect(self.showPrevPage)
        self.pageRightBtn.clicked.connect(self.showNextPage)
        self.select_button.clicked.connect(self.selectBusinessKey)
        self.close_button.clicked.connect(self.closeWindow)

        # setup search filter timer
        # this is to prevent a database call after every character typed. wait until the user finishes typing.
        self.searchTimer = QTimer()
        self.searchTimer.setInterval(1000)
        self.searchTimer.timeout.connect(self.executeSearch)

    def showPrevPage(self):
        newPageNum = self.pageNum - 1
        if newPageNum < 0:
            newPageNum = 0
        self.showPage(newPageNum)

    def showNextPage(self):
        newPageNum = self.pageNum + 1
        if newPageNum < self.totalPages:
            self.showPage(newPageNum)

    def showPage(self, pageNum):
        self.pageNum = pageNum
        self.populateTable()

    def onNumRowsComboChanged(self):
        self.pageSize = int(self.pageSizeComboBox.currentText())
        self.populateTable()

    def closeWindow(self):
        self.close()

    def setBusinessData(self, searchParams):
        self.dbConnStr = searchParams["database"]["connectionString"]
        self.sqlQuery = searchParams["database"]["sqlQuery"]

        self.searchTableColMap = searchParams["search_table_column_mapping"]
        # set search filter combobox
        self.businessSearchComboBox.addItems(self.searchTableColMap.keys())

        self.tableFields = self.searchTableColMap.keys()
        self.model.setHorizontalHeaderLabels(self.tableFields)
        # adjust columns
        self.businessTable.setColumnWidth(0, 120)
        self.businessTable.setColumnWidth(2, 120)
        self.businessTable.setColumnWidth(4, 170)

        self.populateTable()
        
    def populateTable(self):
        # clear old data from table before adding data to it
        self.model.setRowCount(0)

        # Get the dataset
        conn = pyodbc.connect(self.dbConnStr)
        cursor = conn.cursor()
        # Add pagination query
        sqlQueryWithParams = self.sqlQuery + ' order by (select null) offset ? rows fetch next ? rows only'
        offset = self.pageNum * self.pageSize
        cursor.execute(sqlQueryWithParams, offset, self.pageSize)

        # add business data to business table
        for row in cursor:
            # insert a new blank row before you add the data
            newRowIndex = self.model.rowCount()
            self.model.insertRow(newRowIndex)
            # add data for each column for the current row
            for index, db_column_obj in enumerate(self.searchTableColMap.values()):   
                self.model.setItem(newRowIndex, index, QStandardItem(getattr(row, db_column_obj["column"])))
        
        # update data count info
        # need to replace "select *" with "select count(*)"
        sqlCountQuery = self.sqlQuery.replace("*", "count(*)", 1)
        cursor.execute(sqlCountQuery)
        self.dataCount = cursor.fetchone()[0]
        self.totalPages = 1 if (self.dataCount / self.pageSize) == 0 else (self.dataCount / self.pageSize)
        self.pageInfoLabel.setText(str(self.pageNum+1) + " of " + str(int(self.totalPages)))

    def onBusSearchFilterTyped(self):
        print("business search filter: " + self.businessSearchFilterEdit.text())
        searchText = self.businessSearchFilterEdit.text().strip()
        if (self.oldSearchText != searchText):
            self.oldSearchText = searchText
            self.searchTimer.start()
        else:
            print("stopping timer")
            self.searchTimer.stop()
    
    def executeSearch(self):
        print("executing search...")
        self.searchTimer.stop()
    
    def selectBusinessKey(self, signal):
        if self.businessTable.selectionModel().hasSelection():
            currentSelectedData = self.businessTable.selectionModel().currentIndex().data()
            self.businessKeySelectSignal.emit(currentSelectedData)
        else:
            QMessageBox.about(self, "No data selected", "No data was selected. Please select a value in the table.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = LetterTemplates()
    sys.exit(app.exec_())