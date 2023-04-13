import sys
import time
import lookups
import functions
import pandas as pd
import dataframe_image
from os.path import exists, join
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QFileDialog, QVBoxLayout, QProgressBar, QMessageBox, QComboBox, QStyledItemDelegate, QProgressDialog
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QEvent
from PyQt5 import QtGui
from PyQt5.uic import loadUi


class CheckableComboBox(QComboBox):

    # Subclass Delegate to increase item height
    class Delegate(QStyledItemDelegate):
        def sizeHint(self, option, index):
            size = super().sizeHint(option, index)
            size.setHeight(25)
            return size


    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Make the combo editable to set a custom text, but readonly
        self.setEditable(True)
        self.lineEdit().setReadOnly(True)
        # Make the lineedit the same color as QPushButton
        palette = QtWidgets.QApplication.palette()
        palette.setBrush(QtGui.QPalette.Base, palette.button())
        self.lineEdit().setPalette(palette)

        # Use custom delegate
        self.setItemDelegate(CheckableComboBox.Delegate())

        # Update the text when an item is toggled
        self.model().dataChanged.connect(self.updateText)

        # Hide and show popup when clicking the line edit
        self.lineEdit().installEventFilter(self)
        self.closeOnLineEditClick = False

        # Prevent popup from closing when clicking on an item
        self.view().viewport().installEventFilter(self)


    def resizeEvent(self, event):
        # Recompute text to elide as needed
        self.updateText()
        super().resizeEvent(event)


    def eventFilter(self, object, event):

        if object == self.lineEdit():
            if event.type() == QEvent.MouseButtonRelease:
                if self.closeOnLineEditClick:
                    self.hidePopup()
                else:
                    self.showPopup()
                return True
            return False

        if object == self.view().viewport():
            if event.type() == QEvent.MouseButtonRelease:
                index = self.view().indexAt(event.pos())
                item = self.model().item(index.row())

                if item.checkState() == Qt.Checked:
                    item.setCheckState(Qt.Unchecked)
                else:
                    item.setCheckState(Qt.Checked)
                return True
        return False


    def showPopup(self):
        super().showPopup()
        # When the popup is displayed, a click on the lineedit should close it
        self.closeOnLineEditClick = True


    def hidePopup(self):
        super().hidePopup()
        # Used to prevent immediate reopening when clicking on the lineEdit
        self.startTimer(100)
        # Refresh the display text when closing
        self.updateText()


    def timerEvent(self, event):
        # After timeout, kill timer, and reenable click on line edit
        self.killTimer(event.timerId())
        self.closeOnLineEditClick = False


    def updateText(self):
        texts = []
        for i in range(self.model().rowCount()):
            if self.model().item(i).checkState() == Qt.Checked:
                texts.append(self.model().item(i).text())
        text = ", ".join(texts)

        # Compute elided text (with "...")
        metrics = QtGui.QFontMetrics(self.lineEdit().font())
        elidedText = metrics.elidedText(text, Qt.ElideRight, self.lineEdit().width())
        if elidedText:
            self.lineEdit().setText(elidedText)
        else:
            self.lineEdit().setText('Select...')


    def addItem(self, text, data=None):
        item = QtGui.QStandardItem()
        item.setText(text)
        if data is None:
            item.setData(text)
        else:
            item.setData(data)
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
        item.setData(Qt.Unchecked, Qt.CheckStateRole)
        self.model().appendRow(item)


    def addItems(self, texts, datalist=None):
        for i, text in enumerate(texts):
            try:
                data = datalist[i]
            except (TypeError, IndexError):
                data = None
            self.addItem(text, data)


    def currentData(self):
        # Return the list of selected items data
        res = []
        for i in range(self.model().rowCount()):
            if self.model().item(i).checkState() == Qt.Checked:
                res.append(self.model().item(i).data())
        return res


class CustomComboBox(QComboBox):

    # Subclass Delegate to increase item height
    class Delegate(QStyledItemDelegate):
        def sizeHint(self, option, index):
            size = super().sizeHint(option, index)
            size.setHeight(25)
            return size


    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Make the combo editable to set a custom text, but readonly
        self.setEditable(True)
        self.lineEdit().setReadOnly(True)
        # Make the lineedit the same color as QPushButton
        palette = QtWidgets.QApplication.palette()
        palette.setBrush(QtGui.QPalette.Base, palette.button())
        self.lineEdit().setPalette(palette)

        # Use custom delegate
        self.setItemDelegate(CustomComboBox.Delegate())

        # Hide and show popup when clicking the line edit
        self.lineEdit().installEventFilter(self)
        self.closeOnLineEditClick = False


    def resizeEvent(self, event):
        super().resizeEvent(event)


    def eventFilter(self, object, event):

        if object == self.lineEdit():
            if event.type() == QEvent.MouseButtonRelease:
                if self.closeOnLineEditClick:
                    self.hidePopup()
                else:
                    self.showPopup()
                return True
            return False

        if object == self.view().viewport():
            if event.type() == QEvent.MouseButtonRelease:
                index = self.view().indexAt(event.pos())
                item = self.model().item(index.row())

                if item.checkState() == Qt.Checked:
                    item.setCheckState(Qt.Unchecked)
                else:
                    item.setCheckState(Qt.Checked)
                return True
        return False


    def showPopup(self):
        super().showPopup()
        # When the popup is displayed, a click on the lineedit should close it
        self.closeOnLineEditClick = True


    def hidePopup(self):
        super().hidePopup()
        # Used to prevent immediate reopening when clicking on the lineEdit
        self.startTimer(100)


    def timerEvent(self, event):
        # After timeout, kill timer, and reenable click on line edit
        self.killTimer(event.timerId())
        self.closeOnLineEditClick = False


class CustomMessageBox(QMessageBox):
    def __init__(self):
        QtWidgets.QMessageBox.__init__(self)
        self.setSizeGripEnabled(True)

    def event(self, e):
        result = QtWidgets.QMessageBox.event(self, e)

        self.setMinimumHeight(0)
        self.setMaximumHeight(16777215)
        self.setMinimumWidth(0)
        self.setMaximumWidth(16777215)
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)

        textEdit = self.findChild(QtWidgets.QTextEdit)
        if textEdit != None :
            textEdit.setMinimumHeight(0)
            textEdit.setMaximumHeight(16777215)
            textEdit.setMinimumWidth(0)
            textEdit.setMaximumWidth(16777215)
            textEdit.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)

        return result


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        loadUi('DevSupGUI.ui', self)
        
        # Source Files
        self.source_path = r'C:\Users'
        
            # Filename EditLines
        self.filename_dm.setPlaceholderText("File path")
        self.filename_dms.setPlaceholderText("File path")
        self.filename_rfic.setPlaceholderText("File path")
        self.filename_spec.setPlaceholderText("File path")
        
            # Browse Buttons
        self.browse_dm_file.clicked.connect(self.browse_dm)
        self.browse_dms_file.clicked.connect(self.browse_dms)
        self.browse_rfic_file.clicked.connect(self.browse_rfic)
        self.browse_spec_file.clicked.connect(self.browse_spec)
        
            # Upload Buttons
        self.upload_btn_dm_file.clicked.connect(self.upload_dm)
        self.upload_btn_dms_file.clicked.connect(self.upload_dms)
        self.upload_btn_rfic_file.clicked.connect(self.upload_rfic)
        self.upload_btn_spec_file.clicked.connect(self.upload_spec)
        
            # Upload Checkboxes
        self.upload_cb_dm_file.setEnabled(False)
        self.upload_cb_dms_file.setEnabled(False)
        self.upload_cb_rfic_file.setEnabled(False)
        self.upload_cb_spec_file.setEnabled(False)
        
        self.upload_cb_dm_file.setText('Pending')
        self.upload_cb_dms_file.setText('Pending')
        self.upload_cb_rfic_file.setText('Pending')
        self.upload_cb_spec_file.setText('Pending')
        
            # Database
        self.create_cb_db.setEnabled(False)
        self.create_cb_db.setText('Pending')
        self.create_btn_db.clicked.connect(self.create_database)
        
        
        # Parameters
        self.params_layout.setEnabled(False)
        
            # Jurisdictions
        self.jdxs = {'Austria': 'at', 'Belgium': 'be', 'Switzerland': 'ch', 'Germany': 'de',
                     'Denmark': 'dk', 'Spain': 'es', 'FINREP': 'fin', 'France': 'fr',
                     'Great Britain': 'gb', 'Ireland':'ie', 'Italy': 'it', 'Luxemburg': 'lu',
                     'Netherlands': 'nl', 'Sweden': 'se'}
        
        self.comb_jdx = CustomComboBox()
        self.comb_jdx.addItems(['Select...'] + list(self.jdxs.keys()))
        self.gridL_Params.addWidget(self.comb_jdx, 0, 1)
        self.comb_jdx.setCurrentText('Select...')
        
            #Tab Names
        self.ccombox_tab_names = CheckableComboBox()
        self.gridL_Params.addWidget(self.ccombox_tab_names, 1, 1)
        
        
        # Functions
        self.funcs_layout.setEnabled(False)
        
            # Lookups
        self.create_btn_lkups.clicked.connect(self.create_lookups)
        
        self.view_btn_lkups.clicked.connect(self.show_lkups)
        self.save_btn_lkups.clicked.connect(self.save_lkups)
        
            # Overlaps
        self.analyze_btn_overlaps.clicked.connect(self.analyze_overlaps)
        self.comb_jdx.activated.connect(self.activate_func_gb)
        self.ccombox_tab_names.currentTextChanged.connect(self.activate_func_gb)
        
        self.view_btn_overlaps.clicked.connect(self.show_overlaps)
        
        self.dm_path_prev = None
        self.dms_path_prev = None
        
    
    def upload_dm(self):
        """Data Model Load"""
        # save prev dm path
        if hasattr(self, 'dm_path'):
            self.dm_path_prev = self.dm_path

        # dm file path
        fpath = self.filename_dm.text().strip()
        
        # if dm already exists
        if hasattr(self, 'dm_path'):
            dm_indicator = True if fpath != self.dm_path else False
        else:
            dm_indicator = True
        
        # upload checks block
        if dm_indicator:
            if fpath:                                                                   # fpath must be not null
                if exists(fpath):                                                       # file must exists
                    # dm upload and save the path
                    self.data_model = pd.ExcelFile(fpath)
                    self.dm_path = fpath
                    
                    # dm checkbox - success
                    self.upload_cb_dm_file.setChecked(True)
                    self.upload_cb_dm_file.setStyleSheet('color: green')
                    self.upload_cb_dm_file.setText('Success')
                    
                    # dm button disable
                    self.upload_btn_dm_file.setEnabled(False)
                    
                    print(f'Data Model file uploaded by path: {self.dm_path}')
                    
                    # if dms exists, enable db to create
                    if hasattr(self, 'data_model_specialist'):
                        self.label_db.setEnabled(True)
                        self.create_btn_db.setEnabled(True)
                        self.create_cb_db.setText('Pending')
                        
                else:
                    # warning if file not extists
                    directory = '/'.join(fpath.split('/')[:-1])
                    file = fpath.split('/')[-1]
                    self.show_warning(text=f'No such file {file} in directory {directory}.')
                    return
            else:
                self.show_warning(text='Please ensure Data Model file path is specified.')
                return
    
 
    def upload_dms(self):
        """Data Model Specialist Load"""
        
        # save prev dm path
        if hasattr(self, 'dms_path'):
            self.dms_path_prev = self.dms_path
        
        # dms file path
        fpath = self.filename_dms.text().strip()
        
        # if dms already exists
        if hasattr(self, 'dms_path'):
            dms_indicator = True if fpath != self.dms_path else False
        else:
            dms_indicator = True
        
        # upload checks block
        if dms_indicator:
            if fpath:                                                                   # fpath must be not null
                if exists(fpath):                                                       # file must exists  
                    # dms upload and save the path
                    self.data_model_specialist = pd.ExcelFile(fpath)
                    self.dms_path = fpath
                    
                    # dms checkbox - success
                    self.upload_cb_dms_file.setChecked(True)
                    self.upload_cb_dms_file.setStyleSheet('color: green')
                    self.upload_cb_dms_file.setText('Success')
                    
                    # dms button disable
                    self.upload_btn_dms_file.setEnabled(False)
                    
                    print(f'Data Model Specialist file uploaded by path: {self.dms_path}')
                    
                    # if dm exists, enable db to create
                    if hasattr(self, 'data_model'):
                        self.label_db.setEnabled(True)
                        self.create_btn_db.setEnabled(True)
                        self.create_cb_db.setText('Pending')
                       
                else:
                    # warning if file not extists
                    directory = '/'.join(fpath.split('/')[:-1])
                    file = fpath.split('/')[-1]
                    self.show_warning(text=f'No such file {file} in directory {directory}.')
                    return
            else:
                self.show_warning(text='Please ensure Data Model Specialist file path is specified.')
                return


    def upload_rfic(self):
        """Rules For Interval Calculation Load"""
        
        # rfic file path
        fpath = self.filename_rfic.text().strip()
        
        # if rfic already exists
        if hasattr(self, 'rfic_path'):
            rfic_indicator = True if fpath != self.rfic_path else False
        else:
            rfic_indicator = True
        
        # upload checks block
        if rfic_indicator:
            if fpath:                                                                   # fpath must be not null
                if exists(fpath):                                                       # file must exists
                    # rfic upload and save the path
                    self.int_cals = pd.ExcelFile(fpath)
                    self.rfic_path = fpath
                    
                    # rfic checkbox - success
                    self.upload_cb_rfic_file.setChecked(True)
                    self.upload_cb_rfic_file.setStyleSheet('color: green')
                    self.upload_cb_rfic_file.setText('Success')
                    
                    # rfic button disable
                    self.upload_btn_rfic_file.setEnabled(False)
                    
                    print(f'Rules For Interval Calculation file uploaded by path: {self.rfic_path}')
                else:
                    # warning if file not extists
                    directory = '/'.join(fpath.split('/')[:-1])
                    file = fpath.split('/')[-1]
                    self.show_warning(text=f'No such file {file} in directory {directory}.')
                    return
            else:
                self.show_warning(text='Please ensure Rules For Interval Calculation file path is specified.')
                return


    def upload_spec(self):
        """Specification Load"""
        
        # spec file path
        fpath = self.filename_spec.text().strip()
        
        # if spec already exists
        if hasattr(self, 'spec_path'):
            spec_indicator = True if fpath != self.spec_path else False
        else:
            spec_indicator = True
        
        # upload checks block
        if spec_indicator:
            if fpath:                                                                   # fpath must be not null
                if exists(fpath):                                                       # file must exists
                    # spec upload, set and save the path
                    self.spec = pd.ExcelFile(fpath)
                    self.set_spec()
                    self.spec_path = fpath
                    
                    # enable layout of parameters
                    self.params_layout.setEnabled(True)
                    
                    # spec checkbox - success
                    self.upload_cb_spec_file.setChecked(True)
                    self.upload_cb_spec_file.setStyleSheet('color: green')
                    self.upload_cb_spec_file.setText('Success')
                    
                    # spec button disable
                    self.upload_btn_spec_file.setEnabled(False)
                    print(f'Specification file uploaded by path: {self.spec_path}')
                else:
                    # warning if file not extists
                    directory = '/'.join(fpath.split('/')[:-1])
                    file = fpath.split('/')[-1]
                    self.show_warning(text=f'No such file {file} in directory {directory}.')
                    return
            else:
                self.show_warning(text='Please ensure Specification file path is specified.')
                return


    def create_database(self):
        try:
            # db creation
            self.tables, self.database = functions.load_model(self.data_model, self.data_model_specialist)
            self.show_info('Database', 'Database Successfully Created')
            print('Database created')
            
            # db button disable
            self.create_btn_db.setEnabled(False)
            
            # db checkbox
            self.create_cb_db.setChecked(True)
            self.create_cb_db.setStyleSheet('color: green')
            self.create_cb_db.setText('Success')
            
            
            self.activate_func_gb()
        except Exception as e:
            self.show_error('Database Creation', str(e))
            
            # delete incorrect vars
            if hasattr(self, 'tables'):
                del self.tables
            
            if hasattr(self, 'database'):
                del self.database


    def browse_dm(self):
        # browse dm file
        fname = QFileDialog.getOpenFileName(self, 'Open File', self.source_path, 'Excel (*.xls *.xlsx)')
        fpath = fname[0].strip()
        
        # enable dm button if path is not null
        if fpath:
            self.upload_btn_dm_file.setEnabled(True)
        
        # dm check block
        if hasattr(self, 'dm_path'):                                                # if dm already exists
            if self.dm_path == fpath:                                               # if new path equals prev path
                # disable dm button
                self.upload_btn_dm_file.setEnabled(False)
                
                # dm checkbox - success
                self.upload_cb_dm_file.setChecked(True)
                self.upload_cb_dm_file.setStyleSheet('color: green')
                self.upload_cb_dm_file.setText('Success')

                if hasattr(self, 'tables') and hasattr(self, 'database'):           # if db exists
                    # db - successful creation
                    self.label_db.setEnabled(True)
                    self.create_btn_db.setEnabled(False)
                    self.create_cb_db.setChecked(True)
                    self.create_cb_db.setStyleSheet('color: green')
                    self.create_cb_db.setText('Success')
            else:
                # disable overlaps btn
                self.analyze_btn_overlaps.setEnabled(False)
                
                # overlaps checkbox - disable
                self.analyze_cb_overlaps.setChecked(False)
                self.analyze_cb_overlaps.setStyleSheet('color: grey')
                self.analyze_cb_overlaps.setText('Pending')
                
                # overlaps view, save buttons - disable
                self.view_btn_overlaps.setEnabled(False)
                self.save_btn_overlaps.setEnabled(False)
                
                # dm checkbox - disable
                self.upload_cb_dm_file.setChecked(False)
                self.upload_cb_dm_file.setStyleSheet('color: grey')
                self.upload_cb_dm_file.setText('Pending')
                
                # db - disable  
                self.label_db.setEnabled(False)
                self.create_btn_db.setEnabled(False)
                self.create_cb_db.setChecked(False)
                self.create_cb_db.setStyleSheet('color: grey')
                self.create_cb_db.setText('Pending')

        self.activate_func_gb()

        # set new file path
        self.filename_dm.setText(fpath)
        self.source_path = '/'.join(fpath.split('/')[:-1])


    def browse_dms(self):
        # browse dms file
        fname = QFileDialog.getOpenFileName(self, 'Open File', self.source_path, 'Excel (*.xls *.xlsx)')
        fpath = fname[0].strip()
        
        
        # enable dms button if path is not null
        if fpath:   
            self.upload_btn_dms_file.setEnabled(True)
        
        # dms check block
        if hasattr(self, 'dms_path'):                                               # if dms already exists
            if self.dms_path == fpath:                                              # if new path equals prev path
                # disable dms button
                self.upload_btn_dms_file.setEnabled(False)
                
                # dms checkbox - success
                self.upload_cb_dms_file.setChecked(True)
                self.upload_cb_dms_file.setStyleSheet('color: green')
                self.upload_cb_dms_file.setText('Success')
                
                if hasattr(self, 'tables') and hasattr(self, 'database'):           # if db exists
                    # db - successful creation
                    self.label_db.setEnabled(True)
                    self.create_btn_db.setEnabled(False)
                    self.create_cb_db.setChecked(True)
                    self.create_cb_db.setStyleSheet('color: green')
                    self.create_cb_db.setText('Success')
            else:
                # disable overlaps btn
                self.analyze_btn_overlaps.setEnabled(False)
                
                # overlaps checkbox - disable
                self.analyze_cb_overlaps.setChecked(False)
                self.analyze_cb_overlaps.setStyleSheet('color: grey')
                self.analyze_cb_overlaps.setText('Pending')
                
                # overlaps view, save buttons - disable
                self.view_btn_overlaps.setEnabled(False)
                self.save_btn_overlaps.setEnabled(False)
                
                # dms checkbox - disable
                self.upload_cb_dms_file.setChecked(False)
                self.upload_cb_dms_file.setStyleSheet('color: grey')
                self.upload_cb_dms_file.setText('Pending')
                
                # db - disable
                self.label_db.setEnabled(False)
                self.create_btn_db.setEnabled(False)
                self.create_cb_db.setChecked(False)
                self.create_cb_db.setStyleSheet('color: grey')
                self.create_cb_db.setText('Pending')
        
        self.activate_func_gb()
        
        # set new file path
        self.filename_dms.setText(fpath)
        self.source_path = '/'.join(fpath.split('/')[:-1])


    def browse_rfic(self):
        # browse rfic file
        fname = QFileDialog.getOpenFileName(self, 'Open File', self.source_path, 'Excel (*.xls *.xlsx)')
        fpath = fname[0].strip()
        
        # enable rfic button if path is not null
        if fpath:   
            self.upload_btn_rfic_file.setEnabled(True)
        
        # rfic check block
        if hasattr(self, 'rfic_path'):                                              # if rfic already exists
            if self.rfic_path == fpath:                                             # if new path equals prev path
                # disable rfic button
                self.upload_btn_rfic_file.setEnabled(False)
                
                # rfic checkbox - success
                self.upload_cb_rfic_file.setChecked(True)
                self.upload_cb_rfic_file.setStyleSheet('color: green')
                self.upload_cb_rfic_file.setText('Success')
            else:
                # disable overlaps btn
                self.analyze_btn_overlaps.setEnabled(False)
                
                # overlaps checkbox - disable
                self.analyze_cb_overlaps.setChecked(False)
                self.analyze_cb_overlaps.setStyleSheet('color: grey')
                self.analyze_cb_overlaps.setText('Pending')
            
                # dms checkbox - disable
                self.upload_cb_rfic_file.setChecked(False)
                self.upload_cb_rfic_file.setStyleSheet('color: grey')
                self.upload_cb_rfic_file.setText('Pending')
        
        self.activate_func_gb()
        
        # set new file path
        self.filename_rfic.setText(fpath)
        self.source_path = '/'.join(fpath.split('/')[:-1])


    def browse_spec(self):
        # browse spec file
        fname = QFileDialog.getOpenFileName(self, 'Open File', self.source_path, 'Excel (*.xls *.xlsx)')
        fpath = fname[0].strip()
        
        # enable spec button if path is not null
        if fpath:   
            self.upload_btn_spec_file.setEnabled(True)
        
        # spec check block
        if hasattr(self, 'spec_path'):                                              # if spec already exists
            if self.spec_path == fpath:                                             # if new path equals prev path
                # disable spec button
                self.upload_btn_spec_file.setEnabled(False)
                
                # enable layout of parameters
                self.params_layout.setEnabled(True)
                
                # spec checkbox - success
                self.upload_cb_spec_file.setChecked(True)
                self.upload_cb_spec_file.setStyleSheet('color: green')
                self.upload_cb_spec_file.setText('Success')
            else:
                # enable layout of parameters
                self.params_layout.setEnabled(False)
                
                # spec checkbox - disable
                self.ccombox_tab_names.setCurrentText('Select...')
                self.upload_cb_spec_file.setChecked(False)
                self.upload_cb_spec_file.setStyleSheet('color: grey')
                self.upload_cb_spec_file.setText('Pending')
        
        self.activate_func_gb()
        
        # set new file path
        self.filename_spec.setText(fpath)
        self.source_path = '/'.join(fpath.split('/')[:-1])


    def set_spec(self):
        # set tabs
        self.tabs = self.spec.sheet_names
        self.tab_count = len(self.tabs)
        
        # Tab Names
        if self.ccombox_tab_names:
            self.ccombox_tab_names.clear()
        self.ccombox_tab_names.addItems(self.tabs)
        self.ccombox_tab_names.setCurrentText('Select...')
    

    def activate_func_gb(self):
        # enable/disable functions layout
        if self.ccombox_tab_names.currentText() and self.ccombox_tab_names.currentText() != 'Select...':            # if something selected in tab names combobox
            # enable functions layout and lkups button
            self.funcs_layout.setEnabled(True)
            self.create_btn_lkups.setEnabled(True)
            
            if (str(self.comb_jdx.currentText()) != 'Select...' and hasattr(self, 'tables') and
                    hasattr(self, 'database') and hasattr(self, 'int_cals')):                                       # if jdx selected and db, rfic uploaded          
                
                # enable overlaps button
                self.analyze_btn_overlaps.setEnabled(True)
                
                # overlaps checkbox - enable
                if hasattr(self, 'jdx_prev'):
                    if (self.jdxs[str(self.comb_jdx.currentText())] == self.jdx_prev and self.ccombox_tab_names.currentData() == self.tab_names_prev and
                            self.dm_path_prev == self.dm_path and self.dms_path_prev == self.dms_path):
                        # overlaps checkbox - enable
                        self.analyze_cb_overlaps.setChecked(True)
                        self.analyze_cb_overlaps.setStyleSheet('color: green')
                        self.analyze_cb_overlaps.setText('Success')
                        
                        # disable overlaps button
                        self.analyze_btn_overlaps.setEnabled(False)
                        
                        # overlaps view, save buttons - disable
                        self.view_btn_overlaps.setEnabled(True)
                        self.save_btn_overlaps.setEnabled(True)
                    else:
                        # overlaps checkbox - disable
                        self.analyze_cb_overlaps.setChecked(False)
                        self.analyze_cb_overlaps.setStyleSheet('color: grey')
                        self.analyze_cb_overlaps.setText('Pending')
                        
                        # disable overlaps button
                        if self.filename_dm.text().strip() and self.filename_dms.text().strip():
                            self.analyze_btn_overlaps.setEnabled(True)
                        else:
                            self.analyze_btn_overlaps.setEnabled(False)
                        
                        # overlaps view, save buttons - disable
                        self.view_btn_overlaps.setEnabled(False)
                        self.save_btn_overlaps.setEnabled(False)
            else:
                # overlaps checkbox - disable
                self.analyze_cb_overlaps.setChecked(False)
                self.analyze_cb_overlaps.setStyleSheet('color: grey')
                self.analyze_cb_overlaps.setText('Pending')
                
                # disable overlaps button
                self.analyze_btn_overlaps.setEnabled(False)
                
                # overlaps view, save buttons - disable
                self.view_btn_overlaps.setEnabled(False)
                self.save_btn_overlaps.setEnabled(False)
        
        else:
            # disable functions layout and lkups button
            self.funcs_layout.setEnabled(False)
            self.create_btn_lkups.setEnabled(False)
            
            # overlaps checkbox - disable
            self.analyze_cb_overlaps.setChecked(False)
            self.analyze_cb_overlaps.setStyleSheet('color: grey')
            self.analyze_cb_overlaps.setText('Pending')
            
            # disable overlaps button
            self.analyze_btn_overlaps.setEnabled(False)
            
            # overlaps view, save buttons - disable
            self.view_btn_overlaps.setEnabled(False)
            self.save_btn_overlaps.setEnabled(False)


    def analyze_overlaps(self):
        # check jdx
        try:
            jdx = str(self.comb_jdx.currentText())
            self.jdx = self.jdxs[jdx]
        except Exception as e:
            self.show_error('Jurisdiction Error', str(e))
        
        # get selected tabs
        self.tab_names = self.ccombox_tab_names.currentData()
        
        # try to analyze overlaps
        try:
            self.overlaps = functions.main(self.tables, self.database, self.spec, self.int_cals, self.jdx, self.tab_names)
            
            # prev params to check in future
            self.jdx_prev = self.jdx
            self.tab_names_prev = self.tab_names
            
            # disable overlaps button
            self.analyze_btn_overlaps.setEnabled(False)
            
            # overlaps checkbox - disable
            self.analyze_cb_overlaps.setChecked(True)
            self.analyze_cb_overlaps.setStyleSheet('color: green')
            self.analyze_cb_overlaps.setText('Success')
            
            # enable view and save buttons
            self.view_btn_overlaps.setEnabled(True)
            self.save_btn_overlaps.setEnabled(True)
        except Exception as e:
            self.show_error('Analyze Overlaps Error', str(e))

    
    def show_warning(self, title=None, text=None, inform_text=None):
        warn = CustomMessageBox()
        warn.setIcon(QMessageBox.Warning)
        
        if title:
            warn.setWindowTitle(title)
        
        if text:
            warn.setText(text)
        
        if inform_text:
            warn.setInformativeText(inform_text)
        
        warn.setStandardButtons(QMessageBox.Ok)
        warn.exec_()


    def show_error(self, title=None, text=None, inform_text=None):
        err = CustomMessageBox()
        err.setIcon(QMessageBox.Critical)
        
        if title:
            err.setWindowTitle(title)
        
        if text:
            err.setText(text)
        
        if inform_text:
            err.setInformativeText(inform_text)
        
        err.setStandardButtons(QMessageBox.Ok)
        err.exec_()


    def show_info(self, title=None, text=None, inform_text=None):
        info = CustomMessageBox()
        info.setIcon(QMessageBox.Information)
        
        if title:
            info.setWindowTitle(title)
        
        if text:
            info.setText(text)
        
        if inform_text:
            info.setInformativeText(inform_text)
        
        info.setStandardButtons(QMessageBox.Ok)
        info.exec_()


    def show_detailed_info(self, title=None, text=None, inform_text=None, detailed_text=None):
        info = CustomMessageBox()
        info.setIcon(QMessageBox.Information)
        
        if title:
            info.setWindowTitle(title)
        
        if text:
            info.setText(text)
        
        if inform_text:
            info.setInformativeText(inform_text)
        
        if detailed_text:
            info.setDetailedText(detailed_text)
        
        info.setStandardButtons(QMessageBox.Ok)
        info.exec_()
 

    def show_overlaps(self):
        detailed_text = ''
        
        for tab_name, tab in self.overlaps.items():
            
            detailed_text += f'{tab_name} Tab'
            
            for form_name, form in tab.items():
                ols, tots = form
                
                ols_ind = False
                for k, v in ols.items():
                    if v:
                        ols_ind = True
                        break
                                    
                if ols_ind:
                    detailed_text += f"\n\n\t{form_name}' overlaps:\n"
                    print(f"\n\t{form_name}' overlaps:\n")
                    
                    for table, ovlps in ols.items():
                        if ovlps:
                            detailed_text += '\t\t' + ' '.join(list(map(lambda s: s.capitalize(), table.replace('_', ' ').split()))) + ':' + '\n'
                            print('\t\t', ' '.join(list(map(lambda s: s.capitalize(), table.replace('_', ' ').split()))), ':', sep='', end='\n')
                            for item, fields in ovlps.items():
                                detailed_text += '\t\t\t' + str(item) + '\n'
                                print(item)
                            detailed_text += '\n'    
                            print()
                else:
                    detailed_text += f'\n\tNO OVERLAPS IN {form_name}!'
                    print(f'\n\tNO OVERLAPS IN {form_name}!')
                
                                    
                if tots:
                    detailed_text += f"\n\n\t{form_name}' totals:\n"
                    print(f"\n\t{form_name}' totals:\n")
                                
                    for item in tots:
                        detailed_text += '\t\t' + str(item) + '\n'
                        print(item)
                        
        
        
        title = 'Overlaps'
        text = 'Overlaps successfully analyzed with the following parameters: ' + f'Jurisdiction - {self.comb_jdx.currentText()}; ' \
                            + ('Tab' if len(self.tab_names) == 1 else 'Tabs') + f': ' + ', '.join(self.tab_names) + '.'
        inform_text = 'See results in detailes below.'
        
        self.show_detailed_info(title, text, inform_text, detailed_text)


    def create_lookups(self):
        self.tab_names = self.ccombox_tab_names.currentData()
        
        try:
            report_tabs = functions.load_spec(self.spec, self.tab_names)
        except Exception as e:
            self.show_error('Load Specification Error', str(e))
        
        try:
            self.lkups = lookups.collect_lkups(report_tabs, self.tab_names)
            
            # disable lookups button
            self.create_btn_lkups.setEnabled(False)
            
            # overlaps checkbox - disable
            self.create_cb_lkups.setChecked(True)
            self.create_cb_lkups.setStyleSheet('color: green')
            self.create_cb_lkups.setText('Success')
            
            # enable view and save buttons
            self.view_btn_lkups.setEnabled(True)
            self.save_btn_lkups.setEnabled(True)
            
        except Exception as e:
            self.show_error('Create Lookups Error', str(e))
            

    def show_lkups(self):
        detailed_text = ''
        
        for tab_name, tab in self.lkups.items():
            detailed_text += tab_name + ':\n'
            
            for form_name, form in tab.items():
                detailed_text += '\t' + form_name + ':\n'
                
                for index, row in form.iterrows():
                    detailed_text += '\t\t' + row[0] + '\t' + row[1] + '\t' + row[2] + '\t' + row[3] + '\n'
        
                detailed_text += '\n'
            
            detailed_text += '\n\n'
        
        
        title = 'Lookups'
        text = 'Lookups successfully created for the following ' + \
                        ('tab' if len(self.tab_names) == 1 else 'tabs') + f': ' + ', '.join(self.tab_names) + '.'
        inform_text = 'See results in detailes below.'
        
        self.show_detailed_info(title, text, inform_text, detailed_text)
        
   
    def save_lkups(self):
        default_filename = join(self.source_path, 'untitled.xlsx')
        filename, _ = QFileDialog.getSaveFileName(self, 'Save File', default_filename, 'Excel (*.xls *.xlsx)')
        
        if filename:
            Excelwriter = pd.ExcelWriter(filename, engine="xlsxwriter")

            for tab_name, tab in self.lkups.items():
                for form_name, form in tab.items():
                    sheet_name = tab_name + '_' + form_name
                    form.to_excel(Excelwriter, sheet_name=sheet_name, index=False)
                    
            Excelwriter.save()
        
            
if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())