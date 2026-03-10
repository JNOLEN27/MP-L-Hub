import pandas as pd
from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTabWidget, QTextEdit, QTableWidget, QTableWidgetItem, QMessageBox, QScrollArea, QFileDialog, QComboBox, QListWidget, QListWidgetItem, QCheckBox, QFrame, QApplication, QLineEdit, QGridLayout, QProgressDialog 
from PyQt5.QtCore import Qt, pyqtSignal, QTimer
from PyQt5.QtGui import QFont
from datetime import datetime, timedelta

from app.utils.config import APPWINDOWSIZE
from app.data.import_manager import DataImportManager
from app.supply_chain_coordination.coverage_analysis import CoverageAnalysisEngine
from app.supply_chain_coordination.waterfall_analysis import WaterfallAnalysisEngine

class SupplyChainCoordinationWindow(QMainWindow):
    def __init__(self, userdata, parent=None):
        super().__init__(parent)
        self.userdata = userdata
        self.import_manager = DataImportManager()
        self.coverageengine = CoverageAnalysisEngine(self.import_manager)
        self.waterfallengine = WaterfallAnalysisEngine(self.import_manager)
        self.setWindowTitle("Supply Chain Coordination Application")
        self.resize(*APPWINDOWSIZE)
        self.setupui()
        
    def setupui(self):
        centralwidget = QWidget()
        self.setCentralWidget(centralwidget)
        
        layout = QVBoxLayout()
        
        title = QLabel("Supply Chain Coordination")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        tabs = QTabWidget()
        
        coveragedashtab = self.createcoveragedashboard()
        coverageindivtab = self.createindividualpartcoverage()
        calloffforecasttab = self.createcalloffforecasttab()
        
        tabs.addTab(coveragedashtab, "Coverage Dashboard")
        tabs.addTab(coverageindivtab, "Individual Part Coverage")
        tabs.addTab(calloffforecasttab, "Call-off Forecast and Waterfall")
        
        layout.addWidget(tabs)
        
        self.statusBar().showMessage(f"Logged in as: {self.userdata['username']}")
        centralwidget.setLayout(layout)
        
    def createfiltersection(self):
        widget = QWidget()
        widget.setMaximumHeight(190)
        widget.setVisible(False)
        
        layout = QHBoxLayout()
        
        filterlabel = QLabel("Filters:")
        filterlabel.setFont(QFont("Arial", 12, QFont.Bold))
        filterlabel.setAlignment(Qt.AlignTop)
        layout.addWidget(filterlabel)
        
        self.sccfilter = self.createmultiselectdropdown("Select SCC Names...")
        self.sccfilter.selectionChanged.connect(self.applyfilters)
        layout.addWidget(self.sccfilter)
        
        self.regionfilter = self.createmultiselectdropdown("Select Regions...", "Region")
        self.regionfilter.selectionChanged.connect(self.applyfilters)
        layout.addWidget(self.regionfilter)
        
        self.dayalertfilter = self.createmultiselectdropdown("Select Day Alert...", "Day Alert")
        self.dayalertfilter.selectionChanged.connect(self.applyfilters)
        layout.addWidget(self.dayalertfilter)
        
        searchgridwidget = QWidget()
        searchgridlayout = QGridLayout(searchgridwidget)
        searchgridlayout.setSpacing(5)
        
        self.searchfilters = []
        
        partfilter = self.createsearchfilter("Part", "PART_NO")
        self.searchfilters.append(partfilter)
        searchgridlayout.addWidget(partfilter, 0, 0)
        
        mfgfilter = self.createsearchfilter("MFG", "SUPP_MFG")
        self.searchfilters.append(mfgfilter)
        searchgridlayout.addWidget(mfgfilter, 0, 1)
        
        shpfilter = self.createsearchfilter("SHP", "SUPP_SHP")
        self.searchfilters.append(shpfilter)
        searchgridlayout.addWidget(shpfilter, 1, 0)
        
        countryfilter = self.createsearchfilter("Country", "SUPP_SHP_COUNTRY")
        self.searchfilters.append(countryfilter)
        searchgridlayout.addWidget(countryfilter, 1, 1)
        
        layout.addWidget(searchgridwidget)
        
        clearfiltersbtn = QPushButton("Clear All Filters")
        clearfiltersbtn.clicked.connect(self.clearfilters)
        clearfiltersbtn.setMaximumWidth(120)
        clearfiltersbtn.setMaximumHeight(30)
        layout.addWidget(clearfiltersbtn)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget
    
    def createmultiselectdropdown(self, placeholdertext, filtertype = "SCC"):
        class SimpleMultiSelectFilter(QWidget):
            selectionChanged = pyqtSignal()
            
            def __init__(self, placeholder="Select items...", filtertype = "SCC"):
                super().__init__()
                self.filtertype = filtertype
                self.setFixedHeight(170)
                
                layout = QVBoxLayout()
                layout.setContentsMargins(5, 5, 5, 5)
                
                self.label = QLabel(f"{filtertype} Filter:")
                self.label.setFont(QFont("Arial", 10, QFont.Bold))
                layout.addWidget(self.label)
                
                self.list_widget = QListWidget()
                self.list_widget.setMaximumHeight(150)
                self.list_widget.itemChanged.connect(self.on_item_changed)
                self.list_widget.setStyleSheet("""
                    QListWidget {
                        border: 1px solid #ccc;
                        font-size: 11px;
                    }
                    QListWidget::item {
                        padding: 2px;
                    }
                """)
                layout.addWidget(self.list_widget)
                
                self.setLayout(layout)
                self.selected_items = set()
                self.all_items = []
                
            def additems(self, items, presorted = False):
                self.list_widget.clear()
                self.selected_items.clear()
                self.all_items.clear()
                
                select_all_item = QListWidgetItem(f"✓ All {self.filtertype}'s")
                select_all_item.setFlags(select_all_item.flags() | Qt.ItemIsUserCheckable)
                select_all_item.setCheckState(Qt.Checked)
                select_all_item.setData(Qt.UserRole, "SELECT_ALL")
                select_all_item.setFont(QFont("Arial", 9, QFont.Bold))
                self.list_widget.addItem(select_all_item)
                
                maxtextlength = len("✓ All SCC Names")
                
                items_to_use = items if presorted else sorted(items)
                
                for item in items_to_use:
                    if item and str(item).strip():
                        list_item = QListWidgetItem(str(item))
                        list_item.setFlags(list_item.flags() | Qt.ItemIsUserCheckable)
                        list_item.setCheckState(Qt.Checked)
                        self.list_widget.addItem(list_item)
                        self.selected_items.add(str(item))
                        maxtextlength = max(maxtextlength, len(str(item)))
                
                self.autosizewidth(maxtextlength)
                self.update_label()
                
            def autosizewidth(self, maxtextlength):
                fontmetrics = self.list_widget.fontMetrics()
                estimatedwidth = fontmetrics.averageCharWidth() * maxtextlength + 60
                
                if self.parent():
                    parentwidth = self.parent().width()
                    availablewidth = parentwidth - 280
                else:
                    availablewidth = 400
                    
                minwidth = 150
                maxwidth = min(400, availablewidth)
                optimalwidth = max(minwidth, min(estimatedwidth, maxwidth))
                
                self.setFixedWidth(optimalwidth)
                
            def on_item_changed(self, item):
                self.list_widget.blockSignals(True)
                
                try:
                    if item.data(Qt.UserRole) == "SELECT_ALL":
                        if item.checkState() == Qt.Checked:
                            for i in range(1, self.list_widget.count()):
                                other_item = self.list_widget.item(i)
                                other_item.setCheckState(Qt.Checked)
                                self.selected_items.add(other_item.text())
                        else:
                            for i in range(1, self.list_widget.count()):
                                other_item = self.list_widget.item(i)
                                other_item.setCheckState(Qt.Unchecked)
                            self.selected_items.clear()
                    else:
                        if item.checkState() == Qt.Checked:
                            self.selected_items.add(item.text())
                        else:
                            self.selected_items.discard(item.text())
                            select_all_item = self.list_widget.item(0)
                            select_all_item.setCheckState(Qt.Unchecked)
                        
                        if len(self.selected_items) == len(self.all_items):
                            select_all_item = self.list_widget.item(0)
                            select_all_item.setCheckState(Qt.Checked)
                    
                finally:
                    self.list_widget.blockSignals(False)
                
                self.update_label()
                self.selectionChanged.emit()
                
            def show_progress_for_large_operation(self):
                
                progress = QProgressDialog("Loading full dataset...", "Cancel", 0, 100, self)
                progress.setWindowTitle("Processing")
                progress.setWindowModality(Qt.WindowModal)
                progress.setValue(50)
                progress.show()
                
                QApplication.processEvents()
                
                return progress
                
            def update_label(self):
                if len(self.selected_items) == 0:
                    self.label.setText(f"{self.filtertype} Filter:")
                elif len(self.selected_items) == len(self.all_items):
                    self.label.setText(f"{self.filtertype} Filter:")
                else:
                    self.label.setText(f"{self.filtertype} Filter:")
                    
            def getselecteditems(self):
                return list(self.selected_items)
                
            def selectallitems(self):
                if len(self.all_items) > 50:
                    progress = self.parent().show_progress_for_large_operation()
                else:
                    progress = None
                
                try:
                    self.list_widget.blockSignals(True)
                    
                    self.selected_items = set(self.all_items)
                    
                    for i in range(self.list_widget.count()):
                        item = self.list_widget.item(i)
                        item.setCheckState(Qt.Checked)
                        
                finally:
                    self.list_widget.blockSignals(False)
                    if progress:
                        progress.close()
                
                self.update_label()
        
        return SimpleMultiSelectFilter(placeholdertext, filtertype)
    
    def createsearchfilter(self, filtertype="MFG", columnname="SUPP_MFG"):
        widget = QWidget()
        widget.setFixedHeight(120)
        widget.setFixedWidth(180)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(3, 3, 3, 3)
        layout.addSpacing(2)
        
        label = QLabel(f"{filtertype} Search")
        label.setFont(QFont("Arial", 9, QFont.Bold))
        layout.addWidget(label)
        
        searchinput = QLineEdit()
        searchinput.setPlaceholderText(f"Enter {filtertype}...")
        searchinput.setStyleSheet("""QLineEdit {border: 2px solid #ccc; border-radius: 3px; padding: 5px; font-size: 9px;} QLineEdit:focus {border-color: #4CAF50;}""")
        layout.addWidget(searchinput)
        
        btnlayout = QHBoxLayout()
        btnlayout.setSpacing(2)
        
        searchbtn = QPushButton("✓")
        searchbtn.setMaximumHeight(22)
        searchbtn.setMaximumWidth(25)
        searchbtn.clicked.connect(self.applyfilters)
        searchbtn.setStyleSheet("""QPushButton {background-color: #4CAF50; color: white; border: none; padding: 2px; border-radius: 3px;} QPushButton:hover {background-color: #45a049;}""")
        btnlayout.addWidget(searchbtn)
        
        clearsearchbtn = QPushButton("✖")
        clearsearchbtn.setMaximumHeight(22)
        clearsearchbtn.setMaximumWidth(25)
        clearsearchbtn.clicked.connect(lambda: self.clearsearchfilter(widget))
        clearsearchbtn.setStyleSheet("""QPushButton {background-color: #f44336; color: white; border: none; padding: 2px; border-radius: 3px;} QPushButton:hover {background-color: #da190b;}""")
        btnlayout.addWidget(clearsearchbtn)
        
        btnlayout.addStretch()
        
        layout.addLayout(btnlayout)
        
        statuslabel = QLabel("")
        statuslabel.setStyleSheet("color: #666; font-size: 9px;")
        statuslabel.setWordWrap(True)
        layout.addWidget(statuslabel)
        
        layout.addStretch()
        
        searchinput.returnPressed.connect(self.applyfilters)
        
        widget.searchinput = searchinput
        widget.statuslabel = statuslabel
        widget.filtertype = filtertype
        widget.columnname = columnname
        
        widget.setLayout(layout)
        return widget
    
    def clearsearchfilter(self, searchwidget):
        searchwidget.searchinput.clear()
        searchwidget.statuslabel.setText("")
        self.applyfilters()
    
    def populatefilters(self, coveragedf):
        if 'SCC Name' in coveragedf.columns:
            uniquescc = coveragedf['SCC Name'].dropna().unique()
            self.sccfilter.additems(uniquescc)
        
        if 'Region' in coveragedf.columns:
            uniqueregions = coveragedf['Region'].dropna().unique()
            self.regionfilter.additems(uniqueregions)
            
        if 'Day Alert' in coveragedf.columns:
            uniqueday = coveragedf['Day Alert'].dropna().unique()
            validdays = []
            for d in uniqueday:
                try:
                    dayint = int(d)
                    validdays.append(dayint)
                except (ValueError, TypeError):
                    continue
                
            sorteddays = sorted(validdays)
            dayoptions = []
            
            for dayint in sorteddays:
                if dayint == 0:
                    dayoptions.append("Day 0")
                elif dayint >= 999:
                    dayoptions.append("Covered")
                else:
                    dayoptions.append(f"Day {dayint}")
                    
        self.dayalertfilter.additems(dayoptions, presorted=True)
            
    def applyfilters(self):
        if not hasattr(self, 'originalcoveragedf'):
            return
        
        self.coveragetable.setEnabled(False)
        
        try:
            filtereddf = self.originalcoveragedf.copy()
            
            selectedscc = self.sccfilter.getselecteditems()
            if selectedscc and 'SCC Name' in filtereddf.columns:
                filtereddf = filtereddf[filtereddf['SCC Name'].isin(selectedscc)]
            
            selectedregion = self.regionfilter.getselecteditems()
            if selectedregion and 'Region' in filtereddf.columns:
                filtereddf = filtereddf[filtereddf['Region'].isin(selectedregion)]
                
            selecteddays = self.dayalertfilter.getselecteditems()
            if selecteddays and 'Day Alert' in filtereddf.columns:
                dayvalues = []
                for daystr in selecteddays:
                    if "Day 0" in daystr:
                        dayvalues.append(0)
                    elif "Covered" in daystr:
                        dayvalues.append(999)
                    elif "Day " in daystr:
                        try:
                            daynum = int(daystr.split()[1])
                            dayvalues.append(daynum)
                        except (ValueError, IndexError):
                            continue
                if dayvalues:
                    filtereddf = filtereddf[filtereddf['Day Alert'].isin(dayvalues)]
            
            searchstatus = []
            searchcolumnmapping = {
                'MFG': 'MFG Code',
                'Part': 'Part Number',
                'SHP': 'SHP Code',
                'Country': 'SHP Country'}
            
            for searchfilter in self.searchfilters:
                searchtext = searchfilter.searchinput.text().strip().upper()
                filtertype = searchfilter.filtertype
                columnname = searchcolumnmapping.get(filtertype, searchfilter.columnname)
            
                if searchtext and columnname in filtereddf.columns:
                    if ',' in searchtext:
                        codes = [code.strip() for code in searchtext.split(',') if code.strip()]
                        filtereddf = filtereddf[filtereddf[columnname].astype(str).str.upper().isin(codes)]
                        searchfilter.statuslabel.setText(f"Searching: {', '.join(codes)}")
                        searchstatus.append(f"{filtertype}: {len(codes)} values")
                    else:
                        filtereddf = filtereddf[filtereddf[columnname].astype(str).str.upper().str.contains(searchtext, na=False)]
                        searchfilter.statuslabel.setText(f"Searching: {searchtext}")
                        searchstatus.append(f"{filtertype}: {searchtext}")
                else:
                    searchfilter.statuslabel.setText("")
            
            original_count = len(filtereddf)
            max_display_rows = 2000
            
            if len(filtereddf) > max_display_rows:
                filtereddf = filtereddf.head(max_display_rows)
                print(f"Performance: Showing first {max_display_rows} of {original_count} rows")
            else:
                print(f"Filtered to {len(filtereddf)} rows")
            
            self.displaycoveragetable(filtereddf)
            
        finally:
            self.coveragetable.setEnabled(True)
            
    def clearfilters(self):
        if hasattr(self, 'sccfilter'):
            self.sccfilter.selectallitems()
        if hasattr(self, 'regionfilter'):
            self.regionfilter.selectallitems()
        if hasattr(self, 'dayalertfilter'):
            self.dayalertfilter.selectallitems()
        if hasattr(self, 'searchfilters'):
            for searchfilter in self.searchfilters:
                searchfilter.searchinput.clear()
                searchfilter.statuslabel.setText("")
        
        if hasattr(self, 'originalcoveragedf'):
            self.displaycoveragetable(self.originalcoveragedf)
        
    def createcoveragedashboard(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        title = QLabel("Coverage Dashboard")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        buttonlayout = QHBoxLayout()
        
        loadbtn = QPushButton("Generate Coverage Analysis")
        loadbtn.clicked.connect(self.generatecoverageanalysis)
        loadbtn.setStyleSheet("""QPushButton {background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
        buttonlayout.addWidget(loadbtn)
        
        refreshbtn = QPushButton("Refresh Data")
        refreshbtn.clicked.connect(self.refreshcoveragedata)
        buttonlayout.addWidget(refreshbtn)
        
        exportbtn = QPushButton("Export to CSV")
        exportbtn.clicked.connect(self.exportcoveragetable)
        buttonlayout.addWidget(exportbtn)
        
        buttonlayout.addStretch()
        layout.addLayout(buttonlayout)
        
        self.filtersection = self.createfiltersection()
        layout.addWidget(self.filtersection)
        
        scrollarea = QScrollArea()
        scrollarea.setWidgetResizable(True)
        self.coveragetable = QTableWidget()
        self.coveragetable.setSortingEnabled(True)
        scrollarea.setWidget(self.coveragetable)
        layout.addWidget(scrollarea)
        
        self.coverageengine = CoverageAnalysisEngine(self.import_manager)
        
        widget.setLayout(layout)
        return widget
    
    def createcalloffforecasttab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        title = QLabel("Call-of Forecast and Waterfall")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        searchsection = self.createcalloffsearchsection()
        layout.addWidget(searchsection)
        
        self.dailyviewsection = self.createdailyviewsection()
        self.dailyviewsection.setVisible(False)
        layout.addWidget(self.dailyviewsection)
        
        widget.setLayout(layout)
        return widget
    
    def createindividualpartcoverage(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        title = QLabel("Individual Part Coverage")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        searchsection = self.createpartsearchsection()
        layout.addWidget(searchsection)
        
        self.partinfosection = self.createpartinfosection()
        self.partinfosection.setVisible(False)
        layout.addWidget(self.partinfosection)
        
        self.transactionsection = self.createtransactionsection()
        self.transactionsection.setVisible(False)
        layout.addWidget(self.transactionsection)
        
        widget.setLayout(layout)
        return widget
    
    def createpartsearchsection(self):
        widget = QWidget()
        widget.setMaximumHeight(150)
        layout = QHBoxLayout()
        
        searchlabel = QLabel("Part Number:")
        searchlabel.setFont(QFont("Arial", 12, QFont.Bold))
        layout.addWidget(searchlabel)
        
        self.partnumbersearch = QLineEdit()
        self.partnumbersearch.setPlaceholderText("Enter part number...")
        self.partnumbersearch.setMaximumWidth(200)
        self.partnumbersearch.setStyleSheet("""QLineEdit {border: 2px solid #ccc; border-radius: 5px; padding: 8px; font-size: 12px;} QLineEdit:focus {border-color: #4CAF50;}""")
        self.partnumbersearch.returnPressed.connect(self.searchpartcoverage)
        layout.addWidget(self.partnumbersearch)
        
        searchbtn = QPushButton("Search")
        searchbtn.clicked.connect(self.searchpartcoverage)
        searchbtn.setStyleSheet("""QPushButton {background-color: #4CAF50; color: white; padding: 8px 16px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
        layout.addWidget(searchbtn)
        
        clearbtn = QPushButton("Clear")
        clearbtn.clicked.connect(self.clearpartsearch)
        clearbtn.setStyleSheet("""QPushButton {background-color: #f44336; color: white; padding: 8px 16px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #da190b;}""")
        layout.addWidget(clearbtn)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget
    
    def createpartinfosection(self):
        widget = QWidget()
        widget.setMaximumHeight(200)
        
        layout = QVBoxLayout()
        infotitle = QLabel("Part Information")
        infotitle.setFont(QFont("Arial", 14, QFont.Bold))
        layout.addWidget(infotitle)
        
        gridwidget = QWidget()
        gridlayout = QGridLayout(gridwidget)
        gridlayout.setSpacing(10)
        
        self.partinfolabels = {}
        infofields = [
            ("Part Description", 0, 0), ("Supplier Name", 0, 1), ("MFG Code", 0, 2), ("SHP Code", 0, 3),
            ("Unit Load Qty", 1, 0), ("Multi Unit Load", 1, 1), ("Piece Price", 1, 2),
            ("Safety Stock", 2, 0), ("Safety Days", 2, 1), ("Initial Stock", 2, 2)]
        
        for fieldname, row, col in infofields:
            label = QLabel(f"{fieldname}:")
            label.setFont(QFont("Arial", 10, QFont.Bold))
            gridlayout.addWidget(label, row * 2, col)
            
            valuelabel = QLabel("--")
            valuelabel.setStyleSheet("color: #333; font-size: 11px; padding: 2px;")
            gridlayout.addWidget(valuelabel, row * 2 + 1, col)
            
            self.partinfolabels[fieldname] = valuelabel
        
        layout.addWidget(gridwidget)
        widget.setLayout(layout)
        return widget
    
    def createtransactionsection(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        transtitle = QLabel("Transaction Projections")
        transtitle.setFont(QFont("Arial", 14, QFont.Bold))
        layout.addWidget(transtitle)
        
        scrollarea = QScrollArea()
        scrollarea.setWidgetResizable(True)
        
        self.transactiontable = QTableWidget()
        self.transactiontable.setSortingEnabled(False)
        
        columnheaders = ["Date", "Transaction Type", "Receipt/Reqmt", "Available QTY", "ASN"]
        self.transactiontable.setColumnCount(len(columnheaders))
        self.transactiontable.setHorizontalHeaderLabels(columnheaders)
        
        self.transactiontable.setColumnWidth(0, 100)
        self.transactiontable.setColumnWidth(1, 120)
        self.transactiontable.setColumnWidth(2, 110)
        self.transactiontable.setColumnWidth(3, 110)
        self.transactiontable.setColumnWidth(4, 100)
        
        scrollarea.setWidget(self.transactiontable)
        layout.addWidget(scrollarea)
        
        widget.setLayout(layout)
        return widget
    
    def searchpartcoverage(self):
        partnumber = self.partnumbersearch.text().strip().upper()
        
        if not partnumber:
            QMessageBox.warning(self, "Missing Input", "Please enter a part number.")
            return
        
        try:
            success, message, datadict = self.coverageengine.loadrequireddata()
            if not success:
                QMessageBox.warning(self, "Missing Data", message)
                return
            
            partinfo, transactions = self.coverageengine.analyzeindivpart(partnumber, datadict)
            if not partinfo:
                QMessageBox.information(self, "Part Not Found", f"Part number {partnumber} was not found in the system.")
                return
            
            self.displaypartinfo(partinfo)
            self.displaytransactiontable(transactions)
            self.partinfosection.setVisible(True)
            self.transactionsection.setVisible(True)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Search failed: {str(e)}")
            
    def searchcalloffdata(self):
        partnumber = self.calloffpartsearch.text().strip().upper()
        
        if not partnumber:
            QMessageBox.warning(self, "Missing Input", "Please enter a part number.")
            return
        
        try:
            waterfalldata = self.waterfallengine.generatecalloffwaterfall(partnumber)
            
            if not waterfalldata:
                QMessageBox.information(self, "No Data", f"No call-off data found for part {partnumber}.")
                return
            
            self.displaydailyviewtable(waterfalldata)
            self.dailyviewsection.setVisible(True)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Call-off analysis failed: {str(e)}")
            
    def displaydailyviewtable(self, waterfalldata):
        if not waterfalldata:
            return
        
        today = datetime.now().date()
        shippingdates = self.waterfallengine.generateshippingdaterange(today, 90)
        headers = ['Call off Date'] + [date.strftime('%m/%d') for date in shippingdates]
        
        self.dailyviewtable.setRowCount(len(waterfalldata))
        self.dailyviewtable.setColumnCount(len(headers))
        self.dailyviewtable.setHorizontalHeaderLabels(headers)
        self.dailyviewtable.setColumnWidth(0, 120)
        
        for col in range(1, len(headers)):
            self.dailyviewtable.setColumnWidth(col, 80)
            
        for rowindex, rowdata in enumerate(waterfalldata):
            archivedate = rowdata['archive_date']
            calloffs = rowdata['call_offs']
            filefound = rowdata.get('file_found', True)
            
            dateitem = QTableWidgetItem(archivedate.strftime('%m/%d'))
            dateitem.setFlags(dateitem.flags() & ~Qt.ItemIsEditable)
            
            if filefound:
                dateitem.setBackground(Qt.lightGray)
            else:
                dateitem.setBackground(Qt.red)
                dateitem.setForeground(Qt.white)
                
            self.dailyviewtable.setItem(rowindex, 0, dateitem)
            
            for colindex, shippingdate in enumerate(shippingdates, 1):
                quantity = calloffs.get(shippingdate, 0)
                
                if quantity > 0:
                    qtyitem = QTableWidgetItem(f"{int(quantity):,}")
                    qtyitem.setBackground(Qt.cyan)
                else:
                    qtyitem = QTableWidgetItem("")
                
                qtyitem.setFlags(qtyitem.flags() & ~Qt.ItemIsEditable)
                self.dailyviewtable.setItem(rowindex, colindex, qtyitem)
  
    def displaypartinfo(self, partinfo):
        for fieldname, value in partinfo.items():
            if fieldname in self.partinfolabels:
                if fieldname in ['Initial Stock', 'Unit Load Qty', 'Multi Unit Load', 'Safety Stock']:
                    display_value = f"{int(value):,}" if value else "0"
                elif fieldname == 'Piece Price':
                    display_value = f"${float(value):.2f}" if value else "$0.00"
                elif fieldname == 'Safety Days':
                    display_value = f"{float(value):.2f}" if value else "0.00"
                else:
                    display_value = str(value) if value else "--"
            
                self.partinfolabels[fieldname].setText(display_value)
                
    def clearpartsearch(self):
        self.partnumbersearch.clear()
        self.partinfosection.setVisible(False)
        self.transactionsection.setVisible(False)
    
    def generatecoverageanalysis(self):
        try:
            success, message, datadict = self.coverageengine.loadrequireddata()
            if not success:
                QMessageBox.warning(self, "Missing Data", message)
                return
            
            coveragedf = self.coverageengine.buildcoverageanalysis(datadict, daysforward=90)
            
            if coveragedf.empty:
                QMessageBox.information(self, "No Data", "No parts with consumption found.")
                return
            
            self.originalcoveragedf = coveragedf.copy()
            self.currentcoveragedf = coveragedf
            self.populatefilters(coveragedf)
            self.filtersection.setVisible(True)
            self.displaycoveragetable(coveragedf)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Analysis failed: {str(e)}")
            
    def displaycoveragetable(self, coveragedf):
        if coveragedf.empty:
            return
            
        try:
            self.coveragetable.itemChanged.disconnect()
        except:
            pass
        
        if (not hasattr(self, '_last_column_count') or 
            self._last_column_count != len(coveragedf.columns)):
            
            self.coveragetable.setColumnCount(len(coveragedf.columns))
            self.coveragetable.setHorizontalHeaderLabels(coveragedf.columns.tolist())
            self._last_column_count = len(coveragedf.columns)
            
            self.comments_col = None
            if 'Comments' in coveragedf.columns:
                self.comments_col = list(coveragedf.columns).index('Comments')
        
        self.coveragetable.setRowCount(len(coveragedf))
        
        for row in range(len(coveragedf)):
            for col in range(len(coveragedf.columns)):
                value = coveragedf.iloc[row, col]
                
                if isinstance(value, (int, float)) and pd.notna(value):
                    if abs(value) >= 1:
                        display_value = f"{value:,.0f}"
                    else:
                        display_value = f"{value:.2f}"
                else:
                    display_value = str(value) if pd.notna(value) else ""
                
                item = QTableWidgetItem(display_value)
                
                if col == self.comments_col:
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                else:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                
                self.coveragetable.setItem(row, col, item)
        
        self.apply_color_coding(coveragedf)
        
        if self.comments_col is not None:
            self.coveragetable.setColumnWidth(self.comments_col, 200)
        
        self.coveragetable.itemChanged.connect(self.oncommentchanged)
        
        print(f"Table updated: {len(coveragedf)} rows, {len(coveragedf.columns)} columns")
    
    def apply_color_coding(self, coveragedf):
        if 'Unit Load Qty' not in coveragedf.columns:
            return
        
        dailystartcol = None
        for i, col in enumerate(coveragedf.columns):
            try:
                datetime.strptime(col, '%m/%d')
                dailystartcol = i
                break
            except:
                continue
            
        if dailystartcol is None:
            return
        
        for row in range(len(coveragedf)):
            unit_load_qty = coveragedf.iloc[row]['Unit Load Qty']
            if pd.isna(unit_load_qty) or unit_load_qty <= 0:
                unit_load_qty = 1
            
            for col in range(dailystartcol, len(coveragedf.columns)):
                value = coveragedf.iloc[row, col]
                
                if isinstance(value, (int, float)) and pd.notna(value):
                    item = self.coveragetable.item(row, col)
                    if item:
                        if value <= 0:
                            item.setBackground(Qt.red)
                            item.setForeground(Qt.white)
                        elif value < unit_load_qty:
                            item.setBackground(Qt.yellow)
                        else:
                            item.setBackground(Qt.white)
                            item.setForeground(Qt.black)
        
    def oncommentchanged(self, item):
        try:
            row = item.row()
            col = item.column()
            
            if 'Comments' in self.currentcoveragedf.columns:
                commentscol = list(self.currentcoveragedf.columns).index('Comments')
                
                if col == commentscol:
                    partno = str(self.currentcoveragedf.iloc[row]['Part Number'])
                    commenttext = item.text()
                    comments = self.coverageengine.loadcoveragecomments()
                    
                    if commenttext.strip():
                        comments[partno] = commenttext.strip()
                    else:
                        comments.pop(partno, None)
                    
                    self.coverageengine.savecoveragecomments(comments)
        
        except Exception as e:
            print(f"Error saving comment: {e}")
            
    def displaytransactiontable(self, transactiondata):
        if not transactiondata:
            return
        
        self.transactiontable.setRowCount(len(transactiondata))
        
        for row, transaction in enumerate(transactiondata):
            dateitem = QTableWidgetItem(transaction['Date'])
            dateitem.setFlags(dateitem.flags() & ~Qt.ItemIsEditable)
            self.transactiontable.setItem(row, 0, dateitem)
            typeitem = QTableWidgetItem(transaction['Transaction Type'])
            typeitem.setFlags(typeitem.flags() & ~Qt.ItemIsEditable)
            
            if transaction['Transaction Type'] == 'Stock':
                typeitem.setBackground(Qt.lightGray)
            elif transaction['Transaction Type'] == 'Req':
                typeitem.setBackground(Qt.red)
                typeitem.setForeground(Qt.white)
            elif transaction['Transaction Type'] == 'GR':
                typeitem.setBackground(Qt.green)
                typeitem.setForeground(Qt.white)
            
            self.transactiontable.setItem(row, 1, typeitem)
            receiptitem = QTableWidgetItem(transaction['Receipt/Reqmt'])
            receiptitem.setFlags(receiptitem.flags() & ~Qt.ItemIsEditable)
            self.transactiontable.setItem(row, 2, receiptitem)
            qtyitem = QTableWidgetItem(f"{transaction['Available QTY']:,}")
            qtyitem.setFlags(qtyitem.flags() & ~Qt.ItemIsEditable)
            
            if transaction['Available QTY'] <= 0:
                qtyitem.setBackground(Qt.red)
                qtyitem.setForeground(Qt.white)
            elif transaction['Available QTY'] < 100:
                qtyitem.setBackground(Qt.yellow)
                
            self.transactiontable.setItem(row, 3, qtyitem)
            asnitem = QTableWidgetItem(transaction['ASN'])
            asnitem.setFlags(asnitem.flags() & ~Qt.ItemIsEditable)
            self.transactiontable.setItem(row, 4, asnitem)
        
    def refreshcoveragedata(self):
        self.generatecoverageanalysis()
        
    def exportcoveragetable(self):
        if not hasattr(self, 'currentcoveragedf') or self.currentcoveragedf.empty:
            QMessageBox.warning(self, "No Data", "Generate analysis first.")
            return
        
        filename, _ = QFileDialog.getSaveFileName(self, "Export Coverage Analysis", f"Coverage_Analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", "CSV files (*.csv)")
        
        if filename:
            success = self.coverageengine.exporttocsv(self.currentcoveragedf, filename)
            if success:
                QMessageBox.information(self, "Export Complete", f"Exported to:\n{filename}")
            else:
                QMessageBox.critical(self, "Export Failed", "Could not export file.")
        
    def closeEvent(self, event):
        event.accept()