import numpy as np
import pandas as pd
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTabWidget, QTextEdit, QTableWidget, QTableWidgetItem, QMessageBox, QScrollArea, QFileDialog, QComboBox, QListWidget, QListWidgetItem, QCheckBox, QFrame, QApplication, QLineEdit, QGridLayout, QProgressDialog)
from PyQt5.QtCore import Qt, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QColor
from datetime import datetime, timedelta
from app.utils.config import APPWINDOWSIZE
from app.data.import_manager import DataImportManager
from app.supply_chain_coordination.coverage_analysis import CoverageAnalysisEngine
from app.supply_chain_coordination.waterfall_analysis import WaterfallAnalysisEngine
from app.supply_chain_coordination.ldjis_coverage import LDJISCoverageEngine
 
 
class SupplyChainCoordinationWindow(QMainWindow):
    def __init__(self, userdata, parent=None):
        super().__init__(parent)
        self.userdata = userdata
        self.import_manager = DataImportManager()
        self.coverageengine = CoverageAnalysisEngine(self.import_manager)
        self.waterfallengine = WaterfallAnalysisEngine(self.import_manager)
        self.ldjiscoverageengine = LDJISCoverageEngine(self.import_manager)
        self.ldjismodelnames = ['EX90', 'PS3']
        self._comments_cache = None
        self.setWindowTitle("Supply Chain Coordination Application")
        self.resize(*APPWINDOWSIZE)
        self.setupui()
 
    def setupui(self):
        centralwidget = QWidget()
        self.setCentralWidget(centralwidget)
 
        centralwidget.setStyleSheet("QWidget#centralwidget { background-color: #156082; }")
        centralwidget.setObjectName("centralwidget")
 
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
 
        headerwidget = QWidget()
        headerwidget.setStyleSheet("background-color: #156082;")
        headerlayout = QVBoxLayout(headerwidget)
        headerlayout.setContentsMargins(10, 8, 10, 8)
 
        title = QLabel("Supply Chain Coordination")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: white; background-color: transparent;")
        headerlayout.addWidget(title)
 
        layout.addWidget(headerwidget)
 
        self.setMinimumWidth(1000)
 
        tabs = QTabWidget()
        tabs.tabBar().setElideMode(Qt.ElideNone)
        tabs.tabBar().setExpanding(False)
        tabs.setStyleSheet("""
            QTabWidget {
                background-color: #CCECFF;
            }
            QTabWidget::pane {
                background-color: white;
                border: 1px solid #99CCEE;
                border-top: 0px;
            }
            QTabBar {
                background-color: #156082;
            }
            QTabBar QToolButton {
                background-color: #156082;
                border: 1px solid #99CCEE;
                color: white;
            }
            QTabBar::tab {
                background-color: white;
                color: #1A3A6B;
                font-weight: bold;
                padding: 6px 14px;
                min-width: 172px;
                border: 1px solid #99CCEE;
                border-bottom: none;
                border-radius: 4px 4px 0px 0px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: #CCECFF;
                color: black;
                border: 1px solid #99CCEE;
                border-bottom: 1px solid #CCECFF;
            }
            QTabBar::tab:hover:!selected {
                background-color: #E8F6FF;
            }
        """)
 
        coveragedashtab = self.createcoveragedashboard()
        coverageindivtab = self.createindividualpartcoverage()
        calloffforecasttab = self.createcalloffforecasttab()
        ldjiscoveragetab = self.createldjiscoveragetab()
        alertstab = self.createalertstab()
        piwdtab = self.createpiwdtab()
 
        tabs.addTab(coveragedashtab, "Coverage Dashboard")
        tabs.addTab(coverageindivtab, "Individual Part Coverage")
        tabs.addTab(calloffforecasttab, "Call-off Forecast and Waterfall")
        tabs.addTab(ldjiscoveragetab, "LDJIS Coverage")
        tabs.addTab(alertstab, "Alerts Breakdown")
        tabs.addTab(piwdtab, "PIWD Report")

        layout.addWidget(tabs)
 
        self.statusBar().showMessage(f"Logged in as: {self.userdata['username']}")
        centralwidget.setLayout(layout)
 
    def createfiltersection(self):
        widget = QWidget()
        widget.setMaximumHeight(190)
 
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
 
    def createalertfiltersection(self):
        widget = QWidget()
        widget.setMaximumHeight(190)
 
        layout = QHBoxLayout()
 
        filterlabel = QLabel("Filters:")
        filterlabel.setFont(QFont("Arial", 12, QFont.Bold))
        filterlabel.setAlignment(Qt.AlignTop)
        layout.addWidget(filterlabel)
 
        self.alerts_scc_filter = self.createmultiselectdropdown("Select SCC Names...", "SCC")
        self.alerts_scc_filter.selectionChanged.connect(self.applyalertfilters)
        layout.addWidget(self.alerts_scc_filter)
 
        self.alerts_type_filter = self.createmultiselectdropdown("Select Alert...", "Alert")
        self.alerts_type_filter.selectionChanged.connect(self.applyalertfilters)
        layout.addWidget(self.alerts_type_filter)
 
        self.alerts_part_filter = self.createmultiselectdropdown("Select Part-Alert...", "Part-Alert")
        self.alerts_part_filter.selectionChanged.connect(self.applyalertfilters)
        layout.addWidget(self.alerts_part_filter)
 
        clearfiltersbtn = QPushButton("Clear All Filters")
        clearfiltersbtn.clicked.connect(self.clearalertfilters)
        clearfiltersbtn.setMaximumWidth(120)
        clearfiltersbtn.setMaximumHeight(30)
        layout.addWidget(clearfiltersbtn)
 
        layout.addStretch()
        widget.setLayout(layout)
        return widget
 
    def createmultiselectdropdown(self, placeholdertext, filtertype="SCC"):
        class SimpleMultiSelectFilter(QWidget):
            selectionChanged = pyqtSignal()
 
            def __init__(self, placeholder="Select items...", filtertype="SCC"):
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
 
            def additems(self, items, presorted=False):
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
                        self.all_items.append(str(item))
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
                self.label.setText(f"{self.filtertype} Filter:")
 
            def getselecteditems(self):
                return list(self.selected_items)
 
            def selectallitems(self):
                if len(self.all_items) > 50:
                    progress = self.show_progress_for_large_operation()
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
        searchbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; border: none; padding: 2px; border-radius: 3px;} QPushButton:hover {background-color: #45a049;}""")
        btnlayout.addWidget(searchbtn)
 
        clearsearchbtn = QPushButton("✖")
        clearsearchbtn.setMaximumHeight(22)
        clearsearchbtn.setMaximumWidth(25)
        clearsearchbtn.clicked.connect(lambda: self.clearsearchfilter(widget))
        clearsearchbtn.setStyleSheet("""QPushButton {background-color: #E97132; color: white; border: none; padding: 2px; border-radius: 3px;} QPushButton:hover {background-color: #da190b;}""")
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
                    validdays.append(int(d))
                except (ValueError, TypeError):
                    continue
 
            dayoptions = []
            for dayint in sorted(validdays):
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
                            dayvalues.append(int(daystr.split()[1]))
                        except (ValueError, IndexError):
                            continue
                if dayvalues:
                    filtereddf = filtereddf[filtereddf['Day Alert'].isin(dayvalues)]
 
            searchcolumnmapping = {
                'MFG': 'MFG Code',
                'Part': 'Part Number',
                'SHP': 'SHP Code',
                'Country': 'SHP Country',
            }
 
            for searchfilter in self.searchfilters:
                searchtext = searchfilter.searchinput.text().strip().upper()
                filtertype = searchfilter.filtertype
                columnname = searchcolumnmapping.get(filtertype, searchfilter.columnname)
 
                if searchtext and columnname in filtereddf.columns:
                    if ',' in searchtext:
                        codes = [code.strip() for code in searchtext.split(',') if code.strip()]
                        filtereddf = filtereddf[
                            filtereddf[columnname].astype(str).str.upper().isin(codes)
                        ]
                        searchfilter.statuslabel.setText(f"Searching: {', '.join(codes)}")
                    else:
                        filtereddf = filtereddf[
                            filtereddf[columnname].astype(str).str.upper().str.contains(
                                searchtext, na=False
                            )
                        ]
                        searchfilter.statuslabel.setText(f"Searching: {searchtext}")
                else:
                    searchfilter.statuslabel.setText("")
 
            max_display_rows = 2000
            if len(filtereddf) > max_display_rows:
                filtereddf = filtereddf.head(max_display_rows)
 
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
        loadbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 10px 20px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
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
 
        widget.setLayout(layout)
        return widget
 
    def createcalloffforecasttab(self):
        widget = QWidget()
        layout = QVBoxLayout()
 
        title = QLabel("Call-off Forecast and Waterfall")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
 
        searchsection = self.createcalloffsearchsection()
        layout.addWidget(searchsection)
 
        self.dailyviewsection = self.createdailyviewsection()
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
        layout.addWidget(self.partinfosection)
 
        self.transactionsection = self.createtransactionsection()
        layout.addWidget(self.transactionsection)
 
        widget.setLayout(layout)
        return widget
 
    def createldjiscoveragetab(self):
        widget = QWidget()
        layout = QVBoxLayout()
 
        title = QLabel("LDJIS Coverage")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
 
        btnlayout = QHBoxLayout()
        generatebtn = QPushButton("Generate LDJIS Coverage")
        generatebtn.clicked.connect(self.generateldjiscoverage)
        generatebtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 10px 20px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
        btnlayout.addWidget(generatebtn)
        btnlayout.addStretch()
        layout.addLayout(btnlayout)
 
        tablestyle = """QTableWidget {gridline-color: #d0d0d0; background-color: white;} QTableWidget::item {padding: 6px; border: 1px solid #d0d0d0;} QHeaderView::section {background-color: #f0f0f0; padding: 8px; border: 1px solid #d0d0d0; font-weight: bold;}"""
 
        scrollarea = QScrollArea()
        scrollarea.setWidgetResizable(True)
        self.ldjistable = QTableWidget()
        self.ldjistable.setSortingEnabled(False)
        self.ldjistable.setStyleSheet(tablestyle)
        scrollarea.setWidget(self.ldjistable)
        layout.addWidget(scrollarea)
 
        widget.setLayout(layout)
        return widget
 
    def createalertstab(self):
        widget = QWidget()
        layout = QVBoxLayout()
 
        title = QLabel("Alerts Breakdown")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
 
        btnlayout = QHBoxLayout()
        generatebtn = QPushButton("Generate Alerts Breakdown")
        generatebtn.clicked.connect(self.generatealertsbreakdown)
        generatebtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 10px 20px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
        btnlayout.addWidget(generatebtn)
 
        exportbtn = QPushButton("Export to CSV")
        exportbtn.clicked.connect(self.exportalertstable)
        btnlayout.addWidget(exportbtn)
 
        btnlayout.addStretch()
        layout.addLayout(btnlayout)
 
        self.alertfiltersection = self.createalertfiltersection()
        layout.addWidget(self.alertfiltersection)
 
        tablestyle = """QTableWidget {gridline-color: #d0d0d0; background-color: white;} QTableWidget::item {padding: 6px; border: 1px solid #d0d0d0;} QHeaderView::section {background-color: #f0f0f0; padding: 8px; border: 1px solid #d0d0d0; font-weight: bold;}"""
 
        scrollarea = QScrollArea()
        scrollarea.setWidgetResizable(True)
        self.alertstable = QTableWidget()
        self.alertstable.setSortingEnabled(True)
        self.alertstable.setStyleSheet(tablestyle)
        scrollarea.setWidget(self.alertstable)
        layout.addWidget(scrollarea)
 
        widget.setLayout(layout)
        return widget
    
    def createpiwdtab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        title = QLabel("PIWD Report")
        title.setFont(QFont("Arial", 15, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        btnlayout = QHBoxLayout()
        generatebtn = QPushButton("Generate PIWD Report")
        generatebtn.clicked.connect(self.generatepiwdreport)
        generatebtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 10px 20px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
        btnlayout.addWidget(generatebtn)
        
        exportbtn = QPushButton("Export to CSV")
        exportbtn.clicked.connect(self.exportpiwdreport)
        btnlayout.addWidget(exportbtn)
        
        btnlayout.addStretch()
        layout.addLayout(btnlayout)

        self.piwdfiltersection = self.createpiwdfiltersection()
        layout.addWidget(self.piwdfiltersection)
        
        tablestyle = """QTableWidget {gridline-color: #d0d0d0; background-color: white;} QTableWidget::item {padding: 6px; border: 1px solid #d0d0d0;} QHeaderView::section {background-color: #f0f0f0; padding: 8px; border: 1px solid #d0d0d0; font-weight: bold;}"""
        
        scrollarea = QScrollArea()
        scrollarea.setWidgetResizable(True)
        self.piwdtable = QTableWidget()
        self.piwdtable.setSortingEnabled(True)
        self.piwdtable.setStyleSheet(tablestyle)
        scrollarea.setWidget(self.piwdtable)
        layout.addWidget(scrollarea)
        
        widget.setLayout(layout)
        return widget
 
    def createpiwdfiltersection(self):
        widget = QWidget()
        widget.setMaximumHeight(190)

        layout = QHBoxLayout()

        filterlabel = QLabel("Filters:")
        filterlabel.setFont(QFont("Arial", 12, QFont.Bold))
        filterlabel.setAlignment(Qt.AlignTop)
        layout.addWidget(filterlabel)

        self.piwd_scc_filter = self.createmultiselectdropdown("Select SCC Names...", "SCC")
        self.piwd_scc_filter.selectionChanged.connect(self.applypiwdfilters)
        layout.addWidget(self.piwd_scc_filter)

        self.piwd_part_filter = self.createmultiselectdropdown("Select Parts...", "Part")
        self.piwd_part_filter.selectionChanged.connect(self.applypiwdfilters)
        layout.addWidget(self.piwd_part_filter)

        clearfiltersbtn = QPushButton("Clear All Filters")
        clearfiltersbtn.clicked.connect(self.clearpiwdfilters)
        clearfiltersbtn.setMaximumWidth(120)
        clearfiltersbtn.setMaximumHeight(30)
        layout.addWidget(clearfiltersbtn)

        layout.addStretch()
        widget.setLayout(layout)
        return widget

    def createcalloffsearchsection(self):
        widget = QWidget()
        widget.setMaximumHeight(100)
 
        layout = QHBoxLayout()
 
        searchlabel = QLabel("Part Number:")
        searchlabel.setFont(QFont("Arial", 12, QFont.Bold))
        layout.addWidget(searchlabel)
 
        self.calloffpartsearch = QLineEdit()
        self.calloffpartsearch.setPlaceholderText("Enter part number for call-off analysis...")
        self.calloffpartsearch.setMaximumWidth(250)
        self.calloffpartsearch.setStyleSheet("""QLineEdit {border: 2px solid #ccc; border-radius: 5px; padding: 8px; font-size: 12px;} QLineEdit:focus {border-color: #4CAF50;}""")
        self.calloffpartsearch.returnPressed.connect(self.searchcalloffdata)
        layout.addWidget(self.calloffpartsearch)
 
        searchbtn = QPushButton("Generate Waterfall")
        searchbtn.clicked.connect(self.searchcalloffdata)
        searchbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 8px 16px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #1976D2;}""")
        layout.addWidget(searchbtn)
 
        clearbtn = QPushButton("Clear")
        clearbtn.clicked.connect(self.clearcalloffanalysis)
        clearbtn.setStyleSheet("""QPushButton {background-color: #E97132; color: white; padding: 8px 16px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #da190b;}""")
        layout.addWidget(clearbtn)
 
        layout.addStretch()
        widget.setLayout(layout)
        return widget
 
    def createdailyviewsection(self):
        widget = QWidget()
        layout = QVBoxLayout()
        tablestyle = """QTableWidget {gridline-color: #d0d0d0; background-color: white;} QTableWidget::item {padding: 8px; border: 1px solid #d0d0d0;} QHeaderView::section {background-color: #f0f0f0; padding: 10px; border: 1px solid #d0d0d0; font-weight: bold;}"""
 
        dailytitlerow = QHBoxLayout()
        dailytitle = QLabel("Daily View - Call-off Forecast Waterfall")
        dailytitle.setFont(QFont("Arial", 14, QFont.Bold))
        dailytitlerow.addWidget(dailytitle)
        dailydeltatitle = QLabel("Daily Delta")
        dailydeltatitle.setFont(QFont("Arial", 14, QFont.Bold))
        dailytitlerow.addWidget(dailydeltatitle)
        layout.addLayout(dailytitlerow)
 
        dailytablesrow = QHBoxLayout()
 
        dailyscroll = QScrollArea()
        dailyscroll.setWidgetResizable(True)
        self.dailyviewtable = QTableWidget()
        self.dailyviewtable.setSortingEnabled(False)
        self.dailyviewtable.setStyleSheet(tablestyle)
        dailyscroll.setWidget(self.dailyviewtable)
        dailytablesrow.addWidget(dailyscroll)
 
        dailydeltascroll = QScrollArea()
        dailydeltascroll.setWidgetResizable(True)
        self.dailydeltatable = QTableWidget()
        self.dailydeltatable.setSortingEnabled(False)
        self.dailydeltatable.setStyleSheet(tablestyle)
        dailydeltascroll.setWidget(self.dailydeltatable)
        dailytablesrow.addWidget(dailydeltascroll)
 
        layout.addLayout(dailytablesrow)
 
        weeklytitlerow = QHBoxLayout()
        weeklytitle = QLabel("Weekly Summary - Call-off Forecast Waterfall")
        weeklytitle.setFont(QFont("Arial", 14, QFont.Bold))
        weeklytitlerow.addWidget(weeklytitle)
        weeklydeltatitle = QLabel("Weekly Delta")
        weeklydeltatitle.setFont(QFont("Arial", 14, QFont.Bold))
        weeklytitlerow.addWidget(weeklydeltatitle)
        layout.addLayout(weeklytitlerow)
 
        weeklytablesrow = QHBoxLayout()
 
        weeklyscroll = QScrollArea()
        weeklyscroll.setWidgetResizable(True)
        self.weeklyviewtable = QTableWidget()
        self.weeklyviewtable.setSortingEnabled(False)
        self.weeklyviewtable.setStyleSheet(tablestyle)
        weeklyscroll.setWidget(self.weeklyviewtable)
        weeklytablesrow.addWidget(weeklyscroll)
 
        weeklydeltascroll = QScrollArea()
        weeklydeltascroll.setWidgetResizable(True)
        self.weeklydeltatable = QTableWidget()
        self.weeklydeltatable.setSortingEnabled(False)
        self.weeklydeltatable.setStyleSheet(tablestyle)
        weeklydeltascroll.setWidget(self.weeklydeltatable)
        weeklytablesrow.addWidget(weeklydeltascroll)
 
        layout.addLayout(weeklytablesrow)
 
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
        searchbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 8px 16px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
        layout.addWidget(searchbtn)
 
        clearbtn = QPushButton("Clear")
        clearbtn.clicked.connect(self.clearpartsearch)
        clearbtn.setStyleSheet("""QPushButton {background-color: #E97132; color: white; padding: 8px 16px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #da190b;}""")
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
            ("Safety Stock", 2, 0), ("Safety Days", 2, 1), ("Initial Stock", 2, 2),
        ]
 
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
 
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Search failed: {str(e)}")
 
    def searchcalloffdata(self):
        partnumber = self.calloffpartsearch.text().strip().upper()
 
        if not partnumber:
            QMessageBox.warning(self, "Missing Input", "Please enter a part number.")
            return
 
        try:
            valid, message, dates = self.waterfallengine.validatearchiveavailability()
            if not valid:
                QMessageBox.warning(self, "Archive Not Found", f"{message}\n\nExpected directory:\n{self.waterfallengine.archive_dir}")
                return
 
            waterfalldata = self.waterfallengine.generatecalloffwaterfall(partnumber)
 
            if not waterfalldata:
                QMessageBox.information(self, "No Data", f"No call-off data found for part {partnumber}.")
                return
 
            shippingdates = self.waterfallengine.generateshippingdaterange(datetime.now().date(), 90)
            self.displaydailyviewtable(waterfalldata)
            self.displaydailydeltatable(waterfalldata, shippingdates)
            self.displayweeklyviewtable(waterfalldata, shippingdates)
            self.displayweeklydeltatable(waterfalldata, shippingdates)
 
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
                dateitem.setBackground(QColor(255, 204, 204))
                dateitem.setForeground(QColor(156, 0, 6))
 
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
 
    def displaydailydeltatable(self, waterfalldata, shippingdates):
        if not waterfalldata or not shippingdates:
            return
 
        headers = ['Call off Date'] + [d.strftime('%m/%d') for d in shippingdates]
 
        self.dailydeltatable.setRowCount(len(waterfalldata))
        self.dailydeltatable.setColumnCount(len(headers))
        self.dailydeltatable.setHorizontalHeaderLabels(headers)
        self.dailydeltatable.setColumnWidth(0, 120)
        for col in range(1, len(headers)):
            self.dailydeltatable.setColumnWidth(col, 80)
 
        for rowindex, rowdata in enumerate(waterfalldata):
            dateitem = QTableWidgetItem(rowdata['archive_date'].strftime('%m/%d'))
            dateitem.setFlags(dateitem.flags() & ~Qt.ItemIsEditable)
            dateitem.setBackground(Qt.lightGray)
            self.dailydeltatable.setItem(rowindex, 0, dateitem)
 
            if rowindex == 0:
                for colindex in range(1, len(headers)):
                    item = QTableWidgetItem("")
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.dailydeltatable.setItem(rowindex, colindex, item)
                continue
 
            current = rowdata['call_offs']
            previous = waterfalldata[rowindex - 1]['call_offs']
 
            for colindex, shippingdate in enumerate(shippingdates, 1):
                delta = current.get(shippingdate, 0) - previous.get(shippingdate, 0)
 
                if delta != 0:
                    qtyitem = QTableWidgetItem(f"{int(delta):+,}")
                    if delta > 0:
                        qtyitem.setBackground(QColor(204, 255, 204))
                        qtyitem.setForeground(QColor(0, 156, 6))
                    else:
                        qtyitem.setBackground(QColor(255, 204, 204))
                        qtyitem.setForeground(QColor(156, 0, 6))
                else:
                    qtyitem = QTableWidgetItem("")
 
                qtyitem.setFlags(qtyitem.flags() & ~Qt.ItemIsEditable)
                self.dailydeltatable.setItem(rowindex, colindex, qtyitem)
 
    def displayweeklyviewtable(self, waterfalldata, shippingdates):
        if not waterfalldata or not shippingdates:
            return
 
        weeks = {}
        for d in shippingdates:
            iso = d.isocalendar()
            key = (iso[0], iso[1])
            weeks.setdefault(key, []).append(d)
 
        sortedweeks = sorted(weeks.keys())
        headers = ['Call off Date']
 
        for (yr, wk) in sortedweeks:
            weekdates = weeks[(yr, wk)]
            weekstart = min(weekdates)
            headers.append(f"Wk {wk}\n{weekstart.strftime('%m/%d')}")
 
        self.weeklyviewtable.setRowCount(len(waterfalldata))
        self.weeklyviewtable.setColumnCount(len(headers))
        self.weeklyviewtable.setHorizontalHeaderLabels(headers)
        self.weeklyviewtable.setColumnWidth(0, 120)
 
        for col in range(1, len(headers)):
            self.weeklyviewtable.setColumnWidth(col, 90)
 
        for rowindex, rowdata in enumerate(waterfalldata):
            archivedate = rowdata['archive_date']
            calloffs = rowdata['call_offs']
 
            dateitem = QTableWidgetItem(archivedate.strftime('%m/%d'))
            dateitem.setFlags(dateitem.flags() & ~Qt.ItemIsEditable)
            dateitem.setBackground(Qt.lightGray)
            self.weeklyviewtable.setItem(rowindex, 0, dateitem)
 
            for colindex, (yr, wk) in enumerate(sortedweeks, 1):
                weektotal = sum(calloffs.get(d, 0) for d in weeks[(yr, wk)])
 
                if weektotal > 0:
                    qtyitem = QTableWidgetItem(f"{int(weektotal):,}")
                    qtyitem.setBackground(Qt.cyan)
                else:
                    qtyitem = QTableWidgetItem("")
 
                qtyitem.setFlags(qtyitem.flags() & ~Qt.ItemIsEditable)
                self.weeklyviewtable.setItem(rowindex, colindex, qtyitem)
 
    def displayweeklydeltatable(self, waterfalldata, shippingdates):
        if not waterfalldata or not shippingdates:
            return
 
        weeks = {}
        for d in shippingdates:
            iso = d.isocalendar()
            key = (iso[0], iso[1])
            weeks.setdefault(key, []).append(d)
        sorted_weeks = sorted(weeks.keys())
 
        headers = ['Call off Date']
        for (yr, wk) in sorted_weeks:
            week_start = min(weeks[(yr, wk)])
            headers.append(f"Wk {wk}\n{week_start.strftime('%m/%d')}")
 
        self.weeklydeltatable.setRowCount(len(waterfalldata))
        self.weeklydeltatable.setColumnCount(len(headers))
        self.weeklydeltatable.setHorizontalHeaderLabels(headers)
        self.weeklydeltatable.setColumnWidth(0, 120)
        for col in range(1, len(headers)):
            self.weeklydeltatable.setColumnWidth(col, 90)
 
        for rowindex, rowdata in enumerate(waterfalldata):
            dateitem = QTableWidgetItem(rowdata['archive_date'].strftime('%m/%d'))
            dateitem.setFlags(dateitem.flags() & ~Qt.ItemIsEditable)
            dateitem.setBackground(Qt.lightGray)
            self.weeklydeltatable.setItem(rowindex, 0, dateitem)
 
            if rowindex == 0:
                for colindex in range(1, len(headers)):
                    item = QTableWidgetItem("")
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.weeklydeltatable.setItem(rowindex, colindex, item)
                continue
 
            current = waterfalldata[rowindex]['call_offs']
            previous = waterfalldata[rowindex - 1]['call_offs']
 
            for colindex, (yr, wk) in enumerate(sorted_weeks, 1):
                current_total = sum(current.get(d, 0) for d in weeks[(yr, wk)])
                prev_total = sum(previous.get(d, 0) for d in weeks[(yr, wk)])
                delta = current_total - prev_total
 
                if delta != 0:
                    qtyitem = QTableWidgetItem(f"{int(delta):+,}")
                    if delta > 0:
                        qtyitem.setBackground(QColor(204, 255, 204))
                        qtyitem.setForeground(QColor(0, 156, 5))
                    else:
                        qtyitem.setBackground(QColor(255, 204, 204))
                        qtyitem.setForeground(QColor(156, 0, 6))
                else:
                    qtyitem = QTableWidgetItem("")
 
                qtyitem.setFlags(qtyitem.flags() & ~Qt.ItemIsEditable)
                self.weeklydeltatable.setItem(rowindex, colindex, qtyitem)
 
    def displaypartinfo(self, partinfo):
        for fieldname, value in partinfo.items():
            if fieldname in self.partinfolabels:
                if fieldname in ('Initial Stock', 'Unit Load Qty', 'Multi Unit Load', 'Safety Stock'):
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
        for label in self.partinfolabels.values():
            label.setText("--")
        self.transactiontable.setRowCount(0)
 
    def clearcalloffanalysis(self):
        self.calloffpartsearch.clear()
        self.dailyviewtable.setRowCount(0)
        self.dailydeltatable.setRowCount(0)
        self.weeklyviewtable.setRowCount(0)
        self.weeklydeltatable.setRowCount(0)
 
    def generatecoverageanalysis(self):
        self._comments_cache = None
        try:
            success, message, datadict = self.coverageengine.loadrequireddata()
            if not success:
                QMessageBox.warning(self, "Missing Data", message)
                return
 
            coveragedf = self.coverageengine.buildcoverageanalysis(datadict, daysforward=40)
 
            if coveragedf.empty:
                QMessageBox.information(self, "No Data", "No parts with consumption found.")
                return
 
            self.originalcoveragedf = coveragedf.copy()
            self.currentcoveragedf = coveragedf
            self.populatefilters(coveragedf)
            self.displaycoveragetable(coveragedf)
 
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Analysis failed: {str(e)}")
 
    def generateldjiscoverage(self):
        try:
            success, message, datadict = self.coverageengine.loadrequireddata()
            if not success:
                QMessageBox.warning(self, "Missing Data", message)
                return
 
            ok, msg, ldjisdf = self.ldjiscoverageengine.buildldjiscoveragedata(datadict)
            if not ok:
                QMessageBox.warning(self, "No LDJIS Data", msg)
                return
 
            self.ldjisdf = ldjisdf
            self.ldjisworkingdays = self.ldjiscoverageengine.generateworkingdays(
                datetime.now().date(), 20
            )
            self.displayldjistable()
 
        except Exception as e:
            QMessageBox.critical(self, "Error", f"LDJIS analysis failed: {str(e)}")
 
    def displayldjistable(self):
        if not hasattr(self, 'ldjisdf') or self.ldjisdf.empty:
            return
 
        try:
            self.ldjistable.itemChanged.disconnect()
        except Exception:
            pass
 
        nmodels = len(self.ldjismodelnames)
        nparts = len(self.ldjisdf)
        nfixed = self.ldjiscoverageengine.N_FIXED
        fixedcols = self.ldjiscoverageengine.FIXED_COLS
        datecols = [d.strftime('%d-%b') for d in self.ldjisworkingdays]
        allheaders = fixedcols + datecols
 
        self.ldjistable.horizontalHeader().setVisible(False)
 
        self.ldjistable.setRowCount(1 + nmodels + nparts)
        self.ldjistable.setColumnCount(len(allheaders))
        vlabels = [''] + [str(i) for i in range(1, nmodels + nparts + 1)]
        self.ldjistable.setVerticalHeaderLabels(vlabels)
 
        for col, width in enumerate([230, 80, 130, 130, 130]):
            self.ldjistable.setColumnWidth(col, width)
        for col in range(nfixed, len(allheaders)):
            self.ldjistable.setColumnWidth(col, 90)
 
        headerfont = QFont("Arial", 9, QFont.Bold)
 
        for col, label in enumerate(fixedcols):
            item = QTableWidgetItem(label)
            item.setFont(headerfont)
            item.setBackground(Qt.lightGray)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            item.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
            self.ldjistable.setItem(0, col, item)
            self.ldjistable.setSpan(0, col, nmodels + 1, 1)
 
        for col, label in enumerate(datecols, nfixed):
            item = QTableWidgetItem(label)
            item.setFont(headerfont)
            item.setBackground(Qt.lightGray)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            item.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
            self.ldjistable.setItem(0, col, item)
 
        self.ldjistable.setRowHeight(0, 35)
 
        italicfont = QFont("Arial", 9)
        italicfont.setItalic(True)
 
        for modelrow, modelname in enumerate(self.ldjismodelnames):
            tablerow = 1 + modelrow
 
            nameitem = QTableWidgetItem(f"{modelname} volume ->")
            nameitem.setFont(italicfont)
            nameitem.setForeground(Qt.darkGray)
            nameitem.setFlags(nameitem.flags() | Qt.ItemIsEditable)
            self.ldjistable.setItem(tablerow, nfixed, nameitem)
 
            for col in range(nfixed + 1, len(allheaders)):
                volitem = QTableWidgetItem("0")
                volitem.setFlags(volitem.flags() | Qt.ItemIsEditable)
                self.ldjistable.setItem(tablerow, col, volitem)
 
        for partrow, (_, row) in enumerate(self.ldjisdf.iterrows()):
            tablerow = 1 + nmodels + partrow
            fixedvals = [
                str(row['Part Number']),
                str(row['Part Description']),
                str(row['SHP Code']),
                f"{row['Moves away from COL']:,.0f}",
                f"{int(row['Last Covered Mix']):,}",
                f"{int(row['Starting COL Mix']):,}",
            ]
            for col, val in enumerate(fixedvals):
                item = QTableWidgetItem(val)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.ldjistable.setItem(tablerow, col, item)
 
            for col in range(nfixed, len(allheaders)):
                item = QTableWidgetItem("")
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.ldjistable.setItem(tablerow, col, item)
 
        self.recalculateldjiscoverage()
        self.ldjistable.itemChanged.connect(self.onldjisvolchanged)
 
    def recalculateldjiscoverage(self):
        if not hasattr(self, 'ldjisdf') or self.ldjisdf.empty:
            return
 
        try:
            self.ldjistable.itemChanged.disconnect()
        except Exception:
            pass
 
        try:
            nmodels = len(self.ldjismodelnames)
            nfixed = self.ldjiscoverageengine.N_FIXED
            ncols = self.ldjistable.columnCount()
            ndays = ncols - nfixed
 
            dayvolumes = np.zeros(ndays)
            for daycol in range(ndays):
                tablecol = nfixed + daycol
                for modelrow in range(1, 1 + nmodels):
                    item = self.ldjistable.item(modelrow, tablecol)
                    if item:
                        try:
                            dayvolumes[daycol] += float(item.text().replace(',', '').strip() or 0)
                        except ValueError:
                            pass
 
            moves_away = self.ldjisdf['Moves away from COL'].astype(float).values
            starting_mix = self.ldjisdf['Starting COL Mix'].astype(float).values
            n_parts = len(moves_away)
 
            coverage = np.empty((n_parts, ndays))
            coverage[:, 0] = starting_mix - moves_away - dayvolumes[0]
            for d in range(1, ndays):
                coverage[:, d] = coverage[:, d - 1] - dayvolumes[d]
 
            for partrow in range(n_parts):
                tablerow = 1 + nmodels + partrow
                for daycol in range(ndays):
                    val = int(coverage[partrow, daycol])
                    item = QTableWidgetItem(f"{val:,}")
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    if val <= 0:
                        item.setBackground(Qt.red)
                        item.setForeground(Qt.white)
                    else:
                        item.setBackground(Qt.white)
                        item.setForeground(Qt.black)
                    self.ldjistable.setItem(tablerow, nfixed + daycol, item)
 
        finally:
            self.ldjistable.itemChanged.connect(self.onldjisvolchanged)
 
    def onldjisvolchanged(self, item):
        nmodels = len(self.ldjismodelnames)
        nfixed = self.ldjiscoverageengine.N_FIXED
        if 1 <= item.row() <= nmodels and item.column() >= nfixed:
            self.recalculateldjiscoverage()
 
    def displaycoveragetable(self, coveragedf: pd.DataFrame):
        if coveragedf.empty:
            return
 
        try:
            self.coveragetable.itemChanged.disconnect()
        except Exception:
            pass
 
        cols = coveragedf.columns.tolist()
        n_rows, n_cols = len(coveragedf), len(cols)
        str_cols = {cols.index(c) for c in ('Part Number', 'MFG Code', 'SHP Code') if c in cols}
 
        if not hasattr(self, '_last_column_count') or self._last_column_count != n_cols:
            self.coveragetable.setColumnCount(n_cols)
            self.coveragetable.setHorizontalHeaderLabels(cols)
            self._last_column_count = n_cols
            self.comments_col = cols.index('Comments') if 'Comments' in cols else None
 
        self.coveragetable.setSortingEnabled(False)
        self.coveragetable.setUpdatesEnabled(False)
 
        try:
            self.coveragetable.setRowCount(n_rows)
 
            data = coveragedf.values
 
            daily_start = None
            for i, col in enumerate(cols):
                try:
                    datetime.strptime(col, '%m/%d')
                    daily_start = i
                    break
                except ValueError:
                    pass
 
            unit_loads = (
                coveragedf['Unit Load Qty'].fillna(1).clip(lower=1).values
                if 'Unit Load Qty' in cols
                else None
            )
 
            for row in range(n_rows):
                ulq = unit_loads[row] if unit_loads is not None else 1
 
                for col in range(n_cols):
                    value = data[row, col]
                    isnumeric = isinstance(value, (int, float)) and value == value
 
                    if isnumeric and col not in str_cols:
                        display_value = f"{value:,.0f}" if abs(value) >= 1 else f"{value:.2f}"
                    elif isinstance(value, float) and value != value:
                        display_value = ""
                    else:
                        display_value = str(value) if value is not None else ""
 
                    item = QTableWidgetItem(display_value)
                    item.setFlags(
                        (item.flags() | Qt.ItemIsEditable)
                        if col == self.comments_col
                        else (item.flags() & ~Qt.ItemIsEditable)
                    )
 
                    if daily_start is not None and col >= daily_start and isnumeric:
                        if value <= 0:
                            item.setBackground(QColor(255, 204, 204))
                            item.setForeground(QColor(156, 0, 6))
                        elif value < ulq:
                            item.setBackground(QColor(255, 255, 204))
                            item.setForeground(QColor(156, 156, 6))
 
                    self.coveragetable.setItem(row, col, item)
 
        finally:
            self.coveragetable.setUpdatesEnabled(True)
            self.coveragetable.resizeColumnsToContents()
            if self.comments_col is not None:
                self.coveragetable.setColumnWidth(self.comments_col, 200)
            self.coveragetable.setSortingEnabled(True)
            self.coveragetable.itemChanged.connect(self.oncommentchanged)
 
    def oncommentchanged(self, item):
        try:
            col = item.column()
            if 'Comments' not in self.currentcoveragedf.columns:
                return
 
            commentscol = list(self.currentcoveragedf.columns).index('Comments')
            if col != commentscol:
                return
 
            partno = str(self.currentcoveragedf.iloc[item.row()]['Part Number'])
            commenttext = item.text().strip()
 
            if self._comments_cache is None:
                self._comments_cache = self.coverageengine.loadcoveragecomments()
 
            if commenttext:
                self._comments_cache[partno] = commenttext
            else:
                self._comments_cache.pop(partno, None)
 
            self.coverageengine.savecoveragecomments(self._comments_cache)
 
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
                typeitem.setBackground(QColor(255, 204, 204))
                typeitem.setForeground(QColor(156, 0, 6))
            elif transaction['Transaction Type'] == 'GR':
                typeitem.setBackground(QColor(204, 255, 204))
                typeitem.setForeground(QColor(0, 156, 6))
 
            self.transactiontable.setItem(row, 1, typeitem)
 
            receiptitem = QTableWidgetItem(transaction['Receipt/Reqmt'])
            receiptitem.setFlags(receiptitem.flags() & ~Qt.ItemIsEditable)
            self.transactiontable.setItem(row, 2, receiptitem)
 
            qtyitem = QTableWidgetItem(f"{transaction['Available QTY']:,}")
            qtyitem.setFlags(qtyitem.flags() & ~Qt.ItemIsEditable)
 
            if transaction['Available QTY'] <= 0:
                qtyitem.setBackground(QColor(255, 204, 204))
                qtyitem.setForeground(QColor(156, 0, 6))
            elif transaction['Available QTY'] < 100:
                qtyitem.setBackground(QColor(255, 255, 204))
                qtyitem.setForeground(QColor(156, 156, 6))
 
            self.transactiontable.setItem(row, 3, qtyitem)
 
            asnitem = QTableWidgetItem(transaction['ASN'])
            asnitem.setFlags(asnitem.flags() & ~Qt.ItemIsEditable)
            self.transactiontable.setItem(row, 4, asnitem)
 
    _ALERT_COL_MAP = {
        'ALERT_DETAILS':       'Alerts',
        'PART':                'Part',
        'PART_DESCRIPTION':    'Part Description',
        'CURRENT_INVENTORY':   'Inv',
        'ON_YARD_INVENTORY':   'Yard',
        'CURRENT_REQUIREMENT': 'Req',
        'SUPPLIER_NAME':       'Supplier',
        'SUPPLY_SHP_COUNTRY':  'Country',
        'SCC_NAME':            'SCC',
        }
 
    _ALERT_EDITABLE_COLS = [
        'Reason/Cause',
        '#Cars Short',
        'Impact Date',
        'ETA',
        'TO/Container#/Air Freight/Delivery Time',
        'ASN Y/N',
        'Unloading Shop A/C/LOCD',
        ]
 
    def generatealertsbreakdown(self):
        try:
            alertdf = self.import_manager.loaddata("alert_report")
            if alertdf.empty:
                QMessageBox.warning(self, "No Data", "No Alert Report found. Please import it first.")
                return
            
            if 'ALERT_TYPE' in alertdf.columns:
                alertdf = alertdf[alertdf['ALERT_TYPE'] == 'Shortage alert']
                
            if 'ALERT_DETAILS' in alertdf.columns:
                alertdf = alertdf[alertdf['ALERT_DETAILS'].str.contains(r'Day [1-4]\b', case=False, na=False)]
                alertdf['ALERT_DETAILS'] = alertdf['ALERT_DETAILS'].str.extract(r'(Day [1-4])\b', expand=False)
                
            if alertdf.empty:
                QMessageBox.information(self, "No Data", "No Shortage alert rows for Day 1-4 found in the alert report")
                return
            
            available = {src: dst for src, dst in self._ALERT_COL_MAP.items() if src in alertdf.columns}
            displaydf = alertdf[list(available.keys())].rename(columns=available).copy()
 
            for col in self._ALERT_EDITABLE_COLS:
                displaydf[col] = ''
 
            self.originalalertsdf = displaydf.copy()
            self.populatealertfilters(displaydf)
            self.displayalertstable(displaydf)
 
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Alerts analysis failed: {str(e)}")
            
    _PIWD_COL_MAP = {
        'SCC_NAME': 'SCC Name',
        'PART': 'Part',
        'PART_DESCRIPTION': 'Part Description',
        'CURRENT_INVENTORY': 'Inv',
        'ON_YARD_INVENTORY': 'Yard',
        'PORT_INVENTORY': 'Port',
        'CURRENT_REQUIREMENT': 'Req',
        'SUPPLIER_NAME': 'Supplier',
        'SUPPLY_SHP_COUNTRY': 'Country',
        }
    
    _PIWD_EDITABLE_COLS = [
        'Qty',
        'Impact with Material in Transit (No Gap)',
        'Comments'
        ]
            
    def generatepiwdreport(self):
        try:
            piwddf = self.import_manager.loaddata("alert_report")
            if piwddf.empty:
                QMessageBox.warning(self, "No Data", "No Alert Report found. Please import it first.")
                return
            
            if 'ALERT_DETAILS' in piwddf.columns:
                piwddf = piwddf[piwddf['ALERT_DETAILS'] == 'PIWED below zero using GC ETA']
            
            if piwddf.empty:
                QMessageBox.information(self, "No Data", "No PIWD alert rows in the alert report")
                return
            
            available = {src: dst for src, dst in self._PIWD_COL_MAP.items() if src in piwddf.columns}
            displaydf = piwddf[list(available.keys())].rename(columns=available).copy()
            
            for col in self._PIWD_EDITABLE_COLS:
                displaydf[col] = ''
                
            self.originalpiwddf = displaydf.copy()
            self.populatepiwdfilters(displaydf)
            self.displaypiwdtable(displaydf)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"PIWD Report failed: {str(e)}")
 
    def populatealertfilters(self, alertdf):
        if 'SCC' in alertdf.columns:
            self.alerts_scc_filter.additems(alertdf['SCC'].dropna().unique())
 
        if 'Alerts' in alertdf.columns:
            self.alerts_type_filter.additems(alertdf['Alerts'].dropna().unique())
 
        if 'Alerts' in alertdf.columns and 'Part' in alertdf.columns:
            identifiers = (alertdf['Alerts'].astype(str) + ' - ' + alertdf['Part'].astype(str)).dropna().unique()
            self.alerts_part_filter.additems(identifiers)
 
    def applyalertfilters(self):
        if not hasattr(self, 'originalalertsdf'):
            return
 
        filtereddf = self.originalalertsdf.copy()
 
        selected_scc = self.alerts_scc_filter.getselecteditems()
        if selected_scc and 'SCC' in filtereddf.columns:
            filtereddf = filtereddf[filtereddf['SCC'].isin(selected_scc)]
 
        selected_alerts = self.alerts_type_filter.getselecteditems()
        if selected_alerts and 'Alerts' in filtereddf.columns:
            filtereddf = filtereddf[filtereddf['Alerts'].isin(selected_alerts)]
 
        selected_parts = self.alerts_part_filter.getselecteditems()
        if selected_parts and 'Alerts' in filtereddf.columns and 'Part' in filtereddf.columns:
            identifier_series = (filtereddf['Alerts'].astype(str) + ' - ' + filtereddf['Part'].astype(str))
            filtereddf = filtereddf[identifier_series.isin(selected_parts)]
 
        self.displayalertstable(filtereddf)
 
    def clearalertfilters(self):
        if hasattr(self, 'alerts_scc_filter'):
            self.alerts_scc_filter.selectallitems()
        if hasattr(self, 'alerts_type_filter'):
            self.alerts_type_filter.selectallitems()
        if hasattr(self, 'alerts_part_filter'):
            self.alerts_part_filter.selectallitems()
        if hasattr(self, 'originalalertsdf'):
            self.displayalertstable(self.originalalertsdf)
            
    def populatepiwdfilters(self, piwddf):
        if 'SCC Name' in piwddf.columns:
            self.piwd_scc_filter.additems(piwddf['SCC Name'].dropna().unique())
        
        if 'Part' in piwddf.columns:
            self.piwd_part_filter.additems(piwddf['Part'].dropna().unique())
            
    def applypiwdfilters(self):
        if not hasattr(self, 'originalpiwddf'):
            return
        
        filtereddf = self.originalpiwddf.copy()
        
        selected_scc = self.piwd_scc_filter.getselecteditems()
        if selected_scc and 'SCC Name' in filtereddf.columns:
            filtereddf = filtereddf[filtereddf['SCC Name'].isin(selected_scc)]
            
        selected_part = self.piwd_part_filter.getselecteditems()
        if selected_part and 'Part' in filtereddf.columns:
            filtereddf = filtereddf[filtereddf['Part'].isin(selected_part)]
            
        self.displaypiwdtable(filtereddf)
        
    def clearpiwdfilters(self):
        if hasattr(self, 'piwd_scc_filter'):
            self.piwd_scc_filter.selectallitems()
        if hasattr(self, 'piwd_part_filter'):
            self.piwd_part_filter.selectallitems()
 
    def displayalertstable(self, alertdf: pd.DataFrame):
        if alertdf.empty:
            return
 
        cols = alertdf.columns.tolist()
        n_rows, n_cols = len(alertdf), len(cols)
        editable_indices = {i for i, c in enumerate(cols) if c in self._ALERT_EDITABLE_COLS}
 
        self.alertstable.setSortingEnabled(False)
        self.alertstable.setUpdatesEnabled(False)
 
        try:
            self.alertstable.setColumnCount(n_cols)
            self.alertstable.setHorizontalHeaderLabels(cols)
            self.alertstable.setRowCount(n_rows)
 
            data = alertdf.values
 
            for row in range(n_rows):
                for col in range(n_cols):
                    value = data[row, col]
                    display_value = str(value) if value is not None and not (isinstance(value, float) and value != value) else ''
 
                    item = QTableWidgetItem(display_value)
                    if col in editable_indices:
                        item.setFlags(item.flags() | Qt.ItemIsEditable)
                        item.setBackground(QColor(255, 255, 230))
                    else:
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
 
                    self.alertstable.setItem(row, col, item)
 
        finally:
            self.alertstable.setUpdatesEnabled(True)
            self.alertstable.resizeColumnsToContents()
            self.alertstable.setSortingEnabled(True)
 
    def exportalertstable(self):
        if not hasattr(self, 'originalalertsdf') or self.originalalertsdf.empty:
            QMessageBox.warning(self, "No Data", "Generate alerts breakdown first.")
            return
 
        filename, _ = QFileDialog.getSaveFileName(self, "Export Alerts Breakdown", f"Alerts_Breakdown_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", "CSV files (*.csv)")
        if filename:
            try:
                cols = [self.alertstable.horizontalHeaderItem(c).text() for c in range(self.alertstable.columnCount())]
                rows = []
                for r in range(self.alertstable.rowCount()):
                    rows.append([self.alertstable.item(r, c).text() if self.alertstable.item(r, c) else '' for c in range(self.alertstable.columnCount())])
                pd.DataFrame(rows, columns=cols).to_csv(filename, index=False)
                QMessageBox.information(self, "Export Complete", f"Exported to:\n{filename}")
            except Exception as e:
                QMessageBox.critical(self, "Export Failed", str(e))
                
    def displaypiwdtable(self, piwddf: pd.DataFrame):
        if piwddf.empty:
            return
        
        cols = piwddf.columns.tolist()
        n_rows, n_cols = len(piwddf), len(cols)
        editable_indices = {i for i, c in enumerate(cols) if c in self._PIWD_EDITABLE_COLS}
        
        self.piwdtable.setSortingEnabled(False)
        self.piwdtable.setUpdatesEnabled(False)
        
        try:
            self.piwdtable.setColumnCount(n_cols)
            self.piwdtable.setHorizontalHeaderLabels(cols)
            self.piwdtable.setRowCount(n_rows)
            
            data = piwddf.values
            
            for row in range(n_rows):
                for col in range(n_cols):
                    value = data[row, col]
                    display_value = (str(value) if value is not None and not (isinstance(value, float) and value != value) else '')
                    item = QTableWidgetItem(display_value)
                    
                    if col in editable_indices:
                        item.setFlags(item.flags() | Qt.ItemIsEditable)
                        item.setBackground(QColor(255, 255, 230))
                    else:
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        
                    self.piwdtable.setItem(row, col, item)
                    
        finally:
            self.piwdtable.setUpdatesEnabled(True)
            self.piwdtable.resizeColumnsToContents()
            self.piwdtable.setSortingEnabled(True)
                  
    def exportpiwdreport(self):
        if not hasattr(self, 'originalpiwddf') or self.originalpiwddf.empty:
            QMessageBox.warning(self, "No Data", "Generate PIWD report first.")
            return
        
        filename, _ = QFileDialog.getSaveFileName(self, "Export PIWD Report", f"PIWD_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", "CSV files (*.csv)")
        if filename:
            try:
                cols = [self.piwdtable.horizontalHeaderItem(c).text() for c in range(self.piwdtable.columnCount())]
                rows = []
                for r in range(self.piwdtable.rowCount()):
                    rows.append([self.piwdtable.item(r, c).text() if self.piwdtable.item(r, c) else '' for c in range(self.piwdtable.columnCount())])
                pd.DataFrame(rows, columns=cols).to_csv(filename, index=False)
                QMessageBox.information(self, "Export Complete", f"Exported to:\n{filename}")
            except Exception as e:
                QMessageBox.critical(self, "Export Failed", str(e))
 
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
