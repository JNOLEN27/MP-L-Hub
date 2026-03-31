import json
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, Optional, Tuple
from mpl_toolkits.mplot3d import Axes3D
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import logging
import traceback as tb

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTabWidget, QTableWidget, QTableWidgetItem, QMessageBox, QScrollArea,
    QFileDialog, QComboBox, QListWidget, QListWidgetItem, QCheckBox, QFrame,
    QApplication, QLineEdit, QGridLayout, QProgressDialog, QSpinBox
)
from PyQt5.QtCore import Qt, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QColor

from app.utils.config import APPWINDOWSIZE, getsharednetworkpath
from app.data.import_manager import DataImportManager
from app.inventory_by_purpose.ibp_neural_network import InventorybyPurposeNeuralNetwork
import app.inventory_by_purpose.monte_tuc_sim as monte_tuc_sim

# Setup logging to file
LOG_FILE = Path.home() / "InventoryByPurpose_Error.log"
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


class SimpleMultiSelectFilter(QWidget):
    """Reusable multi-select filter widget"""
    selectionChanged = pyqtSignal()

    def __init__(self, placeholder="Select items...", filtertype="Item"):
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

        items_to_use = items if presorted else sorted(items)

        for item in items_to_use:
            if item and str(item).strip():
                list_item = QListWidgetItem(str(item))
                list_item.setFlags(list_item.flags() | Qt.ItemIsUserCheckable)
                list_item.setCheckState(Qt.Checked)
                self.list_widget.addItem(list_item)
                self.selected_items.add(str(item))
                self.all_items.append(str(item))

        self.update_label()

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

    def update_label(self):
        self.label.setText(f"{self.filtertype} Filter:")

    def getselecteditems(self):
        return list(self.selected_items)

    def selectallitems(self):
        self.list_widget.blockSignals(True)
        self.selected_items = set(self.all_items)
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            item.setCheckState(Qt.Checked)
        self.list_widget.blockSignals(False)
        self.update_label()


class InventorybyPurposeWindow(QMainWindow):
    def __init__(self, userdata, parent=None):
        super().__init__(parent)
        try:
            logger.info("=== Initializing InventorybyPurposeWindow ===")
            self.userdata = userdata
            logger.info(f"User data loaded: {userdata.get('username', 'unknown')}")

            logger.info("Initializing DataImportManager...")
            self.import_manager = DataImportManager()
            logger.info("DataImportManager initialized successfully")

            logger.info("Initializing InventorybyPurposeNeuralNetwork...")
            self.nn_engine = InventorybyPurposeNeuralNetwork(self.import_manager)
            logger.info("Neural network engine initialized successfully")

            # Cache for data
            self._master_data = None
            self._current_inventory = None
            self._mc_results = None

            self.setWindowTitle("Inventory by Purpose Application")
            self.resize(*APPWINDOWSIZE)

            logger.info("Setting up UI...")
            self.setupui()
            logger.info("UI setup completed successfully")

        except Exception as e:
            error_msg = f"CRITICAL ERROR in InventorybyPurposeWindow.__init__: {str(e)}\n{tb.format_exc()}"
            logger.error(error_msg)
            print(error_msg)
            # Show error to user
            try:
                QMessageBox.critical(self, "Initialization Error",
                    f"Failed to initialize Inventory by Purpose window:\n{str(e)}\n\nCheck log at: {LOG_FILE}")
            except:
                pass
            raise

    def setupui(self):
        """Setup main UI structure"""
        try:
            logger.info("Starting setupui...")
            centralwidget = QWidget()
            self.setCentralWidget(centralwidget)

            centralwidget.setStyleSheet("QWidget#centralwidget { background-color: #156082; }")
            centralwidget.setObjectName("centralwidget")

            layout = QVBoxLayout()
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(0)

            # Header
            logger.info("Creating header...")
            headerwidget = QWidget()
            headerwidget.setStyleSheet("background-color: #156082;")
            headerlayout = QVBoxLayout(headerwidget)
            headerlayout.setContentsMargins(10, 8, 10, 8)

            title = QLabel("Inventory by Purpose")
            title.setFont(QFont("Arial", 18, QFont.Bold))
            title.setAlignment(Qt.AlignCenter)
            title.setStyleSheet("color: white; background-color: transparent;")
            headerlayout.addWidget(title)

            layout.addWidget(headerwidget)

            self.setMinimumWidth(1000)

            # Tabs
            logger.info("Creating tabs...")
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

            logger.info("Creating Tied-up-capital forecast tab...")
            tieduptab = self.create_tiedup_capital_tab()
            logger.info("Creating Supplier Deep Dive tab...")
            suppliertab = self.create_supplier_deep_dive_tab()
            logger.info("Creating Strategy Analysis tab...")
            strategytab = self.create_strategy_analysis_tab()

            tabs.addTab(tieduptab, "Tied-up-capital forecast")
            tabs.addTab(suppliertab, "Supplier Deep Dive")
            tabs.addTab(strategytab, "Strategy Analysis")

            layout.addWidget(tabs)

            self.statusBar().showMessage(f"Logged in as: {self.userdata.get('username', 'unknown')}")
            centralwidget.setLayout(layout)

            logger.info("setupui completed successfully")

        except Exception as e:
            error_msg = f"ERROR in setupui: {str(e)}\n{tb.format_exc()}"
            logger.error(error_msg)
            print(error_msg)
            raise

    def create_tiedup_capital_tab(self):
        """Create Tied-up-capital forecast tab"""
        try:
            logger.info("Creating Tied-up-capital forecast tab...")
            widget = QWidget()
            layout = QVBoxLayout()

            title = QLabel("Tied-up-capital Forecast")
            title.setFont(QFont("Arial", 16, QFont.Bold))
            title.setAlignment(Qt.AlignCenter)
            layout.addWidget(title)

            # Button section
            buttonlayout = QHBoxLayout()
            generatebtn = QPushButton("Generate Forecast")
            generatebtn.clicked.connect(self.generate_tiedup_forecast)
            generatebtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 10px 20px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
            buttonlayout.addWidget(generatebtn)

            exportbtn = QPushButton("Export to CSV")
            exportbtn.clicked.connect(self.export_tiedup_tables)
            exportbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 10px 20px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
            buttonlayout.addWidget(exportbtn)

            buttonlayout.addStretch()
            layout.addLayout(buttonlayout)

            # Three tables side by side
            tableswidget = QWidget()
            tableslayout = QHBoxLayout()

            # Top 10 Parts
            partsframe = QFrame()
            partslayout = QVBoxLayout(partsframe)
            partslabel = QLabel("Top 10 Parts by Value")
            partslabel.setFont(QFont("Arial", 12, QFont.Bold))
            partslayout.addWidget(partslabel)
            self.tiedup_parts_table = QTableWidget()
            self.tiedup_parts_table.setSortingEnabled(True)
            self.tiedup_parts_table.setStyleSheet("""QTableWidget {gridline-color: #d0d0d0; background-color: white;} QTableWidget::item {padding: 6px; border: 1px solid #d0d0d0;} QHeaderView::section {background-color: #f0f0f0; padding: 8px; border: 1px solid #d0d0d0; font-weight: bold;}""")
            partslayout.addWidget(self.tiedup_parts_table)
            partsframe.setLayout(partslayout)
            tableslayout.addWidget(partsframe)

            # Top 10 Suppliers
            suppliersframe = QFrame()
            supplierlayout = QVBoxLayout(suppliersframe)
            supplierlabel = QLabel("Top 10 Suppliers by Value")
            supplierlabel.setFont(QFont("Arial", 12, QFont.Bold))
            supplierlayout.addWidget(supplierlabel)
            self.tiedup_suppliers_table = QTableWidget()
            self.tiedup_suppliers_table.setSortingEnabled(True)
            self.tiedup_suppliers_table.setStyleSheet("""QTableWidget {gridline-color: #d0d0d0; background-color: white;} QTableWidget::item {padding: 6px; border: 1px solid #d0d0d0;} QHeaderView::section {background-color: #f0f0f0; padding: 8px; border: 1px solid #d0d0d0; font-weight: bold;}""")
            supplierlayout.addWidget(self.tiedup_suppliers_table)
            suppliersframe.setLayout(supplierlayout)
            tableslayout.addWidget(suppliersframe)

            # Regions by Value
            regionsframe = QFrame()
            regionslayout = QVBoxLayout(regionsframe)
            regionslabel = QLabel("Regions by Value")
            regionslabel.setFont(QFont("Arial", 12, QFont.Bold))
            regionslayout.addWidget(regionslabel)
            self.tiedup_regions_table = QTableWidget()
            self.tiedup_regions_table.setSortingEnabled(True)
            self.tiedup_regions_table.setStyleSheet("""QTableWidget {gridline-color: #d0d0d0; background-color: white;} QTableWidget::item {padding: 6px; border: 1px solid #d0d0d0;} QHeaderView::section {background-color: #f0f0f0; padding: 8px; border: 1px solid #d0d0d0; font-weight: bold;}""")
            regionslayout.addWidget(self.tiedup_regions_table)
            regionsframe.setLayout(regionslayout)
            tableslayout.addWidget(regionsframe)

            tableswidget.setLayout(tableslayout)
            layout.addWidget(tableswidget)

            # Chart area
            chartlabel = QLabel("Monte Carlo Simulation Results")
            chartlabel.setFont(QFont("Arial", 12, QFont.Bold))
            layout.addWidget(chartlabel)

            self.tiedup_canvas = FigureCanvas(Figure(figsize=(10, 5), dpi=100))
            layout.addWidget(self.tiedup_canvas)

            widget.setLayout(layout)
            logger.info("Tied-up-capital forecast tab created successfully")
            return widget

        except Exception as e:
            error_msg = f"Error creating Tied-up-capital tab: {str(e)}\n{tb.format_exc()}"
            logger.error(error_msg)
            print(error_msg)
            raise

    def create_supplier_deep_dive_tab(self):
        """Create Supplier Deep Dive tab"""
        try:
            logger.info("Creating Supplier Deep Dive tab...")
            widget = QWidget()
            layout = QVBoxLayout()

            title = QLabel("Supplier Deep Dive")
            title.setFont(QFont("Arial", 16, QFont.Bold))
            title.setAlignment(Qt.AlignCenter)
            layout.addWidget(title)

            # Button to load data
            buttonlayout = QHBoxLayout()
            loadbtn = QPushButton("Load Supplier Data")
            loadbtn.clicked.connect(self.load_supplier_data)
            loadbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 10px 20px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
            buttonlayout.addWidget(loadbtn)
            buttonlayout.addStretch()
            layout.addLayout(buttonlayout)

            # Supplier filter - DO NOT connect signal yet
            filterlayout = QHBoxLayout()
            filterlabel = QLabel("Supplier Filter:")
            filterlabel.setFont(QFont("Arial", 12, QFont.Bold))
            filterlayout.addWidget(filterlabel)

            self.supplier_filter = self.create_multiselect_dropdown("Select Suppliers...", "Supplier")
            # Signal will be connected after filters are fully created
            filterlayout.addWidget(self.supplier_filter)

            filterlayout.addStretch()
            layout.addLayout(filterlayout)

            # Table
            self.supplier_parts_table = QTableWidget()
            self.supplier_parts_table.setSortingEnabled(True)
            self.supplier_parts_table.setStyleSheet("""QTableWidget {gridline-color: #d0d0d0; background-color: white;} QTableWidget::item {padding: 6px; border: 1px solid #d0d0d0;} QHeaderView::section {background-color: #f0f0f0; padding: 8px; border: 1px solid #d0d0d0; font-weight: bold;}""")
            layout.addWidget(self.supplier_parts_table)

            widget.setLayout(layout)
            logger.info("Supplier Deep Dive tab created successfully")
            return widget

        except Exception as e:
            error_msg = f"Error creating Supplier Deep Dive tab: {str(e)}\n{tb.format_exc()}"
            logger.error(error_msg)
            print(error_msg)
            raise

    def create_strategy_analysis_tab(self):
        """Create Strategy Analysis tab"""
        try:
            logger.info("Creating Strategy Analysis tab...")
            widget = QWidget()
            layout = QVBoxLayout()

            title = QLabel("Strategy Analysis")
            title.setFont(QFont("Arial", 16, QFont.Bold))
            title.setAlignment(Qt.AlignCenter)
            layout.addWidget(title)

            # Button to load data
            buttonlayout = QHBoxLayout()
            loadbtn = QPushButton("Load Analysis Data")
            loadbtn.clicked.connect(self.load_strategy_data)
            loadbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 10px 20px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
            buttonlayout.addWidget(loadbtn)
            buttonlayout.addStretch()
            layout.addLayout(buttonlayout)

            # Filter section
            filterlabel = QLabel("Filters:")
            filterlabel.setFont(QFont("Arial", 12, QFont.Bold))
            layout.addWidget(filterlabel)

            filtergrid = QGridLayout()
            filtergrid.setSpacing(10)

            # Create filters WITHOUT connecting signals yet
            self.strategy_part_filter = self.create_multiselect_dropdown("Select Parts...", "Part")
            filtergrid.addWidget(self.strategy_part_filter, 0, 0)

            self.strategy_supplier_filter = self.create_multiselect_dropdown("Select Suppliers...", "Supplier")
            filtergrid.addWidget(self.strategy_supplier_filter, 0, 1)

            self.strategy_region_filter = self.create_multiselect_dropdown("Select Regions...", "Region")
            filtergrid.addWidget(self.strategy_region_filter, 0, 2)

            self.strategy_country_filter = self.create_multiselect_dropdown("Select Countries...", "Country")
            filtergrid.addWidget(self.strategy_country_filter, 1, 0)

            self.strategy_scc_filter = self.create_multiselect_dropdown("Select SCC...", "SCC")
            filtergrid.addWidget(self.strategy_scc_filter, 1, 1)

            clearfilterbtn = QPushButton("Clear Filters")
            clearfilterbtn.clicked.connect(self.clear_strategy_filters)
            clearfilterbtn.setStyleSheet("""QPushButton {background-color: #E97132; color: white; padding: 8px 16px; border: none; border-radius: 5px;} QPushButton:hover {background-color: #da190b;}""")
            filtergrid.addWidget(clearfilterbtn, 1, 2)

            layout.addLayout(filtergrid)

            # 3D Scatter plot
            plotlabel = QLabel("3D Scatter Plot: SAFETY vs STOCK vs Price")
            plotlabel.setFont(QFont("Arial", 12, QFont.Bold))
            layout.addWidget(plotlabel)

            self.strategy_canvas = FigureCanvas(Figure(figsize=(10, 6), dpi=100))
            layout.addWidget(self.strategy_canvas)

            # Table
            tablelabel = QLabel("Filtered Parts Data")
            tablelabel.setFont(QFont("Arial", 12, QFont.Bold))
            layout.addWidget(tablelabel)

            self.strategy_table = QTableWidget()
            self.strategy_table.setSortingEnabled(True)
            self.strategy_table.setStyleSheet("""QTableWidget {gridline-color: #d0d0d0; background-color: white;} QTableWidget::item {padding: 6px; border: 1px solid #d0d0d0;} QHeaderView::section {background-color: #f0f0f0; padding: 8px; border: 1px solid #d0d0d0; font-weight: bold;}""")
            layout.addWidget(self.strategy_table)

            widget.setLayout(layout)
            logger.info("Strategy Analysis tab created successfully")
            return widget

        except Exception as e:
            error_msg = f"Error creating Strategy Analysis tab: {str(e)}\n{tb.format_exc()}"
            logger.error(error_msg)
            print(error_msg)
            raise

    def create_multiselect_dropdown(self, placeholdertext, filtertype="Item"):
        """Create a multi-select filter widget"""
        try:
            logger.info(f"Creating multiselect dropdown for {filtertype}")
            return SimpleMultiSelectFilter(placeholdertext, filtertype)
        except Exception as e:
            error_msg = f"Error creating multiselect dropdown for {filtertype}: {str(e)}\n{tb.format_exc()}"
            logger.error(error_msg)
            print(error_msg)
            raise

    def load_required_data(self):
        """Load all required data for analysis"""
        try:
            if self._master_data is None:
                logger.info("Loading master_data...")
                self._master_data = self.import_manager.loaddata("master_data")
                logger.info(f"Master data loaded: {len(self._master_data)} rows")

            if self._current_inventory is None:
                logger.info("Loading current_inventory_report...")
                self._current_inventory = self.import_manager.loaddata("current_inventory_report")
                logger.info(f"Current inventory loaded: {len(self._current_inventory)} rows")

            if self._master_data.empty or self._current_inventory.empty:
                msg = "Missing required data (master_data or current_inventory_report)"
                logger.warning(msg)
                return False, msg

            logger.info("All required data loaded successfully")
            return True, "Data loaded successfully"
        except Exception as e:
            error_msg = f"Error loading required data: {str(e)}\n{tb.format_exc()}"
            logger.error(error_msg)
            print(error_msg)
            return False, str(e)

    def generate_tiedup_forecast(self):
        """Generate tied-up capital forecast using Monte Carlo"""
        success, message = self.load_required_data()
        if not success:
            QMessageBox.warning(self, "Missing Data", message)
            return

        try:
            progress = QProgressDialog("Running Monte Carlo simulation...", "Cancel", 0, 100, self)
            progress.setWindowTitle("Processing")
            progress.setValue(50)
            progress.show()
            QApplication.processEvents()

            # Get top parts, suppliers, and regions by value
            top_parts = self.compute_top_parts_by_value()
            top_suppliers = self.compute_top_suppliers_by_value()
            regions_value = self.compute_regions_by_value()

            # Display tables
            self.display_top_parts_table(top_parts)
            self.display_top_suppliers_table(top_suppliers)
            self.display_regions_table(regions_value)

            progress.close()
            QMessageBox.information(self, "Success", "Forecast generated successfully")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Forecast generation failed: {str(e)}")

    def compute_top_parts_by_value(self):
        """Compute top 10 parts by value"""
        try:
            inv = self._current_inventory.copy()
            if 'PRICE' not in inv.columns:
                print(f"Available columns: {list(inv.columns)}")
                return pd.DataFrame()

            # Find quantity column (could be BEGINNING_INVENTORY_TODAY, QOH_TOTAL, etc)
            qty_cols = [c for c in inv.columns if 'INVENTORY' in c or 'QOH' in c or 'QUANTITY' in c]
            if not qty_cols:
                print(f"No quantity column found. Available: {list(inv.columns)}")
                return pd.DataFrame()

            qty_col = qty_cols[0]

            # Find part number and description columns
            part_col = next((c for c in inv.columns if c in ['PART_NO', 'PART', 'PART_NUMBER']), None)
            desc_col = next((c for c in inv.columns if 'DESC' in c), None)

            if not part_col or not desc_col:
                print(f"Could not find part/desc columns. Available: {list(inv.columns)}")
                return pd.DataFrame()

            inv['Value'] = (inv['PRICE'] * inv[qty_col]).fillna(0)
            top = inv.nlargest(10, 'Value')[[part_col, desc_col, 'Value']]
            top.columns = ['Part Number', 'Description', 'Value']
            return top
        except Exception as e:
            print(f"Error computing top parts: {e}")
            import traceback
            traceback.print_exc()
            return pd.DataFrame()

    def compute_top_suppliers_by_value(self):
        """Compute top 10 suppliers by value"""
        try:
            inv = self._current_inventory.copy()
            if 'PRICE' not in inv.columns:
                return pd.DataFrame()

            # Find supplier column
            supp_col = next((c for c in inv.columns if 'SUPP' in c and 'NAME' in c), None)
            if not supp_col:
                return pd.DataFrame()

            # Find quantity column
            qty_cols = [c for c in inv.columns if 'INVENTORY' in c or 'QOH' in c or 'QUANTITY' in c]
            if not qty_cols:
                return pd.DataFrame()
            qty_col = qty_cols[0]

            inv['Value'] = (inv['PRICE'] * inv[qty_col]).fillna(0)
            supplier_value = inv.groupby(supp_col)['Value'].sum().reset_index()
            top = supplier_value.nlargest(10, 'Value')
            top.columns = ['Supplier', 'Value']
            return top
        except Exception as e:
            print(f"Error computing top suppliers: {e}")
            import traceback
            traceback.print_exc()
            return pd.DataFrame()

    def compute_regions_by_value(self):
        """Compute regions by value"""
        try:
            inv = self._current_inventory.copy()

            # Find country column
            country_col = next((c for c in inv.columns if 'COUNTRY' in c or 'SHIP_COUNTRY' in c), None)
            if not country_col:
                print(f"No country column found. Available: {list(inv.columns)}")
                return pd.DataFrame()

            # Find quantity column
            qty_cols = [c for c in inv.columns if 'INVENTORY' in c or 'QOH' in c or 'QUANTITY' in c]
            if not qty_cols or 'PRICE' not in inv.columns:
                return pd.DataFrame()
            qty_col = qty_cols[0]

            inv['Value'] = (inv['PRICE'] * inv[qty_col]).fillna(0)
            inv['Region'] = inv[country_col].apply(self.determine_region)
            region_value = inv.groupby('Region')['Value'].sum().reset_index().sort_values('Value', ascending=False)
            region_value.columns = ['Region', 'Value']
            return region_value
        except Exception as e:
            print(f"Error computing regions: {e}")
            import traceback
            traceback.print_exc()
            return pd.DataFrame()

    def determine_region(self, country):
        """Determine region from country"""
        if pd.isna(country) or not country:
            return "Unknown"
        country = str(country).upper().strip()

        if country == 'USA':
            return 'USA'
        if country in ('MEXICO', 'CANADA'):
            return 'MEX'
        if country in ('AUSTRIA', 'BELGIUM', 'BULGARIA', 'CZECH REPUBLIC', 'DENMARK', 'FRANCE',
                      'GERMANY', 'HUNGARY', 'IRELAND', 'ITALY', 'LITHUANIA', 'MOROCCO',
                      'NETHERLANDS', 'NORWAY', 'POLAND', 'PORTUGAL', 'ROMANIA', 'SLOVAK REPUBLIC',
                      'SLOVENIA', 'SPAIN', 'SWEDEN', 'SWITZERLAND', 'TUNISIA', 'TURKEY',
                      'UKRAINE', 'UNITED KINGDOM'):
            return "EMEA"
        if country in ('CHINA', 'SOUTH KOREA', 'THAILAND', 'VIETNAM'):
            return "APAC"
        return 'Other'

    def display_top_parts_table(self, df):
        """Display top parts table"""
        if df.empty:
            self.tiedup_parts_table.setRowCount(0)
            return

        self.tiedup_parts_table.setRowCount(len(df))
        self.tiedup_parts_table.setColumnCount(3)
        self.tiedup_parts_table.setHorizontalHeaderLabels(['Part Number', 'Description', 'Value'])

        for row, (_, rowdata) in enumerate(df.iterrows()):
            for col, colname in enumerate(df.columns):
                value = rowdata[colname]
                if colname == 'Value' and isinstance(value, (int, float)):
                    display_value = f"${float(value):,.2f}"
                else:
                    display_value = str(value)

                item = QTableWidgetItem(display_value)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.tiedup_parts_table.setItem(row, col, item)

        self.tiedup_parts_table.resizeColumnsToContents()

    def display_top_suppliers_table(self, df):
        """Display top suppliers table"""
        if df.empty:
            return

        self.tiedup_suppliers_table.setRowCount(len(df))
        self.tiedup_suppliers_table.setColumnCount(2)
        self.tiedup_suppliers_table.setHorizontalHeaderLabels(['Supplier', 'Value'])

        for row, (_, rowdata) in enumerate(df.iterrows()):
            supplieritem = QTableWidgetItem(str(rowdata['Supplier']))
            supplieritem.setFlags(supplieritem.flags() & ~Qt.ItemIsEditable)
            self.tiedup_suppliers_table.setItem(row, 0, supplieritem)

            valueitem = QTableWidgetItem(f"${float(rowdata['Value']):,.2f}")
            valueitem.setFlags(valueitem.flags() & ~Qt.ItemIsEditable)
            self.tiedup_suppliers_table.setItem(row, 1, valueitem)

        self.tiedup_suppliers_table.resizeColumnsToContents()

    def display_regions_table(self, df):
        """Display regions table"""
        if df.empty:
            return

        self.tiedup_regions_table.setRowCount(len(df))
        self.tiedup_regions_table.setColumnCount(2)
        self.tiedup_regions_table.setHorizontalHeaderLabels(['Region', 'Value'])

        for row, (_, rowdata) in enumerate(df.iterrows()):
            regionitem = QTableWidgetItem(str(rowdata['Region']))
            regionitem.setFlags(regionitem.flags() & ~Qt.ItemIsEditable)
            self.tiedup_regions_table.setItem(row, 0, regionitem)

            valueitem = QTableWidgetItem(f"${float(rowdata['Value']):,.2f}")
            valueitem.setFlags(valueitem.flags() & ~Qt.ItemIsEditable)
            self.tiedup_regions_table.setItem(row, 1, valueitem)

        self.tiedup_regions_table.resizeColumnsToContents()

    def export_tiedup_tables(self):
        """Export tied-up tables to CSV"""
        if self.tiedup_parts_table.rowCount() == 0:
            QMessageBox.warning(self, "No Data", "Generate forecast first.")
            return

        filename, _ = QFileDialog.getSaveFileName(
            self, "Export Tied-up Capital Forecast",
            f"TiedUp_Forecast_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV files (*.csv)"
        )
        if filename:
            try:
                # Collect all data from tables
                data = []
                data.append("=== TOP 10 PARTS BY VALUE ===")
                for row in range(self.tiedup_parts_table.rowCount()):
                    rowdata = [self.tiedup_parts_table.item(row, col).text() for col in range(3)]
                    data.append(",".join(rowdata))

                data.append("\n=== TOP 10 SUPPLIERS BY VALUE ===")
                for row in range(self.tiedup_suppliers_table.rowCount()):
                    rowdata = [self.tiedup_suppliers_table.item(row, col).text() for col in range(2)]
                    data.append(",".join(rowdata))

                data.append("\n=== REGIONS BY VALUE ===")
                for row in range(self.tiedup_regions_table.rowCount()):
                    rowdata = [self.tiedup_regions_table.item(row, col).text() for col in range(2)]
                    data.append(",".join(rowdata))

                with open(filename, 'w') as f:
                    f.write("\n".join(data))

                QMessageBox.information(self, "Export Complete", f"Exported to:\n{filename}")
            except Exception as e:
                QMessageBox.critical(self, "Export Failed", str(e))

    def load_supplier_data(self):
        """Load supplier data and populate filters"""
        success, message = self.load_required_data()
        if not success:
            QMessageBox.warning(self, "Missing Data", message)
            return

        try:
            # Find supplier column
            supp_col = next((c for c in self._master_data.columns if 'SUPP' in c and 'NAME' in c), None)
            if not supp_col:
                raise ValueError("Could not find supplier column in master data")

            suppliers = self._master_data[supp_col].dropna().unique()
            self.supplier_filter.additems(suppliers)

            # NOW connect the signal after filters are populated
            self.supplier_filter.selectionChanged.connect(self.apply_supplier_filter)

            QMessageBox.information(self, "Success", f"Loaded {len(suppliers)} suppliers")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load suppliers: {str(e)}")
            import traceback
            traceback.print_exc()

    def load_strategy_data(self):
        """Load strategy analysis data and populate filters"""
        success, message = self.load_required_data()
        if not success:
            QMessageBox.warning(self, "Missing Data", message)
            return

        try:
            # Find the actual column names in master data
            part_col = next((c for c in self._master_data.columns if c in ['PART', 'PART_NO', 'PART_NUMBER']), None)
            supp_col = next((c for c in self._master_data.columns if 'SUPP' in c and 'NAME' in c), None)
            country_col = next((c for c in self._master_data.columns if 'COUNTRY' in c or 'SHIP_COUNTRY' in c), None)
            scc_col = next((c for c in self._master_data.columns if 'SCC' in c), None)

            if not all([part_col, supp_col, country_col, scc_col]):
                missing = [n for n, c in [('Part', part_col), ('Supplier', supp_col), ('Country', country_col), ('SCC', scc_col)] if c is None]
                raise ValueError(f"Missing columns: {missing}")

            parts = self._master_data[part_col].dropna().unique()
            suppliers = self._master_data[supp_col].dropna().unique()
            countries = self._master_data[country_col].dropna().unique()
            sccs = self._master_data[scc_col].dropna().unique()

            # Compute regions
            self._master_data['Region'] = self._master_data[country_col].apply(self.determine_region)
            regions = self._master_data['Region'].dropna().unique()

            # Populate filters
            self.strategy_part_filter.additems(parts)
            self.strategy_supplier_filter.additems(suppliers)
            self.strategy_region_filter.additems(regions)
            self.strategy_country_filter.additems(countries)
            self.strategy_scc_filter.additems(sccs)

            # NOW connect signals after filters are populated
            try:
                self.strategy_part_filter.selectionChanged.disconnect()
            except:
                pass
            try:
                self.strategy_supplier_filter.selectionChanged.disconnect()
            except:
                pass
            try:
                self.strategy_region_filter.selectionChanged.disconnect()
            except:
                pass
            try:
                self.strategy_country_filter.selectionChanged.disconnect()
            except:
                pass
            try:
                self.strategy_scc_filter.selectionChanged.disconnect()
            except:
                pass

            self.strategy_part_filter.selectionChanged.connect(self.apply_strategy_filters)
            self.strategy_supplier_filter.selectionChanged.connect(self.apply_strategy_filters)
            self.strategy_region_filter.selectionChanged.connect(self.apply_strategy_filters)
            self.strategy_country_filter.selectionChanged.connect(self.apply_strategy_filters)
            self.strategy_scc_filter.selectionChanged.connect(self.apply_strategy_filters)

            # Display all data initially
            self.display_3d_scatterplot(self._master_data)
            self.display_strategy_table(self._master_data)

            QMessageBox.information(self, "Success", "Loaded strategy analysis data")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load strategy data: {str(e)}")
            import traceback
            traceback.print_exc()

    def apply_supplier_filter(self):
        """Apply supplier filter and display parts"""
        success, message = self.load_required_data()
        if not success:
            return

        try:
            # Find supplier column
            supp_col = next((c for c in self._master_data.columns if 'SUPP' in c and 'NAME' in c), None)
            if not supp_col:
                return

            selected_suppliers = self.supplier_filter.getselecteditems()
            if not selected_suppliers:
                return

            # Get master data and filter by selected suppliers
            master = self._master_data.copy()
            filtered = master[master[supp_col].isin(selected_suppliers)]

            if filtered.empty:
                self.supplier_parts_table.setRowCount(0)
                return

            # Display parts for selected suppliers
            self.display_supplier_parts_table(filtered)

        except Exception as e:
            print(f"Filter error: {e}")
            import traceback
            traceback.print_exc()

    def display_supplier_parts_table(self, df):
        """Display parts for selected supplier"""
        if df.empty:
            self.supplier_parts_table.setRowCount(0)
            return

        try:
            # Find the actual column names
            part_col = next((c for c in df.columns if c in ['PART', 'PART_NO', 'PART_NUMBER']), None)
            desc_col = next((c for c in df.columns if 'DESC' in c), None)
            supp_col = next((c for c in df.columns if 'SUPP' in c and 'NAME' in c), None)
            price_col = next((c for c in df.columns if 'PRICE' in c), None)

            # Build column list from what's available
            cols = []
            col_map = {}
            if part_col:
                cols.append(part_col)
                col_map[part_col] = 'Part Number'
            if desc_col:
                cols.append(desc_col)
                col_map[desc_col] = 'Description'
            if supp_col:
                cols.append(supp_col)
                col_map[supp_col] = 'Supplier'
            if price_col:
                cols.append(price_col)
                col_map[price_col] = 'Price'

            if not cols:
                self.supplier_parts_table.setRowCount(0)
                return

            display_cols = [col_map[c] for c in cols]

            self.supplier_parts_table.setRowCount(len(df))
            self.supplier_parts_table.setColumnCount(len(cols))
            self.supplier_parts_table.setHorizontalHeaderLabels(display_cols)

            for row, (_, rowdata) in enumerate(df[cols].iterrows()):
                for col, colname in enumerate(cols):
                    value = rowdata[colname]
                    if colname == price_col and isinstance(value, (int, float)):
                        display_value = f"${float(value):,.2f}"
                    else:
                        display_value = str(value)

                    item = QTableWidgetItem(display_value)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.supplier_parts_table.setItem(row, col, item)

            self.supplier_parts_table.resizeColumnsToContents()
        except Exception as e:
            print(f"Error displaying supplier parts table: {e}")
            import traceback
            traceback.print_exc()

    def apply_strategy_filters(self):
        """Apply all strategy filters and update plot/table"""
        success, message = self.load_required_data()
        if not success:
            return

        try:
            # Find the actual column names in master data
            part_col = next((c for c in self._master_data.columns if c in ['PART', 'PART_NO', 'PART_NUMBER']), None)
            supp_col = next((c for c in self._master_data.columns if 'SUPP' in c and 'NAME' in c), None)
            country_col = next((c for c in self._master_data.columns if 'COUNTRY' in c or 'SHIP_COUNTRY' in c), None)
            scc_col = next((c for c in self._master_data.columns if 'SCC' in c), None)

            if not all([part_col, supp_col, country_col, scc_col]):
                return

            selected_parts = self.strategy_part_filter.getselecteditems()
            selected_suppliers = self.strategy_supplier_filter.getselecteditems()
            selected_regions = self.strategy_region_filter.getselecteditems()
            selected_countries = self.strategy_country_filter.getselecteditems()
            selected_sccs = self.strategy_scc_filter.getselecteditems()

            # Filter master data
            filtered = self._master_data.copy()

            if selected_parts:
                filtered = filtered[filtered[part_col].isin(selected_parts)]
            if selected_suppliers:
                filtered = filtered[filtered[supp_col].isin(selected_suppliers)]
            if selected_regions:
                # Need to determine region from country
                filtered['Region'] = filtered[country_col].apply(self.determine_region)
                filtered = filtered[filtered['Region'].isin(selected_regions)]
            if selected_countries:
                filtered = filtered[filtered[country_col].isin(selected_countries)]
            if selected_sccs:
                filtered = filtered[filtered[scc_col].isin(selected_sccs)]

            if filtered.empty:
                self.strategy_table.setRowCount(0)
                return

            # Update 3D plot and table
            self.display_3d_scatterplot(filtered)
            self.display_strategy_table(filtered)

        except Exception as e:
            print(f"Filter error: {e}")
            import traceback
            traceback.print_exc()

    def display_3d_scatterplot(self, df):
        """Display interactive 3D scatter plot"""
        if df.empty:
            return

        try:
            # Find the actual column names
            safety_col = next((c for c in df.columns if 'SAFETY' in c), None)
            stock_col = next((c for c in df.columns if 'STOCK' in c or 'BEGINNING' in c), None)
            price_col = next((c for c in df.columns if 'PRICE' in c), None)

            if not all([safety_col, stock_col, price_col]):
                missing = [n for n, c in [('SAFETY', safety_col), ('STOCK', stock_col), ('PRICE', price_col)] if c is None]
                print(f"Cannot plot 3D: Missing columns {missing}. Available: {list(df.columns)}")
                self.strategy_canvas.figure.clear()
                ax = self.strategy_canvas.figure.add_subplot(111)
                ax.text(0.5, 0.5, f'Missing columns: {missing}', ha='center', va='center', transform=ax.transAxes)
                self.strategy_canvas.draw()
                return

            self.strategy_canvas.figure.clear()
            ax = self.strategy_canvas.figure.add_subplot(111, projection='3d')

            # Extract data
            safety = df[safety_col].fillna(0).astype(float).values
            stock = df[stock_col].fillna(0).astype(float).values
            price = df[price_col].fillna(0).astype(float).values

            # Create scatter plot
            scatter = ax.scatter(safety, stock, price, c=price, cmap='viridis', marker='o', s=50, alpha=0.6)

            ax.set_xlabel('SAFETY', fontsize=10)
            ax.set_ylabel('STOCK', fontsize=10)
            ax.set_zlabel('Price', fontsize=10)
            ax.set_title('3D Strategy Analysis', fontsize=12, fontweight='bold')

            self.strategy_canvas.figure.colorbar(scatter, ax=ax, label='Price')
            self.strategy_canvas.draw()

        except Exception as e:
            print(f"Error displaying 3D plot: {e}")
            import traceback
            traceback.print_exc()

    def display_strategy_table(self, df):
        """Display strategy analysis table"""
        if df.empty:
            self.strategy_table.setRowCount(0)
            return

        try:
            # Find the actual column names
            part_col = next((c for c in df.columns if c in ['PART', 'PART_NO', 'PART_NUMBER']), None)
            desc_col = next((c for c in df.columns if 'DESC' in c), None)
            supp_col = next((c for c in df.columns if 'SUPP' in c and 'NAME' in c), None)
            safety_col = next((c for c in df.columns if 'SAFETY' in c), None)
            stock_col = next((c for c in df.columns if 'STOCK' in c or 'BEGINNING' in c), None)

            # Build column list from what's available
            cols = []
            col_map = {}
            if part_col:
                cols.append(part_col)
                col_map[part_col] = 'Part Number'
            if desc_col:
                cols.append(desc_col)
                col_map[desc_col] = 'Description'
            if supp_col:
                cols.append(supp_col)
                col_map[supp_col] = 'Supplier'
            if safety_col:
                cols.append(safety_col)
                col_map[safety_col] = 'SAFETY'
            if stock_col:
                cols.append(stock_col)
                col_map[stock_col] = 'STOCK'

            if not cols:
                self.strategy_table.setRowCount(0)
                return

            display_cols = [col_map[c] for c in cols]

            self.strategy_table.setRowCount(len(df))
            self.strategy_table.setColumnCount(len(cols))
            self.strategy_table.setHorizontalHeaderLabels(display_cols)

            for row, (_, rowdata) in enumerate(df[cols].iterrows()):
                for col, colname in enumerate(cols):
                    value = rowdata[colname]
                    if colname in (safety_col, stock_col) and isinstance(value, (int, float)):
                        display_value = f"{float(value):,.2f}"
                    else:
                        display_value = str(value)

                    item = QTableWidgetItem(display_value)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.strategy_table.setItem(row, col, item)

            self.strategy_table.resizeColumnsToContents()
        except Exception as e:
            print(f"Error displaying strategy table: {e}")
            import traceback
            traceback.print_exc()

    def clear_strategy_filters(self):
        """Clear all strategy filters"""
        if hasattr(self, 'strategy_part_filter'):
            self.strategy_part_filter.selectallitems()
        if hasattr(self, 'strategy_supplier_filter'):
            self.strategy_supplier_filter.selectallitems()
        if hasattr(self, 'strategy_region_filter'):
            self.strategy_region_filter.selectallitems()
        if hasattr(self, 'strategy_country_filter'):
            self.strategy_country_filter.selectallitems()
        if hasattr(self, 'strategy_scc_filter'):
            self.strategy_scc_filter.selectallitems()

        # Refresh display
        self.apply_strategy_filters()

    def closeEvent(self, event):
        """Handle window close event"""
        reply = QMessageBox.question(self, "Exit", "Are you sure you want to exit?",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()
