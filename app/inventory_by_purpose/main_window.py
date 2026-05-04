import json
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, Optional, Tuple
import logging
import traceback as tb

LOG_FILE = Path.home() / "InventoryByPurpose_Error.log"
try:
    logging.basicConfig(
        filename=LOG_FILE,
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
except Exception as e:
    print(f"WARNING: Could not set up logging: {e}")

logger = logging.getLogger(__name__)
logger.info("=== Starting imports for InventorybyPurposeWindow ===")

try:
    logger.info("Importing PyQt5...")
    from PyQt5.QtWidgets import (
        QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
        QTabWidget, QTableWidget, QTableWidgetItem, QMessageBox, QScrollArea,
        QFileDialog, QComboBox, QListWidget, QListWidgetItem, QCheckBox, QFrame,
        QApplication, QLineEdit, QGridLayout, QProgressDialog, QSpinBox, QSizePolicy
    )
    from PyQt5.QtCore import Qt, pyqtSignal, QTimer, QThread
    from PyQt5.QtGui import QFont, QColor, QFontMetrics
    logger.info("PyQt5 imported successfully")

    logger.info("Importing app modules...")
    from app.utils.config import APPWINDOWSIZE, getsharednetworkpath
    from app.data.import_manager import DataImportManager
    try:
        from app.inventory_by_purpose.ibp_neural_network import InventorybyPurposeNeuralNetwork
        TORCH_AVAILABLE = True
        logger.info("PyTorch/Neural network imported successfully")
    except ModuleNotFoundError as e:
        TORCH_AVAILABLE = False
        InventorybyPurposeNeuralNetwork = None
        logger.warning(f"PyTorch not available - neural network features disabled: {e}")
    logger.info("App modules imported successfully")

    logger.info("=== All imports completed successfully ===")

except Exception as e:
    error_msg = f"=== CRITICAL ERROR DURING IMPORTS ===\n{str(e)}\n{tb.format_exc()}"
    logger.error(error_msg)
    print(error_msg)
    try:
        from PyQt5.QtWidgets import QMessageBox, QApplication
        app = QApplication([])
        QMessageBox.critical(None, "Import Error", f"Failed to import modules:\n{str(e)}\n\nCheck log at: {LOG_FILE}")
    except:
        pass
    raise

_MC_SIM_THREAD_CLASS = None

def _get_mc_sim_thread_class():
    global _MC_SIM_THREAD_CLASS
    if _MC_SIM_THREAD_CLASS is not None:
        return _MC_SIM_THREAD_CLASS

    class MCSimThread(QThread):
        progress = pyqtSignal(int, str)   
        finished = pyqtSignal(dict)       
        error    = pyqtSignal(str)      

        def __init__(self, import_manager, days: int = 90, n_sims: int = 50):
            super().__init__()
            self.import_manager = import_manager
            self.days = days
            self.n_sims = n_sims
            self._cancelled = False

        def cancel(self):
            self._cancelled = True

        def run(self):
            try:
                try:
                    import app.inventory_by_purpose.monte_tuc_sim as _mc
                except (ImportError, ModuleNotFoundError) as ie:
                    logger.warning(f"MCSimThread: Monte Carlo module unavailable: {ie}")
                    self.finished.emit({})
                    return

                self.progress.emit(10, "Loading simulation data...")
                mc_data = _mc.load_required_data(self.import_manager)
                logger.info("MCSimThread: data loaded")

                if self._cancelled:
                    logger.info("MCSimThread: cancelled before simulation started")
                    self.finished.emit({})
                    return

                self.progress.emit(30, "Running Monte Carlo simulation\n(this may take several minutes)...")
                forecast_result = _mc.plant_forecast_shared_pva(
                    self.days, mc_data, n_sims=self.n_sims
                )
                logger.info("MCSimThread: simulation complete")
                self.progress.emit(95, "Preparing chart...")
                self.finished.emit(forecast_result)

            except Exception as e:
                logger.error(f"MCSimThread error: {e}\n{tb.format_exc()}")
                self.error.emit(str(e))

    _MC_SIM_THREAD_CLASS = MCSimThread
    return _MC_SIM_THREAD_CLASS

_SIMPLE_FILTER_CLASS = None

def _get_simple_filter_class():
    global _SIMPLE_FILTER_CLASS
    if _SIMPLE_FILTER_CLASS is not None:
        return _SIMPLE_FILTER_CLASS

    class SimpleMultiSelectFilter(QWidget):
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

            BATCH_SIZE = 100
            for batch_start in range(0, len(items_to_use), BATCH_SIZE):
                batch = items_to_use[batch_start:batch_start + BATCH_SIZE]
                for item in batch:
                    if item and str(item).strip():
                        list_item = QListWidgetItem(str(item))
                        list_item.setFlags(list_item.flags() | Qt.ItemIsUserCheckable)
                        list_item.setCheckState(Qt.Checked)
                        self.list_widget.addItem(list_item)
                        self.selected_items.add(str(item))
                        self.all_items.append(str(item))
                QApplication.processEvents()

            self.update_label()

            fm = QFontMetrics(self.list_widget.font())
            all_texts = [f"✓ All {self.filtertype}'s"] + [str(i) for i in items_to_use if i and str(i).strip()]
            max_text_w = max((fm.boundingRect(t).width() for t in all_texts), default=60)
            target_w = max_text_w + 58
            target_w = max(80, min(target_w, 300))
            self.setFixedWidth(target_w)

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

    _SIMPLE_FILTER_CLASS = SimpleMultiSelectFilter
    return _SIMPLE_FILTER_CLASS


class InventorybyPurposeWindow(QMainWindow):
    def __init__(self, userdata, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Inventory by Purpose Application")
        def _flush():
            for h in logging.root.handlers:
                try: h.flush()
                except Exception: pass

        try:
            logger.info("=== Initializing InventorybyPurposeWindow ==="); _flush()
            logger.info("setWindowTitle OK — init continues"); _flush()
            self.userdata = userdata
            logger.info(f"User data loaded: {userdata.get('username', 'unknown')}"); _flush()

            logger.info("Initializing DataImportManager..."); _flush()
            self.import_manager = DataImportManager()
            logger.info("DataImportManager initialized successfully"); _flush()

            logger.info("Initializing InventorybyPurposeNeuralNetwork..."); _flush()
            if TORCH_AVAILABLE and InventorybyPurposeNeuralNetwork is not None:
                self.nn_engine = InventorybyPurposeNeuralNetwork(self.import_manager)
                logger.info("Neural network engine initialized successfully"); _flush()
            else:
                self.nn_engine = None
                logger.warning("Neural network engine unavailable (torch not installed)"); _flush()

            self._master_data = None
            self._current_inventory = None
            self._mc_results = None

            self._mc_thread = None
            self._mc_progress_dialog = None
            logger.info("Attributes set, calling resize..."); _flush()
            self.resize(*APPWINDOWSIZE)
            logger.info("resize OK"); _flush()
            self.setAttribute(Qt.WA_DeleteOnClose)
            logger.info("setAttribute OK"); _flush()

            logger.info("Setting up UI..."); _flush()
            self.setupui()
            logger.info("UI setup completed successfully"); _flush()

        except Exception as e:
            error_msg = f"CRITICAL ERROR in InventorybyPurposeWindow.__init__: {str(e)}\n{tb.format_exc()}"
            logger.error(error_msg)
            print(error_msg)
            try:
                QMessageBox.critical(self, "Initialization Error",
                    f"Failed to initialize Inventory by Purpose window:\n{str(e)}\n\nCheck log at: {LOG_FILE}")
            except:
                pass
            raise

    def setupui(self):
        def _flush():
            for h in logging.root.handlers:
                try: h.flush()
                except Exception: pass
        try:
            logger.info("Starting setupui..."); _flush()
            centralwidget = QWidget()
            logger.info("centralwidget created"); _flush()
            self.setCentralWidget(centralwidget)

            centralwidget.setStyleSheet("QWidget#centralwidget { background-color: #156082; }")
            centralwidget.setObjectName("centralwidget")

            layout = QVBoxLayout()
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(0)

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
        try:
            logger.info("Creating Tied-up-capital forecast tab...")
            widget = QWidget()
            layout = QVBoxLayout()

            title = QLabel("Tied-up-capital Forecast")
            title.setFont(QFont("Arial", 16, QFont.Bold))
            title.setAlignment(Qt.AlignCenter)
            layout.addWidget(title)

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

            tableswidget = QWidget()
            tableslayout = QHBoxLayout()

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

            chartlabel = QLabel("Monte Carlo Simulation Results")
            chartlabel.setFont(QFont("Arial", 12, QFont.Bold))
            layout.addWidget(chartlabel)

            self.tiedup_canvas_container = QWidget()
            self.tiedup_canvas_container_layout = QVBoxLayout(self.tiedup_canvas_container)
            self.tiedup_canvas = None  
            self.tiedup_canvas_container_layout.addWidget(QLabel("Click 'Generate Forecast' to display chart"))
            layout.addWidget(self.tiedup_canvas_container)

            widget.setLayout(layout)
            logger.info("Tied-up-capital forecast tab created successfully")
            return widget

        except Exception as e:
            error_msg = f"Error creating Tied-up-capital tab: {str(e)}\n{tb.format_exc()}"
            logger.error(error_msg)
            print(error_msg)
            raise

    def create_supplier_deep_dive_tab(self):
        try:
            logger.info("Creating Supplier Deep Dive tab...")
            widget = QWidget()
            layout = QVBoxLayout()

            title = QLabel("Supplier Deep Dive")
            title.setFont(QFont("Arial", 16, QFont.Bold))
            title.setAlignment(Qt.AlignCenter)
            layout.addWidget(title)

            buttonlayout = QHBoxLayout()
            loadbtn = QPushButton("Load Supplier Data")
            loadbtn.clicked.connect(self.load_supplier_data)
            loadbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 10px 20px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
            buttonlayout.addWidget(loadbtn)
            buttonlayout.addStretch()
            layout.addLayout(buttonlayout)

            filterlayout = QHBoxLayout()
            filterlabel = QLabel("Supplier Filter:")
            filterlabel.setFont(QFont("Arial", 12, QFont.Bold))
            filterlayout.addWidget(filterlabel)

            self.supplier_filter = self.create_multiselect_dropdown("Select Suppliers...", "Supplier")
            filterlayout.addWidget(self.supplier_filter)

            filterlayout.addStretch()
            layout.addLayout(filterlayout)

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
        try:
            logger.info("Creating Strategy Analysis tab...")
            widget = QWidget()
            outer_layout = QVBoxLayout()

            title = QLabel("Strategy Analysis")
            title.setFont(QFont("Arial", 16, QFont.Bold))
            title.setAlignment(Qt.AlignCenter)
            outer_layout.addWidget(title)

            main_split = QHBoxLayout()
            main_split.setSpacing(10)

            left_widget = QWidget()
            left_layout = QVBoxLayout(left_widget)
            left_layout.setContentsMargins(0, 0, 0, 0)
            left_layout.setSpacing(6)

            buttonlayout = QHBoxLayout()
            loadbtn = QPushButton("Load Analysis Data")
            loadbtn.clicked.connect(self.load_strategy_data)
            loadbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 10px 20px; border: none; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
            buttonlayout.addWidget(loadbtn)
            buttonlayout.addStretch()
            left_layout.addLayout(buttonlayout)

            filterlabel = QLabel("Filters:")
            filterlabel.setFont(QFont("Arial", 12, QFont.Bold))
            left_layout.addWidget(filterlabel)

            filtergrid = QHBoxLayout()
            filtergrid.setSpacing(10)

            self.strategy_part_filter = self.create_multiselect_dropdown("Select Parts...", "Part")
            filtergrid.addWidget(self.strategy_part_filter)

            self.strategy_supplier_filter = self.create_multiselect_dropdown("Select Suppliers...", "Supplier")
            filtergrid.addWidget(self.strategy_supplier_filter)

            self.strategy_region_filter = self.create_multiselect_dropdown("Select Regions...", "Region")
            filtergrid.addWidget(self.strategy_region_filter)

            self.strategy_country_filter = self.create_multiselect_dropdown("Select Countries...", "Country")
            filtergrid.addWidget(self.strategy_country_filter)

            self.strategy_scc_filter = self.create_multiselect_dropdown("Select SCC...", "SCC")
            filtergrid.addWidget(self.strategy_scc_filter)

            clearfilterbtn = QPushButton("Clear Filters")
            clearfilterbtn.clicked.connect(self.clear_strategy_filters)
            clearfilterbtn.setStyleSheet("""QPushButton {background-color: #E97132; color: white; padding: 8px 16px; border: none; border-radius: 5px;} QPushButton:hover {background-color: #da190b;}""")
            filtergrid.addWidget(clearfilterbtn)
            filtergrid.setAlignment(clearfilterbtn, Qt.AlignVCenter)
            filtergrid.addStretch()
            left_layout.addLayout(filtergrid)

            tablelabel = QLabel("Filtered Parts Data")
            tablelabel.setFont(QFont("Arial", 12, QFont.Bold))
            left_layout.addWidget(tablelabel)

            self.strategy_table = QTableWidget()
            self.strategy_table.setSortingEnabled(True)
            self.strategy_table.setStyleSheet("""QTableWidget {gridline-color: #d0d0d0; background-color: white;} QTableWidget::item {padding: 6px; border: 1px solid #d0d0d0;} QHeaderView::section {background-color: #f0f0f0; padding: 8px; border: 1px solid #d0d0d0; font-weight: bold;}""")
            left_layout.addWidget(self.strategy_table)

            main_split.addWidget(left_widget, 2) 

            right_widget = QWidget()
            right_layout = QVBoxLayout(right_widget)
            right_layout.setContentsMargins(0, 0, 0, 0)
            right_layout.setSpacing(4)

            plotlabel = QLabel("3D Scatter Plot: SAFETY vs STOCK vs Price")
            plotlabel.setFont(QFont("Arial", 12, QFont.Bold))
            right_layout.addWidget(plotlabel)

            self.strategy_canvas_container = QWidget()
            self.strategy_canvas_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self.strategy_canvas_container_layout = QVBoxLayout(self.strategy_canvas_container)
            self.strategy_canvas = None 
            self.strategy_canvas_container_layout.addWidget(QLabel("Click 'Load Analysis Data' to display 3D plot"))
            right_layout.addWidget(self.strategy_canvas_container)  

            main_split.addWidget(right_widget, 3) 

            outer_layout.addLayout(main_split)
            widget.setLayout(outer_layout)
            logger.info("Strategy Analysis tab created successfully")
            return widget

        except Exception as e:
            error_msg = f"Error creating Strategy Analysis tab: {str(e)}\n{tb.format_exc()}"
            logger.error(error_msg)
            print(error_msg)
            raise

    def create_multiselect_dropdown(self, placeholdertext, filtertype="Item"):
        try:
            logger.info(f"Creating multiselect dropdown for {filtertype}")
            return _get_simple_filter_class()(placeholdertext, filtertype)
        except Exception as e:
            error_msg = f"Error creating multiselect dropdown for {filtertype}: {str(e)}\n{tb.format_exc()}"
            logger.error(error_msg)
            print(error_msg)
            raise

    def load_required_data(self):
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
        success, message = self.load_required_data()
        if not success:
            QMessageBox.warning(self, "Missing Data", message)
            return

        if self._mc_thread is not None and self._mc_thread.isRunning():
            QMessageBox.information(self, "In Progress", "Simulation is already running.")
            return

        try:
            progress = QProgressDialog("Computing inventory values...", "Cancel", 0, 100, self)
            progress.setWindowTitle("Generating Forecast")
            progress.setMinimumDuration(0)
            progress.setWindowModality(Qt.WindowModal)
            progress.setValue(5)
            progress.show()
            QApplication.processEvents()

            top_parts = self.compute_top_parts_by_value()
            top_suppliers = self.compute_top_suppliers_by_value()
            regions_value = self.compute_regions_by_value()
            self.display_top_parts_table(top_parts)
            self.display_top_suppliers_table(top_suppliers)
            self.display_regions_table(regions_value)
            progress.setValue(25)
            QApplication.processEvents()
            
            self._mc_progress_dialog = progress
            MCSimThread = _get_mc_sim_thread_class()
            self._mc_thread = MCSimThread(self.import_manager, days=90, n_sims=50)
            self._mc_thread.progress.connect(self._on_mc_progress)
            self._mc_thread.finished.connect(self._on_mc_finished)
            self._mc_thread.error.connect(self._on_mc_error)
            progress.canceled.connect(self._on_mc_cancel)
            self._mc_thread.start()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Forecast generation failed: {str(e)}")

    def _on_mc_progress(self, value: int, message: str):
        if self._mc_progress_dialog is not None:
            self._mc_progress_dialog.setValue(value)
            self._mc_progress_dialog.setLabelText(message)

    def _on_mc_finished(self, forecast_result: dict):
        if self._mc_progress_dialog is not None:
            self._mc_progress_dialog.close()
            self._mc_progress_dialog = None
        self.display_mc_chart(forecast_result=forecast_result)
        if forecast_result:
            QMessageBox.information(self, "Success", "Forecast generated successfully")
        else:
            QMessageBox.information(
                self, "Complete",
                "Summary tables generated.\n"
                "Monte Carlo was unavailable or cancelled — chart shows historical data only."
            )

    def _on_mc_error(self, error_msg: str):
        if self._mc_progress_dialog is not None:
            self._mc_progress_dialog.close()
            self._mc_progress_dialog = None
        logger.warning(f"MC simulation failed: {error_msg}")
        self.display_mc_chart(forecast_result=None)
        QMessageBox.warning(
            self, "Simulation Warning",
            f"Monte Carlo simulation failed:\n{error_msg}\n\n"
            "Chart shows historical data only."
        )

    def _on_mc_cancel(self):
        if self._mc_thread is not None and self._mc_thread.isRunning():
            logger.info("MC simulation cancel requested by user")
            self._mc_thread.cancel()

    def compute_top_parts_by_value(self):
        try:
            inv = self._current_inventory.copy()
            if 'PRICE' not in inv.columns:
                print(f"Available columns: {list(inv.columns)}")
                return pd.DataFrame()

            qty_cols = [c for c in inv.columns if 'INVENTORY' in c or 'QOH' in c or 'QUANTITY' in c]
            if not qty_cols:
                print(f"No quantity column found. Available: {list(inv.columns)}")
                return pd.DataFrame()

            qty_col = qty_cols[0]

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
        try:
            inv = self._current_inventory.copy()
            if 'PRICE' not in inv.columns:
                return pd.DataFrame()

            supp_col = next((c for c in inv.columns if 'SUPP' in c and 'NAME' in c), None)
            if not supp_col:
                return pd.DataFrame()

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
        try:
            inv = self._current_inventory.copy()

            country_col = next((c for c in inv.columns if 'COUNTRY' in c or 'SHIP_COUNTRY' in c), None)
            if not country_col:
                print(f"No country column found. Available: {list(inv.columns)}")
                return pd.DataFrame()

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

    def _fit_table_to_content(self, table):
        table.resizeColumnsToContents()
        header_h = table.horizontalHeader().height()
        rows_h = sum(table.rowHeight(i) for i in range(table.rowCount()))
        table.setFixedHeight(header_h + rows_h + 4)
        total_col_w = sum(table.columnWidth(i) for i in range(table.columnCount()))
        vheader_w = table.verticalHeader().sizeHint().width()
        table.setFixedWidth(total_col_w + vheader_w + 4)

    def display_top_parts_table(self, df):
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

        self._fit_table_to_content(self.tiedup_parts_table)

    def display_top_suppliers_table(self, df):
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

        self._fit_table_to_content(self.tiedup_suppliers_table)

    def display_regions_table(self, df):
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

        self._fit_table_to_content(self.tiedup_regions_table)

    def display_mc_chart(self, forecast_result=None):
        try:
            import matplotlib.pyplot as plt
            from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
            from matplotlib.figure import Figure
            if self.tiedup_canvas is None:
                self.tiedup_canvas = FigureCanvas(Figure(figsize=(12, 4), dpi=100))
                self.tiedup_canvas_container_layout.takeAt(0).widget().deleteLater()
                self.tiedup_canvas_container_layout.addWidget(self.tiedup_canvas)

            self.tiedup_canvas.figure.clear()
            fig = self.tiedup_canvas.figure
            ax = fig.add_subplot(111)
            has_data = False

            try:
                archive_path = getsharednetworkpath() / "archive" / "Inventory_Archive.csv"
                if archive_path.exists():
                    arc = (pd.read_csv(archive_path, parse_dates=["date"])
                           .dropna(subset=["date", "total_value"])
                           .sort_values("date"))
                    if not arc.empty:
                        ax.plot(arc["date"], arc["total_value"] / 1e6,
                                color="darkgreen", lw=2, label="Historical")
                        has_data = True
                        logger.info(f"display_mc_chart: plotted {len(arc)} historical days")
            except Exception as e:
                logger.warning(f"Could not load historical archive: {e}")

            if forecast_result:
                try:
                    fc_vals = np.array(forecast_result["plant_trajectory"][1:])
                    fc_dates = pd.date_range(
                        pd.Timestamp.today() + pd.Timedelta(days=1),
                        periods=len(fc_vals), freq="D"
                    )
                    ax.plot(fc_dates, fc_vals / 1e6,
                            color="royalblue", lw=2, label="MC Forecast (90-day)")
                    has_data = True
                    logger.info(f"display_mc_chart: plotted {len(fc_vals)}-day MC forecast")
                except Exception as e:
                    logger.warning(f"Could not plot MC forecast: {e}")

            ax.set_title("Tied-Up Capital — Historical & Forecast",
                         fontsize=12, fontweight="bold")
            if has_data:
                ax.axvline(pd.Timestamp.today(), color="red", ls="--", lw=1.5, label="Today")
                ax.set_ylabel("Value ($M)", fontsize=10)
                ax.yaxis.set_major_formatter(
                    plt.FuncFormatter(lambda v, _: f"${v:.1f}M"))
                ax.tick_params(axis="x", rotation=25)
                ax.legend(fontsize=9)
                ax.grid(True, alpha=0.3)
            else:
                ax.text(0.5, 0.5,
                        "No historical archive found.\n"
                        "Ensure the shared network archive CSV exists\n"
                        "at: archive/Inventory_Archive.csv",
                        ha="center", va="center", transform=ax.transAxes,
                        fontsize=11, color="gray")

            fig.tight_layout()
            self.tiedup_canvas.draw()
            logger.info("display_mc_chart: draw complete")

        except Exception as e:
            logger.error(f"Error displaying MC chart: {e}\n{tb.format_exc()}")

    def export_tiedup_tables(self):
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
        success, message = self.load_required_data()
        if not success:
            QMessageBox.warning(self, "Missing Data", message)
            return

        try:
            supp_col = next((c for c in self._master_data.columns if 'SUPP' in c and 'NAME' in c), None)
            if not supp_col:
                raise ValueError("Could not find supplier column in master data")

            suppliers = self._master_data[supp_col].dropna().unique()
            self.supplier_filter.additems(suppliers)

            self.supplier_filter.selectionChanged.connect(self.apply_supplier_filter)

            QMessageBox.information(self, "Success", f"Loaded {len(suppliers)} suppliers")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load suppliers: {str(e)}")
            import traceback
            traceback.print_exc()

    def load_strategy_data(self):
        success, message = self.load_required_data()
        if not success:
            QMessageBox.warning(self, "Missing Data", message)
            return

        try:
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

            self._master_data['Region'] = self._master_data[country_col].apply(self.determine_region)
            regions = self._master_data['Region'].dropna().unique()

            self.strategy_part_filter.additems(parts)
            self.strategy_supplier_filter.additems(suppliers)
            self.strategy_region_filter.additems(regions)
            self.strategy_country_filter.additems(countries)
            self.strategy_scc_filter.additems(sccs)

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

            self.display_3d_scatterplot(self._master_data)
            self.display_strategy_table(self._master_data)

            QMessageBox.information(self, "Success", "Loaded strategy analysis data")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load strategy data: {str(e)}")
            import traceback
            traceback.print_exc()

    def apply_supplier_filter(self):
        success, message = self.load_required_data()
        if not success:
            return

        try:
            supp_col = next((c for c in self._master_data.columns if 'SUPP' in c and 'NAME' in c), None)
            if not supp_col:
                return

            selected_suppliers = self.supplier_filter.getselecteditems()
            if not selected_suppliers:
                return

            master = self._master_data.copy()
            filtered = master[master[supp_col].isin(selected_suppliers)]

            if filtered.empty:
                self.supplier_parts_table.setRowCount(0)
                return
            self.display_supplier_parts_table(filtered)

        except Exception as e:
            print(f"Filter error: {e}")
            import traceback
            traceback.print_exc()

    def display_supplier_parts_table(self, df):
        if df.empty:
            self.supplier_parts_table.setRowCount(0)
            return

        try:
            part_col = next((c for c in df.columns if c in ['PART', 'PART_NO', 'PART_NUMBER']), None)
            desc_col = next((c for c in df.columns if 'DESC' in c), None)
            supp_col = next((c for c in df.columns if 'SUPP' in c and 'NAME' in c), None)
            price_col = next((c for c in df.columns if 'PRICE' in c), None)

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
        success, message = self.load_required_data()
        if not success:
            return

        try:
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

            filtered = self._master_data.copy()

            if selected_parts:
                filtered = filtered[filtered[part_col].isin(selected_parts)]
            if selected_suppliers:
                filtered = filtered[filtered[supp_col].isin(selected_suppliers)]
            if selected_regions:
                filtered['Region'] = filtered[country_col].apply(self.determine_region)
                filtered = filtered[filtered['Region'].isin(selected_regions)]
            if selected_countries:
                filtered = filtered[filtered[country_col].isin(selected_countries)]
            if selected_sccs:
                filtered = filtered[filtered[scc_col].isin(selected_sccs)]

            if filtered.empty:
                self.strategy_table.setRowCount(0)
                return

            self.display_3d_scatterplot(filtered)
            self.display_strategy_table(filtered)

        except Exception as e:
            print(f"Filter error: {e}")
            import traceback
            traceback.print_exc()

    def display_3d_scatterplot(self, df):
        if df.empty:
            logger.warning("display_3d_scatterplot: dataframe is empty, skipping")
            return

        try:
            from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
            from matplotlib.figure import Figure
            from mpl_toolkits.mplot3d import Axes3D

            if self.strategy_canvas is None:
                self.strategy_canvas = FigureCanvas(Figure(figsize=(6, 8), dpi=100))
                self.strategy_canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                self.strategy_canvas_container_layout.takeAt(0).widget().deleteLater()
                self.strategy_canvas_container_layout.addWidget(self.strategy_canvas)

            logger.info(f"display_3d_scatterplot: df shape={df.shape}, columns={list(df.columns)}")

            safety_col = next((c for c in df.columns if 'SAFETY' in c.upper()), None)
            stock_col = next((c for c in df.columns if c.upper() in ('STOCK', 'STOCK_QTY', 'STOCK_VALUE') or ('STOCK' in c.upper() and 'SAFETY' not in c.upper())), None)
            price_col = next((c for c in df.columns if 'PRICE' in c.upper()), None)

            logger.info(f"display_3d_scatterplot: safety_col={safety_col}, stock_col={stock_col}, price_col={price_col}")

            if not all([safety_col, stock_col, price_col]):
                missing = [n for n, c in [('SAFETY', safety_col), ('STOCK', stock_col), ('PRICE', price_col)] if c is None]
                logger.error(f"Cannot plot 3D: Missing columns {missing}. Available: {list(df.columns)}")
                self.strategy_canvas.figure.clear()
                ax = self.strategy_canvas.figure.add_subplot(111)
                ax.text(0.5, 0.5, f'Missing columns: {missing}', ha='center', va='center', transform=ax.transAxes)
                self.strategy_canvas.draw()
                return

            self.strategy_canvas.figure.clear()
            ax = self.strategy_canvas.figure.add_subplot(111, projection='3d')

            safety = pd.to_numeric(df[safety_col], errors='coerce').fillna(0).values
            stock = pd.to_numeric(df[stock_col], errors='coerce').fillna(0).values
            price = pd.to_numeric(df[price_col], errors='coerce').fillna(0).values

            logger.info(f"display_3d_scatterplot: plotting {len(safety)} points, safety range [{safety.min():.2f}, {safety.max():.2f}]")

            scatter = ax.scatter(safety, stock, price, c=price, cmap='viridis', marker='o', s=50, alpha=0.6)

            ax.set_xlabel('SAFETY', fontsize=9)
            ax.set_ylabel('STOCK', fontsize=9)
            ax.set_zlabel('Price', fontsize=9)
            ax.set_title('3D Strategy Analysis', fontsize=11, fontweight='bold')

            self.strategy_canvas.figure.colorbar(scatter, ax=ax, shrink=0.6, label='Price')
            self.strategy_canvas.draw()
            logger.info("display_3d_scatterplot: draw complete")

        except Exception as e:
            logger.error(f"Error displaying 3D plot: {e}\n{tb.format_exc()}")

    def display_strategy_table(self, df):
        if df.empty:
            self.strategy_table.setRowCount(0)
            return

        try:
            part_col = next((c for c in df.columns if c in ['PART', 'PART_NO', 'PART_NUMBER']), None)
            desc_col = next((c for c in df.columns if 'DESC' in c), None)
            supp_col = next((c for c in df.columns if 'SUPP' in c and 'NAME' in c), None)
            safety_col = next((c for c in df.columns if 'SAFETY' in c.upper()), None)
            stock_col = next((c for c in df.columns if c.upper() in ('STOCK', 'STOCK_QTY', 'STOCK_VALUE') or ('STOCK' in c.upper() and 'SAFETY' not in c.upper())), None)

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
        self.apply_strategy_filters()

    def closeEvent(self, event):
        reply = QMessageBox.question(self, "Exit", "Are you sure you want to exit?",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            if self._mc_thread is not None and self._mc_thread.isRunning():
                self._mc_thread.cancel()
                self._mc_thread.wait(5000)          
                if self._mc_thread.isRunning():
                    self._mc_thread.terminate()    
                    self._mc_thread.wait()
            self.hide()
            QApplication.processEvents()
            event.accept()
        else:
            event.ignore()


