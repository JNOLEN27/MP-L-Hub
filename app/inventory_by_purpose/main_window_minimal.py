"""
Minimal test version to diagnose initialization crash
"""
import logging
from pathlib import Path
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QLabel, QTabWidget, QPushButton
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

# Setup logging
LOG_FILE = Path.home() / "InventoryByPurpose_Error.log"
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
)
logger = logging.getLogger(__name__)


class InventorybyPurposeWindow(QMainWindow):
    def __init__(self, userdata, parent=None):
        super().__init__(parent)
        try:
            logger.info("=== MINIMAL TEST: Initializing InventorybyPurposeWindow ===")
            self.userdata = userdata
            logger.info(f"User data loaded: {userdata.get('username', 'unknown')}")

            self.setWindowTitle("Inventory by Purpose Application")
            self.resize(1200, 800)

            logger.info("Setting up minimal UI...")
            self.setupui()
            logger.info("Minimal UI setup completed successfully")

        except Exception as e:
            error_msg = f"CRITICAL ERROR: {str(e)}"
            logger.error(error_msg, exc_info=True)
            print(error_msg)
            raise

    def setupui(self):
        """Setup minimal UI"""
        try:
            logger.info("Starting setupui...")
            centralwidget = QWidget()
            self.setCentralWidget(centralwidget)

            layout = QVBoxLayout()

            # Header
            title = QLabel("Inventory by Purpose [MINIMAL TEST]")
            title.setFont(QFont("Arial", 18, QFont.Bold))
            title.setAlignment(Qt.AlignCenter)
            layout.addWidget(title)

            # Create empty tabs (no filters, no tables, no canvas)
            logger.info("Creating tab widget...")
            tabs = QTabWidget()

            tab1 = QWidget()
            tab1_layout = QVBoxLayout()
            tab1_layout.addWidget(QLabel("Tab 1: Tied-up-capital forecast (EMPTY)"))
            tab1.setLayout(tab1_layout)
            tabs.addTab(tab1, "Tied-up-capital forecast")

            tab2 = QWidget()
            tab2_layout = QVBoxLayout()
            tab2_layout.addWidget(QLabel("Tab 2: Supplier Deep Dive (EMPTY)"))
            tab2.setLayout(tab2_layout)
            tabs.addTab(tab2, "Supplier Deep Dive")

            tab3 = QWidget()
            tab3_layout = QVBoxLayout()
            tab3_layout.addWidget(QLabel("Tab 3: Strategy Analysis (EMPTY)"))
            tab3.setLayout(tab3_layout)
            tabs.addTab(tab3, "Strategy Analysis")

            logger.info("Tabs created successfully")
            layout.addWidget(tabs)

            self.statusBar().showMessage(f"Logged in as: {self.userdata.get('username', 'unknown')}")
            centralwidget.setLayout(layout)

            logger.info("setupui completed successfully")

        except Exception as e:
            error_msg = f"ERROR in setupui: {str(e)}"
            logger.error(error_msg, exc_info=True)
            print(error_msg)
            raise
