import pandas as pd
from PyQt5.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QLineEdit, QTextEdit, QCheckBox,
    QTableWidget, QTableWidgetItem, QMessageBox, QSplitter, QFrame,
    QScrollArea, QApplication
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor


class _TestWorker(QThread):
    finished = pyqtSignal(bool, str, object)

    def __init__(self, import_manager, connection_string, query):
        super().__init__()
        self.import_manager = import_manager
        self.connection_string = connection_string
        self.query = query

    def run(self):
        ok, msg, preview = self.import_manager.testodbcconnection(
            self.connection_string, self.query
        )
        self.finished.emit(ok, msg, preview)


class ODBCConfigDialog(QDialog):
    def __init__(self, import_manager, parent=None):
        super().__init__(parent)
        self.import_manager = import_manager
        self.setWindowTitle("ODBC Connection Settings")
        self.resize(900, 620)
        self._pending = {}
        self._test_worker = None
        self._setupui()
        self._loadcategories()

    # ── UI build ──────────────────────────────────────────────────────────────

    def _setupui(self):
        root = QVBoxLayout()
        root.setContentsMargins(12, 12, 12, 12)

        title = QLabel("ODBC Connection Settings")
        title.setFont(QFont("Arial", 15, QFont.Bold))
        root.addWidget(title)

        subtitle = QLabel(
            "Configure a live ODBC connection for any data category. "
            "When enabled, the app queries the database directly and falls "
            "back to the imported file if the connection fails."
        )
        subtitle.setWordWrap(True)
        subtitle.setStyleSheet("color: #555; font-size: 11px;")
        root.addWidget(subtitle)

        splitter = QSplitter(Qt.Horizontal)

        # ── Left: category list ───────────────────────────────────────────────
        leftpanel = QWidget()
        leftpanel.setFixedWidth(230)
        leftlayout = QVBoxLayout()
        leftlayout.setContentsMargins(0, 0, 6, 0)

        listlabel = QLabel("Data Categories")
        listlabel.setFont(QFont("Arial", 10, QFont.Bold))
        leftlayout.addWidget(listlabel)

        self.categorylist = QListWidget()
        self.categorylist.currentRowChanged.connect(self._oncategorychanged)
        self.categorylist.setStyleSheet("""
            QListWidget { border: 1px solid #ccc; font-size: 11px; }
            QListWidget::item { padding: 6px 8px; }
            QListWidget::item:selected { background: #156082; color: white; }
        """)
        leftlayout.addWidget(self.categorylist)

        leftpanel.setLayout(leftlayout)
        splitter.addWidget(leftpanel)

        # ── Right: connection form ────────────────────────────────────────────
        rightscroll = QScrollArea()
        rightscroll.setWidgetResizable(True)
        rightscroll.setFrameShape(QFrame.NoFrame)

        rightpanel = QWidget()
        self.formlayout = QVBoxLayout()
        self.formlayout.setContentsMargins(10, 0, 0, 0)
        self.formlayout.setSpacing(10)

        # Category name header
        self.categorytitle = QLabel("Select a category")
        self.categorytitle.setFont(QFont("Arial", 13, QFont.Bold))
        self.formlayout.addWidget(self.categorytitle)

        self.categorydesc = QLabel("")
        self.categorydesc.setStyleSheet("color: #666; font-size: 10px;")
        self.formlayout.addWidget(self.categorydesc)

        # Enable checkbox
        self.enablecheck = QCheckBox("Enable ODBC for this category")
        self.enablecheck.setFont(QFont("Arial", 10))
        self.enablecheck.toggled.connect(self._onformchanged)
        self.formlayout.addWidget(self.enablecheck)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #ddd;")
        self.formlayout.addWidget(sep)

        # Connection string
        connlabel = QLabel("Connection String")
        connlabel.setFont(QFont("Arial", 10, QFont.Bold))
        self.formlayout.addWidget(connlabel)

        connhint = QLabel(
            "DSN format:  DSN=YourDSNName;\n"
            "Driver format:  DRIVER={ODBC Driver 17 for SQL Server};"
            "SERVER=host;DATABASE=db;Trusted_Connection=yes;"
        )
        connhint.setStyleSheet(
            "color: #777; font-size: 9px; font-family: Consolas, monospace;"
        )
        connhint.setWordWrap(True)
        self.formlayout.addWidget(connhint)

        self.connstringinput = QLineEdit()
        self.connstringinput.setPlaceholderText(
            "e.g.  DSN=MyDSN;  or  DRIVER={SQL Server};SERVER=...;DATABASE=...;"
        )
        self.connstringinput.setStyleSheet(
            "QLineEdit { border: 2px solid #ccc; border-radius: 3px; "
            "padding: 6px; font-family: Consolas, monospace; font-size: 10px; }"
            "QLineEdit:focus { border-color: #156082; }"
        )
        self.connstringinput.textChanged.connect(self._onformchanged)
        self.formlayout.addWidget(self.connstringinput)

        # SQL Query
        querylabel = QLabel("SQL Query")
        querylabel.setFont(QFont("Arial", 10, QFont.Bold))
        self.formlayout.addWidget(querylabel)

        queryhint = QLabel(
            "Must return the required columns for this category "
            "(shown below). Column names must match exactly."
        )
        queryhint.setStyleSheet("color: #777; font-size: 9px;")
        queryhint.setWordWrap(True)
        self.formlayout.addWidget(queryhint)

        self.queryinput = QTextEdit()
        self.queryinput.setPlaceholderText("SELECT col1, col2, ... FROM your_table WHERE ...")
        self.queryinput.setFixedHeight(100)
        self.queryinput.setStyleSheet(
            "QTextEdit { border: 2px solid #ccc; border-radius: 3px; "
            "padding: 6px; font-family: Consolas, monospace; font-size: 10px; }"
            "QTextEdit:focus { border-color: #156082; }"
        )
        self.queryinput.textChanged.connect(self._onformchanged)
        self.formlayout.addWidget(self.queryinput)

        # Required columns hint
        self.reqcolslabel = QLabel("")
        self.reqcolslabel.setStyleSheet(
            "color: #444; font-size: 9px; font-family: Consolas, monospace; "
            "background: #f8f8f8; padding: 6px; border: 1px solid #ddd; border-radius: 3px;"
        )
        self.reqcolslabel.setWordWrap(True)
        self.formlayout.addWidget(self.reqcolslabel)

        # Test button + status
        testrow = QHBoxLayout()
        self.testbtn = QPushButton("Test Connection")
        self.testbtn.setFixedWidth(150)
        self.testbtn.clicked.connect(self._testconnection)
        self.testbtn.setStyleSheet(
            "QPushButton { background: #156082; color: white; padding: 7px 14px; "
            "border-radius: 4px; font-weight: bold; }"
            "QPushButton:hover { background: #1a7da8; }"
            "QPushButton:disabled { background: #aaa; }"
        )
        testrow.addWidget(self.testbtn)

        self.teststatuslabel = QLabel("")
        self.teststatuslabel.setWordWrap(True)
        self.teststatuslabel.setStyleSheet("font-size: 10px;")
        testrow.addWidget(self.teststatuslabel, 1)
        self.formlayout.addLayout(testrow)

        # Preview table
        self.previewtable = QTableWidget()
        self.previewtable.setEditTriggers(QTableWidget.NoEditTriggers)
        self.previewtable.setFixedHeight(130)
        self.previewtable.setVisible(False)
        self.previewtable.setStyleSheet(
            "QTableWidget { font-size: 9px; } "
            "QHeaderView::section { background: #f0f0f0; font-size: 9px; padding: 3px; }"
        )
        self.formlayout.addWidget(self.previewtable)

        self.formlayout.addStretch()
        rightpanel.setLayout(self.formlayout)
        rightscroll.setWidget(rightpanel)
        splitter.addWidget(rightscroll)
        splitter.setStretchFactor(1, 1)

        root.addWidget(splitter)

        # ── Bottom buttons ────────────────────────────────────────────────────
        btnsep = QFrame()
        btnsep.setFrameShape(QFrame.HLine)
        btnsep.setStyleSheet("color: #ddd;")
        root.addWidget(btnsep)

        btnrow = QHBoxLayout()

        self.savebtn = QPushButton("Save")
        self.savebtn.setFixedWidth(100)
        self.savebtn.clicked.connect(self._savecurrent)
        self.savebtn.setEnabled(False)
        self.savebtn.setStyleSheet(
            "QPushButton { background: #4CAF50; color: white; padding: 7px 14px; "
            "border-radius: 4px; font-weight: bold; }"
            "QPushButton:hover { background: #45a049; }"
            "QPushButton:disabled { background: #aaa; }"
        )
        btnrow.addWidget(self.savebtn)

        removebtn = QPushButton("Remove ODBC for this Category")
        removebtn.setFixedWidth(220)
        removebtn.clicked.connect(self._removecurrent)
        removebtn.setStyleSheet(
            "QPushButton { background: #E97132; color: white; padding: 7px 14px; "
            "border-radius: 4px; }"
            "QPushButton:hover { background: #c0552a; }"
        )
        btnrow.addWidget(removebtn)

        btnrow.addStretch()

        closebtn = QPushButton("Close")
        closebtn.setFixedWidth(90)
        closebtn.clicked.connect(self.accept)
        closebtn.setStyleSheet(
            "QPushButton { padding: 7px 14px; border-radius: 4px; border: 1px solid #ccc; }"
        )
        btnrow.addWidget(closebtn)

        root.addLayout(btnrow)
        self.setLayout(root)
        self._setformenabled(False)

    # ── Population ────────────────────────────────────────────────────────────

    def _loadcategories(self):
        self.categorylist.blockSignals(True)
        self.categorylist.clear()
        categories = self.import_manager.getimportcategories()
        existing = self.import_manager.getodbcconfig()
        for key, info in categories.items():
            item = QListWidgetItem(info["name"])
            item.setData(Qt.UserRole, key)
            if key in existing and existing[key].get("enabled"):
                item.setForeground(QColor("#156082"))
                item.setText(f"● {info['name']}")
            self.categorylist.addItem(item)
        self.categorylist.blockSignals(False)

    def _populateform(self, category_key: str):
        cats = self.import_manager.getimportcategories()
        info = cats.get(category_key, {})

        self.categorytitle.setText(info.get("name", category_key))
        self.categorydesc.setText(info.get("description", ""))

        reqcols = info.get("requiredcolumns", [])
        self.reqcolslabel.setText("Required columns:  " + ",  ".join(reqcols))

        cfg = self._pending.get(
            category_key,
            self.import_manager.getodbcconfig().get(category_key, {})
        )

        self.enablecheck.blockSignals(True)
        self.connstringinput.blockSignals(True)
        self.queryinput.blockSignals(True)

        self.enablecheck.setChecked(cfg.get("enabled", False))
        self.connstringinput.setText(cfg.get("connection_string", ""))
        self.queryinput.setPlainText(cfg.get("query", ""))

        self.enablecheck.blockSignals(False)
        self.connstringinput.blockSignals(False)
        self.queryinput.blockSignals(False)

        self.teststatuslabel.setText("")
        self.previewtable.setVisible(False)
        self.savebtn.setEnabled(False)
        self._setformenabled(True)

    # ── Slots ─────────────────────────────────────────────────────────────────

    def _oncategorychanged(self, row):
        if row < 0:
            return
        key = self.categorylist.item(row).data(Qt.UserRole)
        self._currentkey = key
        self._populateform(key)

    def _onformchanged(self):
        key = getattr(self, "_currentkey", None)
        if key is None:
            return
        self._pending[key] = {
            "enabled": self.enablecheck.isChecked(),
            "connection_string": self.connstringinput.text().strip(),
            "query": self.queryinput.toPlainText().strip(),
        }
        self.savebtn.setEnabled(True)
        self.teststatuslabel.setText("")
        self.previewtable.setVisible(False)

    def _testconnection(self):
        conn = self.connstringinput.text().strip()
        query = self.queryinput.toPlainText().strip()
        if not conn or not query:
            QMessageBox.warning(self, "Missing Input",
                                "Enter a connection string and a SQL query first.")
            return

        self.testbtn.setEnabled(False)
        self.testbtn.setText("Testing…")
        self.teststatuslabel.setText("Connecting…")
        self.teststatuslabel.setStyleSheet("color: #555; font-size: 10px;")
        QApplication.processEvents()

        self._test_worker = _TestWorker(self.import_manager, conn, query)
        self._test_worker.finished.connect(self._ontestfinished)
        self._test_worker.start()

    def _ontestfinished(self, ok: bool, msg: str, preview):
        self.testbtn.setEnabled(True)
        self.testbtn.setText("Test Connection")

        if ok:
            self.teststatuslabel.setText(f"✓  {msg}")
            self.teststatuslabel.setStyleSheet("color: #2e7d32; font-size: 10px;")
            self._showpreview(preview)
        else:
            self.teststatuslabel.setText(f"✗  {msg}")
            self.teststatuslabel.setStyleSheet("color: #c62828; font-size: 10px;")
            self.previewtable.setVisible(False)

    def _showpreview(self, df: pd.DataFrame):
        if df is None or df.empty:
            return
        self.previewtable.setRowCount(len(df))
        self.previewtable.setColumnCount(len(df.columns))
        self.previewtable.setHorizontalHeaderLabels(df.columns.tolist())
        for r, row in df.iterrows():
            for c, val in enumerate(row):
                self.previewtable.setItem(
                    list(df.index).index(r), c,
                    QTableWidgetItem(str(val) if pd.notna(val) else "")
                )
        self.previewtable.resizeColumnsToContents()
        self.previewtable.setVisible(True)

    def _savecurrent(self):
        key = getattr(self, "_currentkey", None)
        if key is None:
            return
        cfg = self._pending.get(key, {})
        conn = cfg.get("connection_string", "").strip()
        query = cfg.get("query", "").strip()
        if cfg.get("enabled") and (not conn or not query):
            QMessageBox.warning(self, "Incomplete",
                                "Enter a connection string and query before saving.")
            return
        self.import_manager.setodbcconfig(
            key, conn, query, enabled=cfg.get("enabled", False)
        )
        self._pending.pop(key, None)
        self.savebtn.setEnabled(False)
        self._loadcategories()
        QMessageBox.information(self, "Saved",
                                f"ODBC settings saved for:\n{self.import_manager.getimportcategories()[key]['name']}")

    def _removecurrent(self):
        key = getattr(self, "_currentkey", None)
        if key is None:
            return
        catname = self.import_manager.getimportcategories()[key]["name"]
        if QMessageBox.question(
            self, "Remove ODBC",
            f"Remove ODBC configuration for\n\"{catname}\"?\n\nThe app will fall back to imported files.",
            QMessageBox.Yes | QMessageBox.No
        ) == QMessageBox.Yes:
            self.import_manager.removeodbcconfig(key)
            self._pending.pop(key, None)
            self._loadcategories()
            self._populateform(key)

    def _setformenabled(self, enabled: bool):
        self.enablecheck.setEnabled(enabled)
        self.connstringinput.setEnabled(enabled)
        self.queryinput.setEnabled(enabled)
        self.testbtn.setEnabled(enabled)
