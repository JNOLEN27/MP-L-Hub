from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGridLayout, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem,
    QProgressDialog,
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from pathlib import Path
import re

import pandas as pd

from app.data.import_manager import DataImportManager
from app.utils.config import COLORPRIMARY, COLORSUCCESS, ADMINUSERS, POWERUSERS, SHAREDNETWORKPATH

ARCHIVE_PATH = SHAREDNETWORKPATH / "archive" / "Inventory_Archive.csv"

class DataImportsWindow(QMainWindow):
    def __init__(self, userdata, parent=None):
        super().__init__(parent)
        self.userdata = userdata if isinstance(userdata, dict) else {'username': userdata}
        self.importmanager = DataImportManager()

        username = self.userdata['username']
        self._isadmin = username in ADMINUSERS
        self._allowedcategories = None if self._isadmin else set(POWERUSERS.get(username, []))

        self.setWindowTitle("Data Imports - Admin Panel" if self._isadmin else "Data Imports")
        self.resize(1000, 700)
        self.setupui()
        
    def setupui(self):
        centralwidget = QWidget()
        self.setCentralWidget(centralwidget)
        
        layout = QVBoxLayout()
        
        title = QLabel("Data Imports Manager")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet(f"color: {COLORPRIMARY};")
        layout.addWidget(title)
        
        subtitle = QLabel("Import data files for all applications")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #666; margin-bottom: 20px;")
        layout.addWidget(subtitle)
        
        importgrid = self.createimportbuttons()
        layout.addWidget(importgrid)
        
        historysection = self.createhistorysection()
        layout.addWidget(historysection)
        
        centralwidget.setLayout(layout)
        
        self.refreshhistory()
        
    def createimportbuttons(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        sectiontitle = QLabel("Import Data Files")
        sectiontitle.setFont(QFont("Arial", 14, QFont.Bold))
        layout.addWidget(sectiontitle)
        
        grid = QGridLayout()
        grid.setSpacing(15)
        
        importbuttons = [
            {"title": "Current Inventory Report",
             "description": "Current Inventory per Part",
             "category": "current_inventory_report",
             "filetypes": "Excel files (*.xlsx *.xls *.xlsm);;CSV files (*.csv)",
             "bulkarchive": True},
            {"title": "Master Data Reports",
             "description": "Part Information",
             "category": "master_data",
             "filetypes": "Excel files (*.xlsx *.xls *.xlsm);;CSV files (*.csv)"},
            {"title": "Part Requirement Split 1",
             "description": "Planned Consumption per Day per Part",
             "category": "part_requirement_split_1",
             "filetypes": "Excel files (*.xlsx *.xls *.xlsm);;CSV files (*.csv)"},
            {"title": "Part Requirement Split 2",
             "description": "Planned Consumption per Day per Part",
             "category": "part_requirement_split_2",
             "filetypes": "Excel files (*.xlsx *.xls *.xlsm);;CSV files (*.csv)"},
            {"title": "Part Requirement Split 3",
             "description": "Planned Consumption per Day per Part",
             "category": "part_requirement_split_3",
             "filetypes": "Excel files (*.xlsx *.xls *.xlsm);;CSV files (*.csv)"},
            {"title": "Goods to be Received Next 365 Days",
             "description": "Goods to be received by part per day for the next 365 days",
             "category": "goods_to_be_received",
             "filetypes": "Excel files (*.xlsx *.xls *.xlsm);;CSV files (*.csv)"},
            {"title": "Manual TTT",
             "description": "Delivery Days per Supplier",
             "category": "manual_TTT",
             "filetypes": "Excel files (*.xlsx *.xls *.xlsm);;CSV files (*.csv)"},
            {"title": "PVA Percentage",
             "description": "Plant Volume Adjustment distribution (SPA2 YTD) — used by Monte Carlo TUC simulation",
             "category": "pva_percentage",
             "filetypes": "CSV files (*.csv);;Excel files (*.xlsx *.xls *.xlsm)"}
            ]
        
        row = 0
        col = 0
        maxcols = 3
        
        for buttoninfo in importbuttons:
            btnwidget = self.createimportbutton(buttoninfo)
            grid.addWidget(btnwidget, row, col)
            
            col += 1
            if col >= maxcols:
                col = 0
                row += 1
                
        layout.addLayout(grid)
        widget.setLayout(layout)
        return widget
    
    def createimportbutton(self, buttoninfo):
        allowed = self._allowedcategories is None or buttoninfo["category"] in self._allowedcategories

        widget = QWidget()
        widget.setMinimumHeight(120)
        if allowed:
            widget.setStyleSheet(f"""QWidget {{background-color: {COLORPRIMARY}; border: 2px solid #ddd; border-radius: 8px; padding: 10px;}} QWidget:hover {{border-color: #f8f9fa; background-color: white;}}""")
        else:
            widget.setStyleSheet("""QWidget {background-color: #e0e0e0; border: 2px solid #ccc; border-radius: 8px; padding: 10px;}""")

        layout = QVBoxLayout()

        titlelayout = QHBoxLayout()
        titlelabel = QLabel(buttoninfo["title"])
        titlelabel.setFont(QFont("Arial", 12, QFont.Bold))
        if not allowed:
            titlelabel.setStyleSheet("color: #999;")
        titlelayout.addWidget(titlelabel)
        titlelayout.addStretch()
        layout.addLayout(titlelayout)

        descriptionlabel = QLabel(buttoninfo["description"])
        descriptionlabel.setStyleSheet("color: #666; font-size: 10px;")
        descriptionlabel.setWordWrap(True)
        layout.addWidget(descriptionlabel)

        importbtn = QPushButton("Import File" if allowed else "No Access")
        importbtn.setEnabled(allowed)
        if allowed:
            importbtn.setStyleSheet(f"""QPushButton {{background-color: {COLORPRIMARY}; color: white; border: none; padding: 8px; border-radius: 4px; font-weight: bold;}} QPushButton:hover {{background-color: #45a049;}}""")
            importbtn.clicked.connect(lambda checked, info=buttoninfo: self.handleimport(info))
        else:
            importbtn.setStyleSheet("""QPushButton {background-color: #bbb; color: #888; border: none; padding: 8px; border-radius: 4px; font-weight: bold;}""")
        layout.addWidget(importbtn)

        # Bulk archive button — CIR only, admins only
        if buttoninfo.get("bulkarchive") and self._isadmin:
            bulkbtn = QPushButton("Bulk Archive History")
            bulkbtn.setStyleSheet("""QPushButton {background-color: #5a7fa8; color: white; border: none; padding: 6px; border-radius: 4px; font-size: 10px;} QPushButton:hover {background-color: #3d6591;}""")
            bulkbtn.clicked.connect(lambda checked, info=buttoninfo: self.handlebulkarchive(info))
            layout.addWidget(bulkbtn)

        widget.setLayout(layout)
        return widget
    
    def handleimport(self, buttoninfo):
        category = buttoninfo["category"]
        filetypes = buttoninfo["filetypes"]
        title = buttoninfo["title"]
        filepath, _ = QFileDialog.getOpenFileName(self, f"Select {title} file to import","",filetypes)
        
        if not filepath:
            return
        
        filepath = Path(filepath)
        try:
            isvalid, errors, previewdata = self.importmanager.validatefile(filepath, category)
            
            if not isvalid:
                QMessageBox.warning(self, "Validation Failed", f"File validation failed:\n\n" + "\n".join(errors))
                return
            
            if errors:
                reply = QMessageBox.question(self, "Import with Warnings?", f"File has warnings but can be imported:\n\n" + "\n".join(errors) + f"\n\nPreview shows {len(previewdata)} rows.\n\nProceed with import?", QMessageBox.Yes | QMessageBox.No)
                if reply != QMessageBox.Yes:
                    return
            username = self.userdata['username']
            success, message = self.importmanager.importfile(filepath, category, username, f"Imported {title} via Data Imports app")
            if success:
                QMessageBox.information(self, "Import Successful", message)
                self.refreshhistory()
            else:
                QMessageBox.critical(self, "Import Failed", message)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Unexpected error during import: {str(e)}")
        
    def createhistorysection(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        title = QLabel("Recent Imports")
        title.setFont(QFont("Arial", 14, QFont.Bold))
        layout.addWidget(title)
        
        self.historytable = QTableWidget()
        self.historytable.setColumnCount(6)
        self.historytable.setHorizontalHeaderLabels(["Date/Time", "File Name", "Category", "Imported By", "Status", "Actions"])
        self.historytable.setEditTriggers(QTableWidget.NoEditTriggers)
        self.historytable.setSelectionBehavior(QTableWidget.SelectRows)
        
        layout.addWidget(self.historytable)
        
        refreshbtn = QPushButton("Refresh History")
        refreshbtn.clicked.connect(self.refreshhistory)
        layout.addWidget(refreshbtn)
        
        widget.setLayout(layout)
        return widget
    
    def refreshhistory(self):
        try:
            history = self.importmanager.getimporthistory()
            self.historytable.setRowCount(len(history))
            
            for row, item in enumerate(history):
                importtime = item['importedat'][:19].replace('T', ' ')
                self.historytable.setItem(row, 0, QTableWidgetItem(importtime))
                self.historytable.setItem(row, 1, QTableWidgetItem(item['originalfilename']))
                self.historytable.setItem(row, 2, QTableWidgetItem(item.get('categoryname', item['category'])))
                self.historytable.setItem(row, 3, QTableWidgetItem(item['importedby']))
                
                status = "Success" if item.get('validationpassed', False) else "Warning"
                self.historytable.setItem(row, 4, QTableWidgetItem(status))
                self.historytable.setItem(row, 5, QTableWidgetItem("View"))
                
            self.historytable.resizeColumnsToContents()
        
        except Exception as e:
            print(f"Error refreshing history: {e}")
        
    # ── Bulk CIR Archive ──────────────────────────────────────────────────────

    def handlebulkarchive(self, buttoninfo):
        """
        Let the user pick a folder of historical CIR files (CSV or XLSX),
        parse each file's date from its filename (pattern: 'CIR m.d.yy.*'),
        calculate total inventory value, and append every entry to the
        shared archive CSV used by the Monte Carlo TUC simulation.
        """
        folder = QFileDialog.getExistingDirectory(
            self, "Select folder containing historical CIR files", ""
        )
        if not folder:
            return

        folder = Path(folder)
        files  = self._findcirfiles(folder)

        if not files:
            QMessageBox.warning(
                self, "No Files Found",
                f"No CIR files (CSV or XLSX) matching 'CIR m.d.yy.*' were found in:\n{folder}"
            )
            return

        # Parse dates and drop files where the date couldn't be extracted
        entries = []
        unparsed = []
        for fp in files:
            d = self._parsecirdate(fp.name)
            if d:
                entries.append((d, fp))
            else:
                unparsed.append(fp.name)

        if not entries:
            QMessageBox.warning(
                self, "Date Parsing Failed",
                "Could not parse dates from any filename.\n"
                "Expected format: CIR m.d.yy.csv  or  CIR mm.dd.yyyy.xlsx"
            )
            return

        entries.sort(key=lambda x: x[0])

        msg = (
            f"Found {len(entries)} CIR file(s) to archive "
            f"({entries[0][0]} → {entries[-1][0]})."
        )
        if unparsed:
            msg += f"\n\n{len(unparsed)} file(s) skipped (unrecognised filename):\n"
            msg += "\n".join(f"  • {n}" for n in unparsed[:5])
            if len(unparsed) > 5:
                msg += f"\n  … and {len(unparsed) - 5} more"
        msg += "\n\nProceed?"

        if QMessageBox.question(self, "Confirm Bulk Archive", msg,
                                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return

        progress = QProgressDialog("Archiving CIR files…", "Cancel", 0, len(entries), self)
        progress.setWindowTitle("Bulk Archive Progress")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)

        results, errors = [], []

        for i, (report_date, fp) in enumerate(entries):
            if progress.wasCanceled():
                break
            progress.setLabelText(f"Processing {fp.name}…")

            try:
                df       = self._loadcirfile(fp)
                val, cnt = self._calculateinventoryvalue(df)
                results.append({"date": pd.Timestamp(report_date), "total_value": val, "part_count": cnt})
            except Exception as e:
                errors.append(f"{fp.name}: {e}")

            progress.setValue(i + 1)

        progress.close()

        if results:
            self._writetoarchive(results)

        summary = f"Archived {len(results)} file(s) successfully."
        if errors:
            summary += f"\n\n{len(errors)} file(s) failed:\n" + "\n".join(f"  • {e}" for e in errors[:10])
        QMessageBox.information(self, "Bulk Archive Complete", summary)
        self.refreshhistory()

    # ── Helpers ───────────────────────────────────────────────────────────────

    _CIR_PATTERN = re.compile(
        r"^CIR\s+(\d{1,2})\.(\d{1,2})\.(\d{2,4})\.(csv|xlsx|xls|xlsm)$",
        re.IGNORECASE,
    )

    def _findcirfiles(self, folder: Path):
        found = []
        for fp in folder.iterdir():
            if fp.is_file() and self._CIR_PATTERN.match(fp.name):
                found.append(fp)
        return sorted(found)

    def _parsecirdate(self, filename: str):
        m = self._CIR_PATTERN.match(filename)
        if not m:
            return None
        month, day, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if year < 100:
            year += 2000
        try:
            return pd.Timestamp(year=year, month=month, day=day).date()
        except Exception:
            return None

    def _loadcirfile(self, filepath: Path) -> pd.DataFrame:
        ext = filepath.suffix.lower()
        if ext == ".csv":
            for enc in ("utf-8", "windows-1252", "iso-8859-1"):
                try:
                    return pd.read_csv(filepath, delimiter=";", encoding=enc)
                except UnicodeDecodeError:
                    continue
            return pd.read_csv(filepath, delimiter=";", encoding="utf-8", errors="replace")
        else:
            return pd.read_excel(filepath)

    def _calculateinventoryvalue(self, df: pd.DataFrame):
        # Try name-based column lookup first, fall back to position
        part_no_col  = next((c for c in df.columns if "PART_NO" in c.upper()), None)
        beg_col      = next((c for c in df.columns if "BEGINNING_INVENTORY" in c.upper()), None)
        yard_col     = next((c for c in df.columns if "INVENTORY_YARD" in c.upper()), None)
        price_col    = next((c for c in df.columns if c.upper() == "PRICE"), None)

        if all([part_no_col, beg_col, yard_col, price_col]):
            stock  = (pd.to_numeric(df[beg_col],   errors="coerce").fillna(0)
                    + pd.to_numeric(df[yard_col],  errors="coerce").fillna(0))
            price  = pd.to_numeric(df[price_col],  errors="coerce").fillna(0)
        else:
            # Positional fallback matching the known CIR column order
            stock  = (pd.to_numeric(df.iloc[:, 5], errors="coerce").fillna(0)
                    + pd.to_numeric(df.iloc[:, 6], errors="coerce").fillna(0))
            price  = pd.to_numeric(df.iloc[:, 11], errors="coerce").fillna(0)

        values = stock * price
        valid  = values[values > 0]
        return float(valid.sum()), int(len(valid))

    def _writetoarchive(self, results: list):
        ARCHIVE_PATH.parent.mkdir(parents=True, exist_ok=True)
        new_df = pd.DataFrame(results)

        if ARCHIVE_PATH.exists():
            existing = pd.read_csv(ARCHIVE_PATH, parse_dates=["date"])
            existing["date"] = existing["date"].dt.normalize()
            new_df["date"]   = pd.to_datetime(new_df["date"]).dt.normalize()

            combined = pd.concat([existing, new_df], ignore_index=True)
            # Keep the most recently added value when dates collide
            combined = (combined.drop_duplicates(subset=["date"], keep="last")
                                .sort_values("date")
                                .reset_index(drop=True))
        else:
            combined = new_df.sort_values("date").reset_index(drop=True)

        combined.to_csv(ARCHIVE_PATH, index=False)

    def closeevent(self, event):
        event.accept()