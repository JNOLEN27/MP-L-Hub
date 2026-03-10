from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QGridLayout, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from pathlib import Path

from app.data.import_manager import DataImportManager
from app.utils.config import COLORPRIMARY, COLORSUCCESS

class DataImportsWindow(QMainWindow):
    def __init__(self, userdata, parent=None):
        super().__init__(parent)
        self.userdata = userdata
        self.importmanager = DataImportManager()
        
        self.setWindowTitle("Data Imports - Admin Panel")
        self.resize(1000,700)
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
             "filetypes": "Excel files (*.xlsx *.xls *.xlsm);;CSV files (*.csv)"},
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
             "filetypes": "Excel files (*.xlsx *.xls *.xlsm);;CSV files (*.csv)"}
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
        widget = QWidget()
        widget.setMinimumHeight(120)
        widget.setStyleSheet(f"""QWidget {{background-color: {COLORPRIMARY}; border: 2px solid #ddd; border-radius: 8px; padding: 10px;}} QWidget:hover {{border-color: #f8f9fa; background-color: white;}}""")
        
        layout = QVBoxLayout()
        
        titlelayout = QHBoxLayout()
        
        titlelabel = QLabel(buttoninfo["title"])
        titlelabel.setFont(QFont("Arial", 12, QFont.Bold))
        titlelayout.addWidget(titlelabel)
        titlelayout.addStretch()
        
        layout.addLayout(titlelayout)
        
        descriptionlabel = QLabel(buttoninfo["description"])
        descriptionlabel.setStyleSheet("color: #666; font-size: 10px;")
        descriptionlabel.setWordWrap(True)
        layout.addWidget(descriptionlabel)
        
        importbtn = QPushButton("Import File")
        importbtn.setStyleSheet(f"""QPushButton {{background-color: {COLORPRIMARY}; color: white; border: none; padding: 8px; border-radius: 4px; font-weight: bold;}} QPushButton:hover {{background-color: #45a049;}}""")
        importbtn.clicked.connect(lambda checked, info=buttoninfo: self.handleimport(info))
        
        layout.addWidget(importbtn)
        
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
            if isinstance(self.userdata, dict):
                username = self.userdata['username']
            else:
                username = self.userdata
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
        
    def closeevent(self, event):
        event.accept()