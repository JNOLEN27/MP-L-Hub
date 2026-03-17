import re

from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QMessageBox, QGridLayout, QSizePolicy
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QIcon


class WrappedButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(parent)
        self._label = QLabel(text, self)
        self._label.setWordWrap(True)
        self._label.setAlignment(Qt.AlignCenter)
        self._label.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)

    def setText(self, text):
        self._label.setText(text)

    def setStyleSheet(self, style):
        super().setStyleSheet(style)
        block = re.search(r'QPushButton\s*\{([^}]*)\}', style)
        color = None
        if block:
            match = re.search(r'(?<!-)color\s*:\s*([^;]+)', block.group(1))
            if match:
                color = match.group(1).strip()
        self._label.setStyleSheet(f"color: {color}; background-color: transparent;" if color else "background-color: transparent;")

    def resizeEvent(self, event):
        self._label.setGeometry(self.rect())
        super().resizeEvent(event)

from app.auth.permissions import PermissionsManager
from app.launcher.access_request_dialog import AccessRequestDialog
from app.utils.config import WINDOWTITLE, LAUNCHERWINDOWSIZE, AVAILABLEAPPS, COLORPRIMARY, COLORSUCCESS, ADMINUSERS

class LauncherWindow(QMainWindow):
    openapprequested = pyqtSignal(str)
    def __init__(self, userdata):
        super().__init__()
        self.userdata = userdata
        self.permissions = PermissionsManager()
        self.setWindowTitle(WINDOWTITLE)
        self.resize(*LAUNCHERWINDOWSIZE)
        self.setupui()
        
    def setupui(self):
        centralwidget = QWidget()
        self.setCentralWidget(centralwidget)
        
        layout = QVBoxLayout()
        layout.setSpacing(20)
        
        header = self.createheader()
        layout.addWidget(header)
        
        appswidget = self.createappsgrid()
        layout.addWidget(appswidget)
        
        layout.addStretch()
        
        footer = self.createfooter()
        layout.addWidget(footer)
        
        centralwidget.setLayout(layout)
        
    def createheader(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        title = QLabel("VCCH Material Planning and Logistics Hub")
        title.setFont(QFont("Arial",18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet(f"color: {COLORPRIMARY};")
        layout.addWidget(title)
        
        welcome = QLabel(f"Welcome, {self.userdata['username']}!")
        welcome.setFont(QFont("Arial",12))
        welcome.setAlignment(Qt.AlignCenter)
        layout.addWidget(welcome)
        
        widget.setLayout(layout)
        return widget
    
    def createappsgrid(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        label = QLabel("Available Applications")
        label.setFont(QFont("Arial", 14, QFont.Bold))
        layout.addWidget(label)
        
        grid = QGridLayout()
        grid.setSpacing(15)
        
        userapps = self.permissions.getuserapps(self.userdata['userid'])
        
        row = 0
        col = 0
        maxcols = 2
        for appkey, appinfo in AVAILABLEAPPS.items():
            appname = appinfo['name']
            hasaccess = appname in userapps
            
            btn = WrappedButton(appinfo['name'])
            btn.setMinimumHeight(100)
            btn.setFont(QFont("Arial", 11))
            
            if hasaccess:
                btn.setStyleSheet(f"""QPushButton {{background-color: #45a049; color: green; border: none; border-radius: 5px; padding: 10px;}} QPushButton:hover {{background-color: {COLORSUCCESS};}}""")
                btn.clicked.connect(lambda checked, key=appkey: self.openapp(key))
            else:
                btn.setStyleSheet("""QPushButton {background-color: #cccccc; color: #666666; border: none; border-radius: 5px; padding: 10px;}""")
        
            grid.addWidget(btn, row, col)
            col += 1
            if col >= maxcols:
                col = 0
                row += 1
        layout.addLayout(grid)
        widget.setLayout(layout)
        return widget
    
    def createfooter(self):
        widget = QWidget()
        layout = QHBoxLayout()
        
        requestbtn = QPushButton("Request Access to Apps")
        requestbtn.clicked.connect(self.requestaccess)
        layout.addWidget(requestbtn)
        
        if self.userdata['username'] in ADMINUSERS:
            adminbtn = QPushButton("Admin Panel")
            adminbtn.clicked.connect(self.openadminpanel)
            adminbtn.setStyleSheet(f"""QPushButton {{background-color: {COLORPRIMARY}; color: white; padding: 8px 15px;}}""")
            layout.addWidget(adminbtn)
            
            dataimportbtn = QPushButton("Data Imports")
            dataimportbtn.clicked.connect(self.opendataimportpanel)
            dataimportbtn.setStyleSheet(f"""QPushButton {{background-color: {COLORPRIMARY}; color: white; padding: 8px 15px;}}""")
            layout.addWidget(dataimportbtn)
        
        refreshbtn = QPushButton("Refresh Applications")
        refreshbtn.clicked.connect(self.refreshui)
        layout.addWidget(refreshbtn)
        
        layout.addStretch()
        
        exitbtn = QPushButton("Exit")
        exitbtn.clicked.connect(self.close)
        layout.addWidget(exitbtn)
        
        widget.setLayout(layout)
        return widget
    
    def refreshui(self):
        if self.centralWidget():
            self.centralWidget().setParent(None)
        self.setupui()
    
    def openapp(self, appkey):
        self.openapprequested.emit(appkey)
    
    def requestaccess(self):
        dialog = AccessRequestDialog(self.userdata['userid'], self.userdata['username'],self)
        if dialog.exec_():
            self.setupui()
    
    def openadminpanel(self):
        from app.admin.request_panel import AdminWindow
        
        adminwindow = AdminWindow(self.userdata['username'], self)
        adminwindow.show()
        
    def opendataimportpanel(self):
        from app.admin.data_imports import DataImportsWindow
        
        dataimport = DataImportsWindow(self.userdata['username'], self)
        dataimport.show()
        
    def closeevent(self, event):
        reply = QMessageBox.question(self, "Exit Application", "Are you sure you want to exit?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()