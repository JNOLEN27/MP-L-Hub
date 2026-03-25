import re

from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QMessageBox, QGridLayout, QSizePolicy
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QIcon

from app.auth.permissions import PermissionsManager
from app.launcher.access_request_dialog import AccessRequestDialog
from app.utils.config import WINDOWTITLE, LAUNCHERWINDOWSIZE, AVAILABLEAPPS, COLORPRIMARY, COLORSUCCESS, ADMINUSERS, POWERUSERS

class WrappedButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(parent)
        self._label = QLabel(text, self)
        self._label.setWordWrap(True)
        self._label.setAlignment(Qt.AlignCenter)
        self._label.setFont(QFont("Arial", 16))
        self._label.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
        self._defaultcolor = None
        self._hovercolor = None
 
    def setText(self, text):
        self._label.setText(text)
 
    def _parsecolor(self, block):
        match = re.search(r'(?<!-)color\s*:\s*([^;]+)', block)
        return match.group(1).strip() if match else None
 
    def _applylabelcolor(self, color):
        self._label.setStyleSheet(f"color: {color}; background-color: transparent;" if color else "background-color: transparent;")
 
    def setStyleSheet(self, style):
        super().setStyleSheet(style)
        default_block = re.search(r'QPushButton\s*\{([^}]*)\}', style)
        self._defaultcolor = self._parsecolor(default_block.group(1)) if default_block else None
        hover_block = re.search(r'QPushButton:hover\s*\{([^}]*)\}', style)
        self._hovercolor = self._parsecolor(hover_block.group(1)) if hover_block else None
        self._applylabelcolor(self._defaultcolor)
 
    def enterEvent(self, event):
        self._applylabelcolor(self._hovercolor or self._defaultcolor)
        super().enterEvent(event)
 
    def leaveEvent(self, event):
        self._applylabelcolor(self._defaultcolor)
        super().leaveEvent(event)
    
    def resizeEvent(self, event):
        self._label.setGeometry(self.rect())
        super().resizeEvent(event)

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
        widget.setStyleSheet("background-color: #156082;")
        widgetlayout = QVBoxLayout(widget)
        widgetlayout.setContentsMargins(0, 4, 0, 4)
        
        title = QLabel("VCCH Material Planning and Logistics Hub")
        title.setFont(QFont("Arial",18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: white; background-color: transparent;")
        widgetlayout.addWidget(title)
        
        welcome = QLabel(f"Welcome, {self.userdata['username']}!")
        welcome.setFont(QFont("Arial",12))
        welcome.setAlignment(Qt.AlignCenter)
        welcome.setStyleSheet("color: white; background-color: transparent;")
        widgetlayout.addWidget(welcome)
        
        layout.addWidget(widget)
        return widget
    
    def createappsgrid(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        label = QLabel("Available Applications")
        label.setFont(QFont("Arial", 20, QFont.Bold))
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
            
            if hasaccess:
                btn.setStyleSheet("""QPushButton {background-color: #156082; color: #0e2841; border: none; border-radius: 5px; padding: 10px;} QPushButton:hover {background-color: #45a049; color: green;}""")
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
        requestbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 8px 15px;} QPushButton:hover {background-color: #a2d8f0; color: grey;}""")
        layout.addWidget(requestbtn)
        
        refreshbtn = QPushButton("Refresh Applications")
        refreshbtn.clicked.connect(self.refreshui)
        refreshbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 8px 15px;} QPushButton:hover {background-color: #a2d8f0; color: grey;}""")
        layout.addWidget(refreshbtn)
        
        username = self.userdata['username']
        if username in ADMINUSERS:
            adminbtn = QPushButton("Admin Panel")
            adminbtn.clicked.connect(self.openadminpanel)
            adminbtn.setStyleSheet("""QPushButton {background-color: #800000; color: white; padding: 8px 15px;} QPushButton:hover {background-color: #ffb3b3; color: grey;}""")
            layout.addWidget(adminbtn)
 
        if username in ADMINUSERS or username in POWERUSERS:
            dataimportbtn = QPushButton("Data Imports")
            dataimportbtn.clicked.connect(self.opendataimportpanel)
            dataimportbtn.setStyleSheet("""QPushButton {background-color: #800000; color: white; padding: 8px 15px;} QPushButton:hover {background-color: #ffb3b3; color: grey;}""")
            layout.addWidget(dataimportbtn)
        
        layout.addStretch()
        
        exitbtn = QPushButton("Exit")
        exitbtn.clicked.connect(self.close)
        exitbtn.setStyleSheet("""QPushButton {background-color: #d0d0d0; color: black; padding: 8px 15px;} QPushButton:hover {background-color: #800000; color: white}""")
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
        
        dataimport = DataImportsWindow(self.userdata, self)
        dataimport.show()
        
    def closeevent(self, event):
        reply = QMessageBox.question(self, "Exit Application", "Are you sure you want to exit?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()
