from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QCheckBox, QMessageBox, QScrollArea, QWidget
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

from app.auth.permissions import PermissionsManager
from app.utils.config import AVAILABLEAPPS

class AccessRequestDialog(QDialog):
    def __init__(self, userid, username, parent=None):
        super().__init__(parent)
        self.userid = userid
        self.username = username
        self.permissions = PermissionsManager()
        self.setWindowTitle("Request Application Access")
        self.setMinimumWidth(500)
        self.setMinimumHeight(400)
        self.setModal(True)
        self.appcheckboxes = {}
        self.setupui()
    
    def setupui(self):
        layout = QVBoxLayout()
        title = QLabel("Request Access to Applications")
        title.setFont(QFont("Arial", 14, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        instructions = QLabel("Select the applications you need to access to. \n" "An administrator will review and approve your request. Reach out to Jack Nolen for escalation.")
        instructions.setWordWrap(True)
        instructions.setAlignment(Qt.AlignCenter)
        layout.addWidget(instructions)
        
        layout.addSpacing(20)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scrollwidget = QWidget()
        scrolllayout = QVBoxLayout()
        
        for appkey, appinfo in AVAILABLEAPPS.items():
            hasaccess = self.permissions.checkaccess(self.userid, appinfo['name'])
            haspending = self.permissions.haspendingrequest(self.userid, appinfo['name'])
            
            checkbox = QCheckBox(f"{appinfo['name']}")
            checkbox.setFont(QFont("Arial", 11))
            
            desclabel = QLabel(f"    {appinfo['description']}")
            desclabel.setWordWrap(True)
            desclabel.setStyleSheet("color: #666; margin-left: 20px;")
            
            if hasaccess:
                checkbox.setChecked(True)
                checkbox.setEnabled(False)
                checkbox.setText(f"{checkbox.text()}✓")
                checkbox.setStyleSheet("color: green;")
            elif haspending:
                checkbox.setChecked(True)
                checkbox.setEnabled(False)
                checkbox.setText(f"{checkbox.text()}⏳")
                checkbox.setStyleSheet("color: orange;")
            else:
                checkbox.setChecked(False)
            
            scrolllayout.addWidget(checkbox)
            scrolllayout.addWidget(desclabel)
            scrolllayout.addSpacing(10)
            self.appcheckboxes[appinfo['name']] = checkbox
        
        scrolllayout.addStretch()
        scrollwidget.setLayout(scrolllayout)
        scroll.setWidget(scrollwidget)
        layout.addWidget(scroll)
        
        buttonlayout = QVBoxLayout()
        submitbtn = QPushButton("Submit Request")
        submitbtn.clicked.connect(self.handlesubmit)
        submitbtn.setDefault(True)
        buttonlayout.addWidget(submitbtn)
        
        cancelbtn = QPushButton("Cancel")
        cancelbtn.clicked.connect(self.reject)
        buttonlayout.addWidget(cancelbtn)
        
        layout.addLayout(buttonlayout)
        
        self.setLayout(layout)
        
    def handlesubmit(self):
        selectedapps = []
        for appname, checkbox in self.appcheckboxes.items():
            if checkbox.isChecked() and checkbox.isEnabled():
                selectedapps.append(appname)
        if not selectedapps:
            QMessageBox.information(self, "No Selection", "No new applications selected for access request.")
            return
        try:
            for appname in selectedapps:
                self.permissions.submitaccessrequest(self.userid, self.username, appname)
            QMessageBox.information(self, "Request Submitted", f"Access request submitted for {len(selectedapps)} application(s). \n\n" "You will be notified when an adimistrator approves your request. \n\n" "Reach out to Jack Nolen for escalation.")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to submit request: {e}")