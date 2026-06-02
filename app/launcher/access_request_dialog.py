from PyQt5.QtWidgets import QApplication, QDialog, QVBoxLayout, QLabel, QPushButton, QCheckBox, QMessageBox, QScrollArea, QWidget
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
        screen = QApplication.primaryScreen()
        if screen:
            avail = screen.availableGeometry()
            scale = min(avail.width() / 1920.0, avail.height() / 1080.0)
            self._ui_scale = max(0.60, min(1.0, scale))
        else:
            self._ui_scale = 1.0
        self.setWindowTitle("Request Application Access")
        self.setMinimumWidth(max(350, int(500 * self._ui_scale)))
        self.setMinimumHeight(max(300, int(400 * self._ui_scale)))
        self.setModal(True)
        self.appcheckboxes = {}
        self.setupui()

    def _sz(self, n):
        return max(1, int(n * self._ui_scale))
    
    def setupui(self):
        layout = QVBoxLayout()
        header = QWidget()
        header.setStyleSheet("background-color: #156082;")
        headerlayout = QVBoxLayout(header)
        headerlayout.setContentsMargins(0, 4, 0, 4)
 
        title = QLabel("Request Access to Applications")
        title.setFont(QFont("Arial", max(10, self._sz(14)), QFont.Bold))
        title.setStyleSheet("background-color: transparent; color: white")
        title.setAlignment(Qt.AlignCenter)
        headerlayout.addWidget(title)

        instructions = QLabel("Select the applications you need to access to. \n" "An administrator will review and approve your request. Reach out to Jack Nolen for escalation.")
        instructions.setWordWrap(True)
        instructions.setStyleSheet("background-color: transparent; color: white")
        instructions.setAlignment(Qt.AlignCenter)
        headerlayout.addWidget(instructions)

        layout.addWidget(header)

        layout.addSpacing(self._sz(20))
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scrollwidget = QWidget()
        scrolllayout = QVBoxLayout()
        
        for appkey, appinfo in AVAILABLEAPPS.items():
            hasaccess = self.permissions.checkaccess(self.userid, appinfo['name'])
            haspending = self.permissions.haspendingrequest(self.userid, appinfo['name'])
            
            checkbox = QCheckBox(f"{appinfo['name']}")
            checkbox.setFont(QFont("Arial", max(8, self._sz(11))))
            
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
        submitbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 8px 15px;} QPushButton:hover {background-color: #a2d8f0; color: grey;}""")
        submitbtn.setDefault(True)
        buttonlayout.addWidget(submitbtn)
        
        cancelbtn = QPushButton("Cancel")
        cancelbtn.clicked.connect(self.reject)
        cancelbtn.setStyleSheet("""QPushButton {background-color: #d0d0d0; color: black; padding: 8px 15px;} QPushButton:hover {background-color: #800000; color: white}""")
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
