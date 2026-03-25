from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

from app.auth.permissions import PermissionsManager
from app.utils.config import COLORPRIMARY, COLORSUCCESS, COLORERROR

class AdminWindow(QMainWindow):
    def __init__(self, adminusername, parent=None):
        super().__init__(parent)
        self.adminusername = adminusername
        self.permissions = PermissionsManager()
        self.setWindowTitle("Admin Panel - Access Requests")
        self.resize(800, 600)
        self.setupui()
        self.loadrequests()
        
    def setupui(self):
        centralwidget = QWidget()
        self.setCentralWidget(centralwidget)
        layout = QVBoxLayout()
        header = QWidget()
        header.setStyleSheet("background-color: #156082;")
        headerlayout = QVBoxLayout(header)
        headerlayout.setContentsMargins(4, 4, 0, 4)
        
        
        title = QLabel("Pending Access Requests")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setStyleSheet("color: white; background-color: transparent;")
        headerlayout.addWidget(title)
        
        instructions = QLabel("Review and approve or deny user access requests. \n""Select a request and click Approve or Deny.")
        instructions.setStyleSheet("color: white; background-color: transparent;")
        headerlayout.addWidget(instructions)
        
        layout.addWidget(header)
        layout.addSpacing(10)
        
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Request ID", "Username", "Application", "Requested At", "Status"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.table)
        
        buttonlayout = QHBoxLayout()
        self.approvebtn = QPushButton("Approve")
        self.approvebtn.clicked.connect(self.approveselected)
        self.approvebtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 8px 15px;} QPushButton:hover {background-color: #a2d8f0; color: grey;}""")
        buttonlayout.addWidget(self.approvebtn)
        
        self.denybtn = QPushButton("Deny")
        self.denybtn.clicked.connect(self.denyselected)
        self.denybtn.setStyleSheet("""QPushButton {background-color: #800000; color: white; padding: 8px 15px;} QPushButton:hover {background-color: #ffb3b3; color: grey;}""")
        buttonlayout.addWidget(self.denybtn)
        
        buttonlayout.addStretch()
        
        refreshbtn = QPushButton("Refresh")
        refreshbtn.clicked.connect(self.loadrequests)
        refreshbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 8px 15px;} QPushButton:hover {background-color: #a2d8f0; color: grey;}""")
        buttonlayout.addWidget(refreshbtn)
        
        closebtn = QPushButton("Close")
        closebtn.clicked.connect(self.close)
        closebtn.setStyleSheet("""QPushButton {background-color: #d0d0d0; color: black; padding: 8px 15px;} QPushButton:hover {background-color: #800000; color: white}""")
        buttonlayout.addWidget(closebtn)
        
        layout.addLayout(buttonlayout)
        centralwidget.setLayout(layout)
        
    def loadrequests(self):
        requests = self.permissions.getpendingrequests()
        self.table.setRowCount(0)
        for request in requests:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(request['requestid']))
            self.table.setItem(row, 1, QTableWidgetItem(request['username']))
            self.table.setItem(row, 2, QTableWidgetItem(request['requestedapp']))
            self.table.setItem(row, 3, QTableWidgetItem(request['requestedat'][:19]))
            self.table.setItem(row, 4, QTableWidgetItem(request['status']))
            
        if requests:
            self.setWindowTitle(f"Admin Panel - {len(requests)} Pending Requests(s)")
        else:
            self.setWindowTitle(f"Admin Panel - No Pending Requests")
    
    def getselectedrequestid(self):
        selectedrow = self.table.selectedIndexes()
        if not selectedrow:
            return None
        row = selectedrow[0].row()
        return self.table.item(row, 0).text()
    
    def approveselected(self):
        requestid = self.getselectedrequestid()
        if not requestid:
            QMessageBox.warning(self, "No Selection", "Please select a request to approve")
            return
        row = self.table.currentRow()
        username = self.table.item(row, 1).text()
        appname = self.table.item(row, 2).text()
        reply = QMessageBox.question(self,"Confirm Approval",f"Approve access for user '{username}' to '{appname}'?",QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            success = self.permissions.approverequest(requestid, self.adminusername)
            if success:
                QMessageBox.information(self, "Success", "Access request approved!")
                self.loadrequests()
            else:
                QMessageBox.critical(self, "Error", "Failed to approve request")
    
    def denyselected(self):
        requestid = self.getselectedrequestid()
        if not requestid:
            QMessageBox.warning(self, "No Selection", "Please select a request to deny")
            return
        row = self.table.currentRow()
        username = self.table.item(row, 1).text()
        appname = self.table.item(row, 2).text()
        reply = QMessageBox.question(self,"Confirm Denial",f"Deny access for user '{username}' to '{appname}'",QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            success = self.permissions.denyrequest(requestid, self.adminusername)
            if success:
                QMessageBox.information(self, "Success", "Access request denied")
                self.loadrequests()
            else:
                QMessageBox.critical(self, "Error", "Failed to deny request")
