from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox, QTabWidget, QWidget
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

from app.auth.local_auth import LocalAuth
from app.utils.config import PASSWORDMINLENGTH

class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.auth = LocalAuth()
        self.userdata = None
        self.setWindowTitle("Material Planning and Logistics Hub - Login")
        self.setMinimumWidth(400)
        self.setModal(True)
        self.setupui()

    def setupui(self):
        layout = QVBoxLayout()
        header = QWidget()
        header.setStyleSheet("background-color: #156082;")
        headerlayout = QVBoxLayout(header)
        headerlayout.setContentsMargins(0, 4, 0, 4)

        title = QLabel("VCCH MP&L Hub")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setStyleSheet("color: white; background-color: transparent;")
        title.setAlignment(Qt.AlignCenter)
        headerlayout.addWidget(title)

        layout.addWidget(header)
        if self.auth.userexists():
            self.setuploginform(layout, headerlayout)
        else:
            self.setupregistrationform(layout, headerlayout)
        self.setLayout(layout)

    def setuploginform(self, layout, headerlayout=None):
        userdata = self.auth.getuserdata()
        existingusername = userdata.get('username', '')

        label = QLabel(f"Welcome back, {existingusername}!")
        label.setFont(QFont("Arial", 12))
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("color: white; background-color: transparent;")
        headerlayout.addWidget(label)

        layout.addSpacing(10)

        usernamelayout = QHBoxLayout()
        usernamelabel = QLabel("Username:")
        usernamelabel.setFont(QFont("Arial", 10, QFont.Bold))
        usernamelayout.addWidget(usernamelabel)
        self.usernameinput = QLineEdit(existingusername)
        self.usernameinput.setReadOnly(True)
        usernamelayout.addWidget(self.usernameinput)
        layout.addLayout(usernamelayout)

        passwordlayout = QHBoxLayout()
        passwordlabel = QLabel("Password:")
        passwordlabel.setFont(QFont("Arial", 10, QFont.Bold))
        passwordlayout.addWidget(passwordlabel)
        self.passwordinput = QLineEdit()
        self.passwordinput.setEchoMode(QLineEdit.Password)
        self.passwordinput.returnPressed.connect(self.handlelogin)
        passwordlayout.addWidget(self.passwordinput)
        layout.addLayout(passwordlayout)

        layout.addSpacing(10)
        buttonlayout = QHBoxLayout()

        loginbtn = QPushButton("Login")
        loginbtn.clicked.connect(self.handlelogin)
        loginbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 8px 16px; border: 2px; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
        loginbtn.setDefault(True)
        buttonlayout.addWidget(loginbtn)

        changebtn = QPushButton("Different User")
        changebtn.clicked.connect(self.handlechangeuser)
        changebtn.setStyleSheet("""QPushButton {background-color: #d0d0d0; color: black; padding: 8px 16px; border: 10px; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #da190b; color: white}""")
        buttonlayout.addWidget(changebtn)

        layout.addLayout(buttonlayout)

    def setupregistrationform(self, layout, headerlayout=None):
        label = QLabel("Create Your Account")
        label.setFont(QFont("Arial", 12))
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("color: white; background-color: transparent;")
        headerlayout.addWidget(label)

        layout.addSpacing(10)

        usernamelayout = QHBoxLayout()
        usernamelabel = QLabel("Username:")
        usernamelabel.setFont(QFont("Arial", 10, QFont.Bold))
        usernamelayout.addWidget(usernamelabel)
        self.usernameinput = QLineEdit()
        self.usernameinput.setPlaceholderText("Enter username")
        usernamelayout.addWidget(self.usernameinput)
        layout.addLayout(usernamelayout)

        passwordlayout = QHBoxLayout()
        passwordlabel = QLabel("Password:")
        passwordlabel.setFont(QFont("Arial", 10, QFont.Bold))
        passwordlayout.addWidget(passwordlabel)
        self.passwordinput = QLineEdit()
        self.passwordinput.setEchoMode(QLineEdit.Password)
        self.passwordinput.setPlaceholderText(f"Minimum {PASSWORDMINLENGTH} characters")
        passwordlayout.addWidget(self.passwordinput)
        layout.addLayout(passwordlayout)

        confirmpasswordlayout = QHBoxLayout()
        confirmpasswordlayout.addWidget(QLabel("Confirm:"))
        self.confirminput = QLineEdit()
        self.confirminput.setEchoMode(QLineEdit.Password)
        self.confirminput.setPlaceholderText("Re-enter password")
        confirmpasswordlayout.addWidget(self.confirminput)
        layout.addLayout(confirmpasswordlayout)

        layout.addSpacing(10)

        registerbtn = QPushButton("Create Account")
        registerbtn.clicked.connect(self.handleregister)
        registerbtn.setStyleSheet("""QPushButton {background-color: #156082; color: white; padding: 8px 16px; border: 2px; border-radius: 5px; font-weight: bold;} QPushButton:hover {background-color: #45a049;}""")
        registerbtn.setDefault(True)
        layout.addWidget(registerbtn)

    def handlelogin(self):
        username = self.usernameinput.text().strip()
        password = self.passwordinput.text()
        if not username or not password:
            QMessageBox.warning(self, "Error", "Please enter username and password")
            return
        success, userdata = self.auth.authenticate(username, password)
        if success:
            self.userdata = userdata
            self.accept()
        else:
            QMessageBox.warning(self, "Login Failed", "Invalid username or password")
            self.passwordinput.clear()
            self.passwordinput.setFocus()

    def handleregister(self):
        username = self.usernameinput.text().strip()
        password = self.passwordinput.text()
        confirm = self.confirminput.text()
        if not username:
            QMessageBox.warning(self, "Error", "Please enter a username")
            return
        if len(password) < PASSWORDMINLENGTH:
            QMessageBox.warning(self, "Error", f"Password must be at least {PASSWORDMINLENGTH} characters")
            return
        if password != confirm:
            QMessageBox.warning(self, "Error", "Passwords do not match")
            return
        try:
            self.userdata = self.auth.createuser(username, password)
            QMessageBox.information(self, "Success", "Account created successfully!")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to create account: {e}")

    def handlechangeuser(self):
        reply = QMessageBox.question(self, "Change User", "WARNING: This will delete the current user account from this computer. Continue?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.auth.deleteuser()
            self.close()

    def getuserdata(self):
        return self.userdata
