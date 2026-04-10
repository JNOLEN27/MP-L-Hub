import sys
from PyQt5.QtWidgets import QApplication, QMessageBox
from PyQt5.QtCore import Qt

from app.launcher.login_dialog import LoginDialog
from app.launcher.launcher_window import LauncherWindow
from app.launcher.access_request_dialog import AccessRequestDialog
from app.auth.permissions import PermissionsManager

class Application:
    def __init__(self):
        # Enable automatic pixel-ratio scaling so all hardcoded sizes stay
        # proportional regardless of the OS display-scaling factor (125 %, 150 %,
        # 200 %, etc.).  Must be set before QApplication is constructed.
        if hasattr(Qt, 'AA_EnableHighDpiScaling'):
            QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
            QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
        self.app = QApplication(sys.argv)
        self.app.setApplicationName("VCCH Material Planning and Logistics Hub")
        self.userdata = None
        self.launcherwindow = None
        self.openappwindows = {}
        self.permissions = PermissionsManager()
    
    def run(self):
        if not self.showlogin():
            return 0
        userapps = self.permissions.getuserapps(self.userdata['userid'])
        if not userapps:
            reply = QMessageBox.question(None, "No Access", "You don't have access to any applications yet.\n\n" "Would you like to request access now?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.showaccessrequest()
        self.launcherwindow = LauncherWindow(self.userdata)
        self.launcherwindow.openapprequested.connect(self.openapplication)
        self.launcherwindow.show()
        return self.app.exec_()
    
    def showlogin(self):
        dialog = LoginDialog()
        if dialog.exec_():
            self.userdata = dialog.getuserdata()
            return True
        return False
    
    def showaccessrequest(self):
        dialog = AccessRequestDialog(self.userdata['userid'], self.userdata['username'])
        dialog.exec_()
        
    def openapplication(self, appkey):
        if appkey in self.openappwindows:
            window = self.openappwindows[appkey]
            window.raise_()
            window.activateWindow()
            return
        
        if appkey == "inventory_by_purpose":
            from app.inventory_by_purpose.main_window import InventorybyPurposeWindow
            window = InventorybyPurposeWindow(self.userdata, self.launcherwindow)
            
        elif appkey == "supply_chain_coordination":
            from app.supply_chain_coordination.main_window import SupplyChainCoordinationWindow
            window = SupplyChainCoordinationWindow(self.userdata, self.launcherwindow)
            
        else:
            QMessageBox.warning(self.launcherwindow, "Not Implemented", f"Application '{appkey}' is not yet implemented.")
            return
        
        if appkey != "inventory_by_purpose":  # IbP sets this internally
            window.setAttribute(Qt.WA_DeleteOnClose)
        window.destroyed.connect(lambda: self.closeapplication(appkey))
        self.openappwindows[appkey] = window
        window.show()
        
    def closeapplication(self, appkey):
        if appkey in self.openappwindows:
            del self.openappwindows[appkey]
            
def main():
    app = Application()
    sys.exit(app.run())

if __name__ == "__main__":
    main()
