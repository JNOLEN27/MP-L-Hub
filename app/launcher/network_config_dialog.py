from pathlib import Path

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QFileDialog, QMessageBox, QFrame
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

from app.utils.config import (
    getsharednetworkpath, setsharednetworkpath, issharednetworkpathconfigured,
    BASEDIR
)


class NetworkConfigDialog(QDialog):
    """Dialog for configuring the shared network drive path."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configure Shared Network Drive")
        self.setMinimumWidth(560)
        self.setModal(True)
        self._setupui()

    def _setupui(self):
        layout = QVBoxLayout()
        layout.setSpacing(12)

        # Title
        title = QLabel("Shared Network Drive Configuration")
        title.setFont(QFont("Arial", 13, QFont.Bold))
        layout.addWidget(title)

        # Explanation
        info = QLabel(
            "All users must point to the <b>same shared folder</b> on your network drive.\n"
            "This folder stores imported data files, coverage comments, alerts notes,\n"
            "and PIWD entries so that everyone sees the same information.\n\n"
            "Example (Windows UNC path):  <tt>\\\\server\\share\\MP-L-Hub-Data</tt>\n"
            "Example (mapped drive):      <tt>Z:\\MP-L-Hub-Data</tt>"
        )
        info.setWordWrap(True)
        info.setStyleSheet("color: #333; padding: 6px;")
        layout.addWidget(info)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #ccc;")
        layout.addWidget(sep)

        # Current path
        current_label = QLabel("Current shared path:")
        current_label.setFont(QFont("Arial", 10, QFont.Bold))
        layout.addWidget(current_label)

        current_path = getsharednetworkpath()
        is_configured = issharednetworkpathconfigured()
        status_text = str(current_path)
        status_color = "#1a7a1a" if is_configured else "#8a6000"
        if not is_configured:
            status_text += "  (using local fallback — shared drive not configured)"

        self._current_display = QLabel(status_text)
        self._current_display.setStyleSheet(
            f"color: {status_color}; font-family: Consolas, monospace; padding: 4px;"
        )
        self._current_display.setWordWrap(True)
        layout.addWidget(self._current_display)

        # Path entry row
        path_label = QLabel("New shared path:")
        path_label.setFont(QFont("Arial", 10, QFont.Bold))
        layout.addWidget(path_label)

        path_row = QHBoxLayout()
        self._path_edit = QLineEdit()
        self._path_edit.setPlaceholderText(r"\\server\share\MP-L-Hub-Data  or  Z:\MP-L-Hub-Data")
        self._path_edit.setText(str(current_path) if is_configured else "")
        path_row.addWidget(self._path_edit)

        browse_btn = QPushButton("Browse...")
        browse_btn.setFixedWidth(90)
        browse_btn.clicked.connect(self._browse)
        path_row.addWidget(browse_btn)
        layout.addLayout(path_row)

        # Note about restart
        note = QLabel(
            "Note: After saving, please restart the application for the new path to take full effect."
        )
        note.setStyleSheet("color: #666; font-style: italic;")
        note.setWordWrap(True)
        layout.addWidget(note)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        save_btn = QPushButton("Save and Close")
        save_btn.setStyleSheet(
            "QPushButton {background-color: #156082; color: white; padding: 8px 18px; border-radius: 4px;}"
            "QPushButton:hover {background-color: #1e88b5;}"
        )
        save_btn.clicked.connect(self._save)
        btn_row.addWidget(save_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet("padding: 8px 18px;")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)

        layout.addLayout(btn_row)
        self.setLayout(layout)

    def _browse(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select Shared Network Folder",
            str(getsharednetworkpath())
        )
        if folder:
            self._path_edit.setText(folder)

    def _save(self):
        path_str = self._path_edit.text().strip()
        if not path_str:
            QMessageBox.warning(self, "No Path", "Please enter or browse to a shared folder path.")
            return

        p = Path(path_str)
        if not p.exists():
            reply = QMessageBox.question(
                self, "Path Not Found",
                f"The folder does not currently exist:\n{path_str}\n\n"
                "Save it anyway? (Useful for a network path that will be available at runtime.)",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        if not setsharednetworkpath(path_str):
            QMessageBox.critical(self, "Save Failed", "Could not save the network path configuration.")
            return

        QMessageBox.information(
            self, "Saved",
            f"Shared network path saved:\n{path_str}\n\n"
            "Please restart the application."
        )
        self.accept()
