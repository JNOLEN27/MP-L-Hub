"""
Auto-updater for MP&L Hub desktop application.

How it works:
  1. On startup the launcher calls checkforupdate().
  2. It reads  <shared_drive>/updates/version.txt  and compares with the
     locally installed version (APP_VERSION in config.py / version.txt).
  3. If the network has a newer version the user is prompted.
  4. On confirmation a small batch script is written to %TEMP%, launched
     detached, and the app exits.  The batch script:
       - waits a few seconds for the process to fully close
       - uses robocopy to copy the new build from the shared drive
       - relaunches MPL_Hub.exe
       - deletes itself

Developer workflow to push an update:
  1. Bump version.txt in the project root (e.g. 1.4.1).
  2. Run build_app.bat — when prompted, choose to push to the shared drive.
     The script will robocopy dist/MPL_Hub to <shared_drive>/updates/MPL_Hub
     and write the new version.txt there automatically.
  Users get prompted on their next app launch.
"""

import sys
import subprocess
import tempfile
from pathlib import Path

from PyQt5.QtWidgets import QMessageBox

from app.utils.config import APP_VERSION, getsharednetworkpath


def _versiontuple(v: str):
    """Convert "1.2.3" → (1, 2, 3) for comparison."""
    try:
        return tuple(int(x) for x in v.strip().split('.'))
    except (ValueError, AttributeError):
        return (0,)


def checkforupdate(parent=None) -> bool:
    """
    Check the shared drive for a newer build.
    Returns True if the user accepted an update (caller should exit the app).
    """
    if not getattr(sys, 'frozen', False):
        # Running from source — never auto-update
        return False

    try:
        network_version_file = getsharednetworkpath() / "updates" / "version.txt"
        if not network_version_file.exists():
            return False

        network_version = network_version_file.read_text().strip()
        if _versiontuple(network_version) <= _versiontuple(APP_VERSION):
            return False

        reply = QMessageBox.question(
            parent,
            "Update Available",
            f"A new version of MP&L Hub is available.\n\n"
            f"  Installed:  {APP_VERSION}\n"
            f"  Available:  {network_version}\n\n"
            f"Update now?  The app will close and restart automatically.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )

        if reply == QMessageBox.Yes:
            return _applyupdate(parent)

    except Exception as e:
        print(f"[Updater] Check failed: {e}")

    return False


def manualcheck(parent=None):
    """Triggered by the 'Check for Updates' button — shows a result either way."""
    if not getattr(sys, 'frozen', False):
        QMessageBox.information(parent, "Updates", "Update checking is only available in the installed build.")
        return

    try:
        network_version_file = getsharednetworkpath() / "updates" / "version.txt"
        if not network_version_file.exists():
            QMessageBox.information(parent, "Updates", "No update file found on the shared drive.")
            return

        network_version = network_version_file.read_text().strip()
        if _versiontuple(network_version) <= _versiontuple(APP_VERSION):
            QMessageBox.information(
                parent, "Up to Date",
                f"You are on the latest version ({APP_VERSION})."
            )
            return

        # Newer version found — reuse the same prompt
        reply = QMessageBox.question(
            parent,
            "Update Available",
            f"A new version of MP&L Hub is available.\n\n"
            f"  Installed:  {APP_VERSION}\n"
            f"  Available:  {network_version}\n\n"
            f"Update now?  The app will close and restart automatically.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if reply == QMessageBox.Yes:
            if _applyupdate(parent):
                parent.close() if parent else None

    except Exception as e:
        QMessageBox.warning(parent, "Update Check Failed", f"Could not reach the shared drive:\n{e}")


def _applyupdate(parent=None) -> bool:
    """
    Write an updater batch script to %TEMP%, launch it detached, and return
    True so the caller knows to exit the application.
    """
    try:
        app_dir = Path(sys.executable).resolve().parent
        exe_path = Path(sys.executable).resolve()
        source_dir = getsharednetworkpath() / "updates" / "MPL_Hub"

        bat_lines = [
            "@echo off",
            "echo MP^&L Hub is updating, please wait...",
            "timeout /t 4 /nobreak > nul",
            f'robocopy "{source_dir}" "{app_dir}" /E /IS /IT /COPY:DAT /R:5 /W:3 > nul',
            "if errorlevel 8 (",
            "    echo Update failed. Please copy the new version manually from:",
            f'    echo {source_dir}',
            "    pause",
            "    exit /b 1",
            ")",
            f'start "" "{exe_path}"',
            'del "%~f0"',
        ]

        bat_path = Path(tempfile.gettempdir()) / "_mplhub_update.bat"
        bat_path.write_text("\r\n".join(bat_lines), encoding="utf-8")

        subprocess.Popen(
            ["cmd.exe", "/c", str(bat_path)],
            creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
            close_fds=True,
        )
        return True

    except Exception as e:
        QMessageBox.critical(
            parent,
            "Update Failed",
            f"Could not launch the updater:\n{e}\n\n"
            f"Please update manually by copying the new build from:\n"
            f"{getsharednetworkpath() / 'updates' / 'MPL_Hub'}",
        )
        print(f"[Updater] Apply failed: {e}")
        return False
