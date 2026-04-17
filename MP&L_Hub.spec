# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec file for MP&L Hub
# Build command:  pyinstaller "MP&L_Hub.spec"
#
# Requirements: pip install pyinstaller
#

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

# Explicitly collect all files for matplotlib and mpl_toolkits
matplotlib_datas, matplotlib_binaries, matplotlib_hiddenimports = collect_all('matplotlib')
mpl_datas, mpl_binaries, mpl_hiddenimports = collect_all('mpl_toolkits')

a = Analysis(
    ['MP&L_Hub.py'],
    pathex=[str(Path('.').resolve())],
    binaries=[*matplotlib_binaries, *mpl_binaries],
    datas=[
        # Include the entire app/ package
        ('app', 'app'),
        # Include version file so the app knows its own version at runtime
        ('version.txt', '.'),
        *matplotlib_datas,
        *mpl_datas,
    ],
    hiddenimports=[
        # PyQt5
        'PyQt5',
        'PyQt5.QtWidgets',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.sip',
        # Data / analysis
        'pandas',
        'numpy',
        'bcrypt',
        'tqdm',
        # Matplotlib - fully collected via collect_all above
        *matplotlib_hiddenimports,
        *mpl_hiddenimports,
        # Standard library modules sometimes missed by PyInstaller
        'json',
        'pathlib',
        'datetime',
        'hashlib',
        'hmac',
        'base64',
        'uuid',
        'shutil',
        # App modules (explicitly list all submodules)
        'app',
        'app.utils',
        'app.utils.config',
        'app.utils.updater',
        'app.data',
        'app.data.import_manager',
        'app.auth',
        'app.auth.local_auth',
        'app.auth.login_encryption',
        'app.auth.permissions',
        'app.launcher',
        'app.launcher.main',
        'app.launcher.login_dialog',
        'app.launcher.launcher_window',
        'app.launcher.access_request_dialog',
        'app.admin',
        'app.admin.request_panel',
        'app.admin.data_imports',
        'app.supply_chain_coordination',
        'app.supply_chain_coordination.main_window',
        'app.supply_chain_coordination.coverage_analysis',
        'app.supply_chain_coordination.waterfall_analysis',
        'app.supply_chain_coordination.ldjis_coverage',
        'app.supply_chain_coordination.adjustment_store',
        'app.supply_chain_coordination.maintenance_tab',
        'app.inventory_by_purpose',
        'app.inventory_by_purpose.main_window',
        'app.inventory_by_purpose.main_window_minimal',
        'app.inventory_by_purpose.ibp_neural_network',
        'app.inventory_by_purpose.monte_tuc_sim',
        'app.inventory_by_purpose.odbc_config_dialog',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Exclude large optional packages not needed at runtime
        'torch',        # Optional - only needed for neural network features
        'scipy',
        'IPython',
        'notebook',
        'pytest',
        'setuptools',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='MPL_Hub',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,           # No console window — GUI only
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='app_icon.ico',   # Uncomment and point to your .ico file
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='MPL_Hub',
)
