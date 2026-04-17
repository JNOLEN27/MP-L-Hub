import os
import json
import sys
from pathlib import Path

if getattr(sys, 'frozen', False):
    BASEDIR = Path(sys.executable).resolve().parent
    _version_file = Path(sys._MEIPASS) / "version.txt"
else:
    BASEDIR = Path(__file__).resolve().parent.parent.parent
    _version_file = BASEDIR / "version.txt"

APP_VERSION = _version_file.read_text().strip() if _version_file.exists() else "0.0.0"

CONFIGPATH = BASEDIR / "config"

LOCALAPPDATA = Path(os.getenv('LOCALAPPDATA', str(Path.home() / 'AppData' / 'Local'))) / "MP&LHUB"
LOCALAPPDATA.mkdir(parents=True, exist_ok=True)

LOCALUSERCONFIG = LOCALAPPDATA / "user_config.json"
NETWORKCONFIGFILE = LOCALAPPDATA / "network_config.json"


def getsharednetworkpath() -> Path:
    return Path(r"W:\_US Operations\P&M Americas\21200 Supplier Quality & Logistics\21220 Material Planning & Logistic\12 MP&L Hub\Data")


SHAREDNETWORKPATH = getsharednetworkpath()

ACCESSREQUESTSFILE = SHAREDNETWORKPATH / "config" / "access_requests.json"
USERPERMISSIONSFILE = SHAREDNETWORKPATH / "config" / "user_permissions.json"

FORECATSDIR = SHAREDNETWORKPATH / "forecasts"
ARCHIVEDIR = SHAREDNETWORKPATH / "archive"
MASTERDATADIR = SHAREDNETWORKPATH / "master_data"
INVENTORYARCHIVEFILE = ARCHIVEDIR / "Inventory_Archive.csv"

AVAILABLEAPPS = {
    "inventory_by_purpose": {
        "name": "Inventory by Purpose",
        "description": "VCCH Inventory Forecasting and Analysis",
        "module": "inventory_by_purpose"},
    "supply_chain_coordination": {
        "name": "Supply Chain Coordination",
        "description": "Runouts, Coverages, and Analysis",
        "module": "supply_chain_coordination"},
    "plant_supply_chain_engineering": {
        "name": "Plant Supply Chain Engineering",
        "description": "PFEP, KPI's and Reports",
        "module": "plant_supply_chain_engineering"},
    "production_control": {
        "name": "Production Control",
        "description": "Production Schedules, Reports, and Planning",
        "module": "production_control"},
    "program_management": {
        "name": "Program Management",
        "description": "Volume Simulations, What-If Analysis, and Program Schedules",
        "module": "program_management"},
    "running_change": {
        "name": "Running Change",
        "description": "Timings, Archives, and Exports",
        "module": "running_change"}
    }

ADMINUSERS = ["jnolen", "ahuntington", "jmccaslin", "mkaindl", "kwhite", "tgarbachevskaya", "esanta", "sscimemi", "mberry"]

POWERUSERS = {
    "ljohnson": ["master_data", "current_inventory_report", "part_requirement_split_1", "part_requirement_split_2", "part_requirement_split_3", "splunk_receiving_data", "goods_to_be_received", "alert_report", "manual_TTT", "goods_to_be_departed"],
    "cwilliams": ["master_data", "current_inventory_report", "part_requirement_split_1", "part_requirement_split_2", "part_requirement_split_3", "splunk_receiving_data", "goods_to_be_received", "alert_report", "manual_TTT", "goods_to_be_departed"],
    "bmcarn": ["master_data", "current_inventory_report", "part_requirement_split_1", "part_requirement_split_2", "part_requirement_split_3", "splunk_receiving_data", "goods_to_be_received", "alert_report", "manual_TTT", "goods_to_be_departed"],
    "dezell": ["master_data", "current_inventory_report", "part_requirement_split_1", "part_requirement_split_2", "part_requirement_split_3", "splunk_receiving_data", "goods_to_be_received", "alert_report", "manual_TTT", "goods_to_be_departed"]
    }

WINDOWTITLE = "VCCH Material Planning and Logistics Management Hub"
LAUNCHERWINDOWSIZE = (600, 400)
APPWINDOWSIZE = (1400, 900)

COLORPRIMARY = "#0057A3"
COLORSUCCESS = "4CAF50"
COLORWARNING = "FF9800"
COLORERROR = "F44336"
COLORBACKGROUND = "#F5F5F5"

PASSWORDMINLENGTH = 6
PASSWORDHASHROUNDS = 12


def ensuredirectories():
    FORECATSDIR.mkdir(parents=True, exist_ok=True)
    ARCHIVEDIR.mkdir(parents=True, exist_ok=True)
    MASTERDATADIR.mkdir(parents=True, exist_ok=True)
    CONFIGPATH.mkdir(parents=True, exist_ok=True)


def getlatestforecastfile():
    if not FORECATSDIR.exists():
        return None
    forecastfiles = list(FORECATSDIR.glob("forecast_*.parquet"))
    if not forecastfiles:
        forecastfiles = list(FORECATSDIR.glob("forecast_*.csv"))
    if not forecastfiles:
        return None
    return max(forecastfiles, key=lambda f: f.stat().st_mtime)


ensuredirectories()
