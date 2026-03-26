import pandas as pd
import json
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple, Optional

from app.utils.config import SHAREDNETWORKPATH

class DataImportManager:
    def __init__(self):
        self.importsdir = SHAREDNETWORKPATH / "imports"
        self.processeddir = SHAREDNETWORKPATH / "processed"
        
        self.importcategories = {
            "current_inventory_report": {
                "name": "Current Inventory Reports",
                "description": "Current Inventory per Part",
                "requiredcolumns": ["MFG", "PART_NO", "PART_DESC", "BEGINNING_INVENTORY_TODAY", "INVENTORY_YARD_TODAY", "INVENTORY_PORT_TODAY", "SUPP_SHP_COUNTRY", "SUPP_NAME", "PRICE"],
                "filetypes": [".csv", ".xlsx", ".xls", ".xlsm"],
                "archive": True},
            "master_data": {
                "name": "Master Data Reports",
                "description": "Part Information",
                "requiredcolumns": ["PART", "PART_DESC", "SUPP_NAME", "SUPP_SHP", "SUPP_SHP_COUNTRY", "SCC_NAME", "PRICE", "UNIT_LOAD_QTY", "MULT_UNIT_LOAD_VALID", "SAFETY", "SHIP_QTY", "STOCK", "FIXED_PERIOD", "TTT_DAYS", "CONSOLIDATOR"],
                "filetypes": [".csv", ".xlsx", ".xls", ".xlsm"],
                "archive": False},
            "part_requirement_split_1": {
                "name": "Part Requirement Split 1",
                "description": "Planned Consumption per Day per Part",
                "requiredcolumns": ["ARTNR", "PART_DESCRIPTION", "PRODDAG", "WEEK", "SKIFT", "ARTAN"],
                "filetypes": [".csv", ".xlsx", ".xls", ".xlsm"],
                "archive": False},
            "part_requirement_split_2": {
                "name": "Part Requirement Split 2",
                "description": "Planned Consumption per Day per Part",
                "requiredcolumns": ["ARTNR", "PART_DESCRIPTION", "PRODDAG", "WEEK", "SKIFT", "ARTAN"],
                "filetypes": [".csv", ".xlsx", ".xls", ".xlsm"],
                "archive": False},
            "part_requirement_split_3": {
                "name": "Part Requirement Split 3",
                "description": "Planned Consumption per Day per Part",
                "requiredcolumns": ["ARTNR", "PART_DESCRIPTION", "PRODDAG", "WEEK", "SKIFT", "ARTAN"],
                "filetypes": [".csv", ".xlsx", ".xls", ".xlsm"],
                "archive": False},
            "goods_to_be_received": {
                "name": "Goods to be Received",
                "description": "Goods to be received by part per day for the next 365 days",
                "requiredcolumns": ["ARTNR", "ARTAN", "FRAKT_TID_TIDIGAST", "FRAKT_TID_SENAST", "ANK_TID_TIDIGAST", "ANK_TID_SENAST"],
                "filetypes": [".csv", ".xlsx", ".xls", ".xlsm"],
                "archive": False},
            "splunk_receiving_data": {
                "name": "Splunk Receiving Data",
                "description": "Goods to be received by part per day for the next 60 days",
                "requiredcolumns": ['TO Number', 'Load Delivery Date Original TO', 'Load Delivery Date Final', 'Part Number', 'Quantity'],
                "filetypes": [".csv", ".xlsx", ".xls", ".xlsm"],
                "archive": False},
            "manual_TTT": {
                "name": "Manual TTT",
                "description": "Delivery Days per Supplier",
                "requiredcolumns": ["FRAKTDAG", "LEVNR"],
                "filetypes": [".csv", ".xlsx", ".xls", ".xlsm"],
                "archive": False},
            "goods_to_be_departed": {
                "name": "Goods to be Departed",
                "description": "Goods to be departed by part per day for the next 90 days",
                "requiredcolumns": ['PART_NO', 'CALL_OFF_QUANTITY', 'EARLIEST_SHIPING_TIME'],
                "filetypes": [".csv", ".xlsx", ".xls", ".xlsm"],
                "archive": True},
            "alert_report": {
                "name": "Alert Report",
                "description": "Alert Report",
                "requiredcolumns": ['PART', 'PART_DESCRIPTION', 'CURRENT_INVENTORY', 'ON_YARD_INVENTORY', 'CURRENT_REQUIREMENT', 'ASN_INTRANSIT', 'SUPPLIER_NAME', 'SUPPLIER_COUNTRY', 'SCC_NAME', 'ALERT_TYPE', 'ALERT_DETAILS'],
                "filetypes": [".csv", ".xlsx", ".xls", ".xlsm"],
                "archive": False},
            "part_matrix": {
                "name": "Part Matrix",
                "description": "Part Matrix with Program Support data",
                "requiredcolumns": ["Part No", "Type 110 (V536)", "Type 100 (P519)"],
                "filetypes": [".csv", ".xlsx", ".xls", ".xlsm"],
                "archive": False}
            }
        self.ensuredirectories()
    
    def ensuredirectories(self):
        for category in self.importcategories.keys():
            (self.importsdir / category).mkdir(parents=True, exist_ok=True)
            (self.processeddir / category).mkdir(parents=True, exist_ok=True)
    
    def getimportcategories(self) -> Dict:
        return self.importcategories
    
    def validatefile(self, filepath: Path, category: str) -> Tuple[bool, List[str], pd.DataFrame]:
        errors = []
        previewdata = pd.DataFrame()
        try:
            if filepath.suffix.lower() not in self.importcategories[category]["filetypes"]:
                errors.append(f"Invalid file type. Expected: {self.importcategories[category]['filetypes']}")
                return False, errors, previewdata
            
            if filepath.suffix.lower() == '.csv':
                for encoding in ['utf-8', 'windows-1252', 'iso-8859-1', 'cp1252']:
                    try:
                        df = pd.read_csv(filepath, delimiter=";", encoding=encoding)
                        print(f"Successfully read with encoding: {encoding}")
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    df = pd.read_csv(filepath, delimiter=";", encoding='utf-8', errors='replace')
            else:
                df = pd.read_excel(filepath)
            
            if df.empty:
                errors.append("File is empty")
                return False, errors, previewdata
            
            requiredcols = self.importcategories[category]["requiredcolumns"]
            missingcols = [col for col in requiredcols if col not in df.columns]
            
            if missingcols:
                errors.append(f"Missing required columns: {missingcols}")
                errors.append(f"Available columns: {list(df.columns)}")
            
            if df.isnull().all().any():
                emptycols = df.columns[df.isnull().all()].tolist()
                errors.append(f"Columns with no data: {emptycols}")
        
            previewdata = df.head(10)
            isvalid = len([e for e in errors if "Missing required columns" in e]) == 0
            
            return isvalid, errors, previewdata
        
        except Exception as e:
            errors.append(f"Error reading file: {str(e)}")
            return False, errors, previewdata
        
    def importfile(self, sourcepath: Path, category: str, username: str, description: str = "") -> Tuple[bool, str]:
        try:
            isvalid, errors, _ = self.validatefile(sourcepath, category)
            if not isvalid:
                return False, f"Validation failed: {'; '.join(errors)}"
            
            shouldarchive = self.importcategories[category].get("archive", False)
            if shouldarchive:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                newfilename = f"{timestamp}_{sourcepath.name}"
            else:
                newfilename = f"latest_{sourcepath.name}"
                existingfiles = list((self.importsdir / category).glob(f"latest_*{sourcepath.suffix}"))
                for existingfile in existingfiles:
                    existingfile.unlink()
                    oldmetadata = existingfile.with_suffix('.json')
                    if oldmetadata.exists():
                        oldmetadata.unlink()
            
            destpath = self.importsdir / category / newfilename
            
            import shutil
            shutil.copy2(sourcepath, destpath)
            
            metadata = {
                "originalfilename": sourcepath.name,
                "importedfilename": newfilename,
                "category": category,
                "importedby": username,
                "importedat": datetime.now().isoformat(),
                "description": description,
                "filesize": destpath.stat().st_size,
                "validationpassed": True,
                "archived": shouldarchive
                }
            
            metadatapath = destpath.with_suffix('.json')
            with open(metadatapath, 'w') as f:
                json.dump(metadata, f, indent=2)
                
            if shouldarchive:
                return True, f"Successfully imported and archived {sourcepath.name} to {category}"
            else:
                return True, f"Successfully imported {sourcepath.name} to {category} (replaced previous version)"
            
        except Exception as e:
            return False, f"Import failed: {str(e)}"
        
    def getimporthistory(self, category: Optional[str] = None) -> List[Dict]:
        history = []
        categories = [category] if category else self.importcategories.keys()
        for cat in categories:
            importdir = self.importsdir / cat
            if not importdir.exists():
                continue
            
            for metadatafile in importdir.glob("*.json"):
                try:
                    with open(metadatafile, 'r') as f:
                        metadata = json.load(f)
                    metadata['categoryname'] = self.importcategories[cat]["name"]
                    history.append(metadata)
                except Exception as e:
                    print(f"Error reading metadata {metadatafile}: {e}")
                    
        history.sort(key=lambda x: x['importedat'], reverse=True)
        return history
    
    def getlatestfile(self, category: str) -> Optional[Path]:
        importdir = self.importsdir / category
        if not importdir.exists():
            return None
        
        shouldarchive = self.importcategories[category].get("archive", False)
        datafiles = []
        for ext in ['.csv', '.xlsx', '.xls', '.xlsm']:
            if shouldarchive:
                datafiles.extend(importdir.glob(f"*{ext}"))
            else:
                datafiles.extend(importdir.glob(f"latest_*{ext}"))
            
        if not datafiles:
            return None
        
        return max(datafiles, key=lambda f: f.stat().st_mtime)
    
    def loaddata(self, category: str, filename: Optional[str] = None) -> pd.DataFrame:
        if filename:
            filepath = self.importsdir / category / filename
        else:
            filepath = self.getlatestfile(category)
        
        if not filepath or not filepath.exists():
            return pd.DataFrame()
        
        try:
            if filepath.suffix.lower() == '.csv':
                for encoding in ['utf-8', 'windows-1252', 'iso-8859-1', 'cp1252']:
                    try:
                        return pd.read_csv(filepath, delimiter=";", encoding=encoding)
                    except UnicodeDecodeError:
                        continue
                return pd.read_csv(filepath, delimiter=";", encoding='utf-8', errors='replace')
            else:
                return pd.read_excel(filepath)
            
        except Exception as e:
            print(f"Error loading data from {filepath}: {e}")
            return pd.DataFrame()
        
    def deleteimport(self, category: str, filename: str, username: str) -> Tuple[bool, str]:
        try:
            filepath = self.importsdir / category / filename
            metadatapath = filepath.with_suffix('.json')
            
            if filepath.exists():
                filepath.unlink()
            if metadatapath.exists():
                metadatapath.unlink()
            
            return True, f"Successfully deleted {filename}"
        
        except Exception as e:
            return False, f"Delete failed: {str(e)}"
