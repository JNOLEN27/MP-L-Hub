import pandas as pd
import os
from datetime import datetime, timedelta, date
from typing import Dict, List, Tuple, Optional
from pathlib import Path

class WaterfallAnalysisEngine:
    def __init__(self, import_manager):
        self.import_manager = import_manager
        self.archive_dir = self.import_manager.importsdir / "goods_to_be_departed"
        
    def generatecalloffwaterfall(self, partnumber: str, daysback: int = 90) -> List[Dict]:
        if not self.archive_dir.exists():
            return []
        
        waterfalldata = []
        today = datetime.now().date()
        
        for daysbackoffset in range(daysback - 1, -1, -1):
            archivedate = today - timedelta(days=daysbackoffset)
            
            if archivedate.weekday() >= 5 and daysbackoffset > 7:
                continue
            
            gtbddata = self.loadgtbdarchive(archivedate)
            
            if gtbddata is None:
                continue
            
            calloffs = self.processpartcalloffs(gtbddata, partnumber)
            
            waterfalldata.append({
                'archive_date': archivedate,
                'call_offs': calloffs,
                'file_found': True})
            
        return waterfalldata
    
    def loadgtbdarchive(self, date: date) -> Optional[pd.DataFrame]: 
        dateprefix = f"GTBD {date.month}.{date.day}.{date.strftime('%y')}"
        
        matching = []
        for ext in ['.csv', '.xlsx', '.xls', '.xlsm']:
            matching.extend(self.archive_dir.glob(f"*{dateprefix}*{ext}"))
            
        if not matching:
            return None
        
        filepath = max(matching, key=lambda f: f.stat().st_mtime)
        
        try:
            if filepath.suffix.lower() == '.csv':
                gtbddf = self.import_manager.loaddata('goods_to_be_departed', filepath.name)
            else:
                gtbddf = pd.read_excel(filepath)
            
            if gtbddf.empty:
                return None
            
            requiredcols = ['PART_NO', 'CALL_OFF_QUANTITY', 'EARLIEST_SHIPING_TIME']
            if any(col not in gtbddf.columns for col in requiredcols):
                return None
            
            return gtbddf
        
        except:
            return None
        
    def processpartcalloffs(self, gtbddf: pd.DataFrame, partnumber: str) -> Dict:
        partdata = gtbddf[gtbddf['PART_NO'].astype(str).str.strip().str.upper() == partnumber.upper()]
        
        if partdata.empty:
            return {}
        
        calloffs = {}
        
        for _, row in partdata.iterrows():
            try:
                if pd.isna(row['EARLIEST_SHIPING_TIME']):
                    continue
                
                shippingdate = pd.to_datetime(row['EARLIEST_SHIPING_TIME']).date()
                quantity = float(row['CALL_OFF_QUANTITY']) if pd.notna(row['CALL_OFF_QUANTITY']) else 0
                
                if quantity <= 0:
                    continue
                
                if shippingdate not in calloffs:
                    calloffs[shippingdate] = 0
                calloffs[shippingdate] += quantity
                
            except:
                continue
        
        return calloffs
    
    def generateshippingdaterange(self, startdate: datetime.date, daysforward: int = 90) -> List[date]:
        shippingdates = []
        
        for daysoffset in range(daysforward):
            futuredate = startdate + timedelta(days=daysoffset)
            shippingdates.append(futuredate)
            
        return shippingdates
    
    def validatearchiveavailability(self) -> Tuple[bool, str, List[date]]:
        if not self.archive_dir.exists():
            return False, f"Archive directory not found: {self.archive_dir}", []
        
        today = datetime.now().date()
        availabledates = []
        
        for daysback in range(30):
            checkdate = today - timedelta(days=daysback)
            dateprefix = f"GTBD {checkdate.month}.{checkdate.day}.{checkdate.strftime('%y')}"
            matches = list(self.archive_dir.glob(f"*{dateprefix}*"))
            if matches:
                availabledates.append(checkdate)
                
        if not availabledates:
            return False, "No GTBD archive files found in the last 30 days", []
        
        return True, f"Found {len(availabledates)} GTBD archive files", availabledates
    
    def getpartsavailability(self, partnumber: str) -> Dict[str, int]:
        waterfalldata = self.generatecalloffwaterfall(partnumber, daysback=30)
        
        summary = {
            'fileschecked': len(waterfalldata),
            'fileswithdata': sum(1 for row in waterfalldata if row['call_offs']),
            'totalcalloffs': 0,
            'uniqueshippingdates': set()}
        
        for row in waterfalldata:
            for shippingdate, quantity in row['call_offs'].items():
                summary['totalcalloffs'] += quantity
                summary['uniqueshippingdates'].add(shippingdate)
                
        summary['uniqueshippingdates'] = len(summary['uniqueshippingdates'])
        
        return summary
