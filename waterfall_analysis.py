import pandas as pd
import os
from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Optional
from pathlib import Path

class WaterfallAnalysisEngine:
    def __init__(self, import_manager):
        self.import_manager = import_manager
        self.archive_dir = self.import_manager.importsdir.parent / "historical_data"
        
    def generatecalloffwaterfall(self, partnumber: str, daysback: int = 90) -> List[Dict]:
        if not self.archive_dir.exists():
            return []
        
        waterfalldata = []
        today = datetime.now().date()
        
        for daysbackoffset in range(daysback):
            archivedate = today - timedelta(days=daysbackoffset)
            
            if archivedate.weekday() >= 5 and daysbackoffset > 7:
                continue
            
            gtbddata = self.loadgtbdarchive(archivedate)
            
            if gtbddata is None:
                waterfalldata.append({
                    'archive_date': archivedate,
                    'call_offs': {},
                    'file_found': False})
                continue
            
            calloffs = self.processpartcalloffs(gtbddata, partnumber)
            
            waterfalldata.append({
                'archive_date': archivedate,
                'call_offs': calloffs,
                'file_found': True})
            
        return waterfalldata
    
    def loadgtbdarchive(self, date: datetime.date) -> Optional[pd.DataFrame]: 
        datestr = date.strftime("%-m.%-d.%y") if os.name != 'nt' else date.strftime("%#m.%#d.%y")
        expectedfilename = f"GTBD {datestr}.csv"
        filepath = self.archive_dir / expectedfilename
       
        if not filepath.exists():
            return None
        try:
            gtbddf = pd.read_csv(filepath)
            requiredcols = ['PART_NO', 'CALL_OFF_QUANTITY', 'EARLIEST_SHIPPING_TIME']
            missingcols = [col for col in requiredcols if col not in gtbddf.columns]
            
            if missingcols:
                return None
            
            return gtbddf
        
        except:
            return None
        
    def processpartcalloffs(self, gtbddf: pd.DataFrame, partnumber: str) -> Dict:
        partdata = gtbddf[gtbddf['PART_NO'].astype(str).upper() == partnumber.upper()]
        
        if partdata.empty:
            return {}
        
        calloffs = {}
        
        for _, row in partdata.iterrows():
            try:
                if pd.isna(row['EARLIEST_SHIPPING_TIME']):
                    continue
                
                shippingdate = pd.to_datetime(row['EARLIEST_SHIPPING_TIME']).date()
                quantity = float(row['CALL_OFF_QUANTITY']) if pd.notna(row['CALL_OFF_QUANTITY']) else 0
                
                if quantity <= 0:
                    continue
                
                if shippingdate not in calloffs:
                    calloffs[shippingdate] = 0
                calloffs[shippingdate] += quantity
                
            except:
                continue
        
        return calloffs
    
    def generateshippingdaterange(self, startdate: datetime.date, daysforward: int = 90) -> List[datetime.date]:
        shippingdates = []
        
        for daysoffset in range(daysforward):
            futuredate = startdate + timedelta(days=daysoffset)
            shippingdates.append(futuredate)
            
        return shippingdates
    
    def validatearchiveavailability(self) -> Tuple[bool, str, List[datetime.date]]:
        if not self.archive_dir.exists():
            return False, f"Archive directory not found: {self.archive_dir}", []
        
        today = datetime.now().date()
        availabledates = []
        
        for daysback in range(30):
            checkdate = today - timedelta(days=daysback)
            datestr = checkdate.strftime("%-m.%-d.%y") if os.name != 'nt' else checkdate.strtime("%#m.%#d.%y")
            expectedfilename = f"GTBD {datestr}.csv"
            filepath = self.archive_dir / expectedfilename
            
            if filepath.exists():
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