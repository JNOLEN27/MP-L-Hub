import pandas as pd
import json
from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Optional

class CoverageAnalysisEngine:
    def __init__(self, import_manager):
        self.import_manager = import_manager
    
    def loadrequireddata(self) -> Tuple[bool, str, Dict[str, pd.DataFrame]]:
        try:
            data = {
                'current_inventory': self.import_manager.loaddata("current_inventory_report"),
                'master_data': self.import_manager.loaddata("master_data"),
                'req_split_1': self.import_manager.loaddata("part_requirement_split_1"),
                'req_split_2': self.import_manager.loaddata("part_requirement_split_2"),
                'req_split_3': self.import_manager.loaddata("part_requirement_split_3"),
                'gtbr_data': self.import_manager.loaddata("goods_to_be_received")
            }
            
            if data['current_inventory'].empty or data['master_data'].empty:
                return False, "Current Inventory Report and Master Data are required.", data
            
            return True, "Data loaded successfully", data
            
        except Exception as e:
            return False, f"Error loading data: {str(e)}", {}
    
    def getpartswithconsumption(self, req1: pd.DataFrame, req2: pd.DataFrame, req3: pd.DataFrame) -> List[str]:
        allparts = set()
        
        for df in [req1, req2, req3]:
            if not df.empty and 'ARTNR' in df.columns:
                allparts.update(df['ARTNR'].dropna().unique())
        
        return sorted(list(allparts))
    
    def sortbydaystozerofriendly(self, coveragedf: pd.DataFrame) -> pd.DataFrame:
        if coveragedf.empty:
            return coveragedf
        
        datecols = []
        for col in coveragedf.columns:
            try:
                datetime.strptime(col, '%m/%d')
                datecols.append(col)
            except:
                continue
        
        datecols.sort(key=lambda x: datetime.strptime(x, '%m/%d'))
        
        def findzeroday(row):
            for i, col in enumerate(datecols):
                if pd.notna(row[col]) and row[col] <= 0:
                    return i
            return 999
        
        coveragedf['Days To Zero'] = coveragedf.apply(findzeroday, axis=1)
        coveragedf = coveragedf.sort_values(['Days To Zero', 'Part Number']).reset_index(drop=True)
        coveragedf = coveragedf.drop('Days To Zero', axis=1)
        return coveragedf
    
    def loadcoveragecomments(self) -> Dict[str, str]:
        commentsfile = self.import_manager.importsdir.parent / "coveragecomments.json"
        
        print(f"DEBUG: Loading comments from: {commentsfile}")
        print(f"DEBUG: File exists: {commentsfile.exists()}")
        
        if commentsfile.exists():
            try:
                with open(commentsfile, 'r') as f:
                    comments = json.load(f)
                    print(f"DEBUG: Loaded {len(comments)} comments: {list(comments.keys())[:5]}")
                    return comments
            except Exception as e:
                print(f"Error loading comments: {e}")
        
        print("DEBUG: Returning empty comments dict")
        return {}
    
    def savecoveragecomments(self, comments: Dict[str, str]):
        commentsfile = self.import_manager.importsdir.parent / "coveragecomments.json"
        
        print(f"DEBUG: Saving {len(comments)} comments to: {commentsfile}")
        print(f"DEBUG: Parent directory: {commentsfile.parent}")
        print(f"DEBUG: Parent directory exists: {commentsfile.parent.exists()}")
        
        try:
            commentsfile.parent.mkdir(parents=True, exist_ok=True)
            
            with open(commentsfile, 'w') as f:
                json.dump(comments, f, indent=2)
            
            print(f"DEBUG: File saved successfully. File now exists: {commentsfile.exists()}")
            
            if commentsfile.exists():
                file_size = commentsfile.stat().st_size
                print(f"DEBUG: File size: {file_size} bytes")
        
        except Exception as e:
            print(f"Error saving comments: {e}")
            import traceback
            traceback.print_exc()
            
    def addcoveragecomments(self, coveragedf: pd.DataFrame) -> pd.DataFrame:
        comments = self.loadcoveragecomments()
        coveragedf['Comments'] = coveragedf['PART_NO'].astype(str).map(comments).fillna('')
        return coveragedf
    
    def adddaysuntilzerocolumn(self, coveragedf: pd.DataFrame) -> pd.DataFrame:
        if coveragedf.empty:
            return coveragedf
        
        dailycols = [col for col in coveragedf.columns if col.startswith('Day_')]
        dailycols.sort()
        
        def calculatedaystozero(row):
            for i, col in enumerate(dailycols):
                try:
                    value = float(row[col]) if pd.notna(row[col]) else float('inf')
                    if value <= 0:
                        return i
                except (ValueError, TypeError):
                    continue
            return 999
        
        coveragedf['Day Alert'] = coveragedf.apply(calculatedaystozero, axis=1)
        
        def ensureintegeronly(days):
            try:
                daysint = int(float(days))
                return min(daysint, 999)
            except (ValueError, TypeError):
                return 999
        
        coveragedf['Day Alert'] = coveragedf['Day Alert'].apply(ensureintegeronly)
        coveragedf['Day Alert'] = coveragedf['Day Alert'].astype(int)
        
        return coveragedf
        
    def renamecolumnstofriendly(self, coveragedf: pd.DataFrame) -> pd.DataFrame:
        columnmapping = {
            'PART_NO': 'Part Number',
            'PART_DESC': 'Part Description',
            'PRICE': 'Price',
            'UNIT_LOAD_QTY': 'Unit Load Qty',
            'SUPP_MFG': 'MFG Code',
            'SUPP_NAME': 'Supplier Name',
            'SUPP_SHP': 'SHP Code',
            'SUPP_SHP_COUNTRY': 'SHP Country',
            'SCC_NAME': 'SCC Name',
            'Region': 'Region',
            'Comments': 'Comments'}
        
        dailycols = [col for col in coveragedf.columns if col.startswith('Day_')]
        for col in dailycols:
            try:
                parts = col.split('_')
                if len(parts) >= 4:
                    datestr = parts[2] + '-' + parts[3] + '-' + parts[4]
                    dateobj = datetime.strptime(datestr, '%Y-%m-%d')
                    newname = dateobj.strftime('%m/%d')
                    columnmapping[col] = newname
            except:
                if len(parts) > 1:
                    daynum = parts[1]
                    columnmapping[col] = f"Day {int(daynum)}"
                else:
                    columnmapping[col] = col
       
        coveragedf = coveragedf.rename(columns=columnmapping)
        return coveragedf
    
    def reordercolumns(self, coveragedf: pd.DataFrame) -> pd.DataFrame:
        if coveragedf.empty:
            return coveragedf
        
        preferredorder = ['Part Number', 'Part Description', 'MFG Code', 'Supplier Name', 'SHP Code', 'SHP Country', 'Region', 'Unit Load Qty', 'Price', 'SCC Name', 'Day Alert', 'Comments']
        
        datecolumns = []
        for col in coveragedf.columns:
            try:
                datetime.strptime(col, '%m/%d')
                datecolumns.append(col)
            except:
                continue
            
        datecolumns.sort(key=lambda x: datetime.strptime(x, '%m/%d'))
        finalorder = []
        
        for col in preferredorder:
            if col in coveragedf.columns:
                finalorder.append(col)
        
        finalorder.extend(datecolumns)
        remainingcols = [col for col in coveragedf.columns if col not in finalorder]
        finalorder.extend(remainingcols)
        
        coveragedf = coveragedf[finalorder]
        
        return coveragedf
    
    def buildcoverageanalysis(self, datadict: Dict[str, pd.DataFrame], daysforward: int = 90) -> pd.DataFrame:
        uniqueparts = self.getpartswithconsumption(
            datadict['req_split_1'], 
            datadict['req_split_2'], 
            datadict['req_split_3']
        )
        
        if not uniqueparts:
            return pd.DataFrame()
        
        coveragedf = self.buildbasetable(uniqueparts, datadict)
        coveragedf = self.addregioncolumn(coveragedf)
        coveragedf = self.addcoveragecomments(coveragedf)
        coveragedf = self.adddailyprojections(coveragedf, datadict, daysforward)
        coveragedf = self.adddaysuntilzerocolumn(coveragedf)
        coveragedf = self.renamecolumnstofriendly(coveragedf)
        coveragedf = self.reordercolumns(coveragedf)
        coveragedf = self.sortbydaystozerofriendly(coveragedf)
        
        return coveragedf
    
    def buildbasetable(self, uniqueparts: List[str], datadict: Dict[str, pd.DataFrame]) -> pd.DataFrame:
        coveragedf = pd.DataFrame({'PART_NO': uniqueparts})
        coveragedf = self.mergemasterdata(coveragedf, datadict['master_data'])
        coveragedf = self.addinitialstock(coveragedf, datadict['current_inventory'])
        coveragedf = coveragedf[coveragedf['Initial_Stock'] > 0].reset_index(drop=True)
        
        ldjisshp = ['AEWMV', 'AEJPM', 'AEXXW', 'AEIK0', 'AE2YR', 'AEX93', 'AE9ET']
        if 'SUPP_SHP' in coveragedf.columns:
            coveragedf = coveragedf[~coveragedf['SUPP_SHP'].isin(ldjisshp)].reset_index(drop=True)
        
        return coveragedf
    
    def mergemasterdata(self, coveragedf: pd.DataFrame, masterdata: pd.DataFrame) -> pd.DataFrame:
        if masterdata.empty:
            return coveragedf
        
        mastercols = ['PART', 'PART_DESC', 'PRICE', 'UNIT_LOAD_QTY', 'SAFETY', 'STOCK', 'MUL', 'SUPP_MFG', 'SUPP_NAME', 'SUPP_SHP', 'SUPP_SHP_COUNTRY', 'SCC_NAME']
        availablecols = [col for col in mastercols if col in masterdata.columns]
    
        part_col = None
        for possible_part_col in ['PART', 'PART_NO', 'PART_NUMBER', 'ARTNR']:
            if possible_part_col in masterdata.columns:
                part_col = possible_part_col
                break
        
        if availablecols and part_col:
            merge_cols = [part_col] + [col for col in availablecols if col != part_col]
            mastersubset = masterdata[merge_cols].drop_duplicates(subset=[part_col])
    
            coveragedf = coveragedf.merge(
                mastersubset, 
                left_on='PART_NO', 
                right_on=part_col, 
                how='left'
            )

            if part_col in coveragedf.columns and part_col != 'PART_NO':
                coveragedf = coveragedf.drop(part_col, axis=1)
        
        return coveragedf
    
    def determineregion(self, country):
        if pd.isna(country) or not country:
            return "No Country Found"
        
        country = str(country).upper().strip()
        
        if country in ['USA']:
            return 'USA'
        
        if country in ['MEXICO', 'CANADA']:
            return 'MEXI'
        
        if country in ['AUSTRIA', 'BELGIUM', 'BULGARIA', 'CZECH REPUBLIC', 'DENMARK', 'FRANCE', 'GERMANY', 'HUNGARY', 'IRELAND', 'ITALY', 'LITHUANIA',
                       'MOROCCO', 'NETHERLANDS', 'NORWAY', 'POLAND', 'PORTUGAL', 'ROMANIA', 'SLOVAK REPUBLIC', 'SLOVENIA', 'SPAIN', 'SWEDEN',
                       'SWITZERLAND', 'TUNISIA', 'TURKEY', 'UKRAINE', 'UNITED KINGDOM']:
            return 'EMEA'
        
        if country in ['CHINA', 'SOUTH KOREA', 'THAILAND', 'VIETNAM']:
            return 'APAC'
        
        return 'Country Not Mapped. Reach out to Admin'
    
    def addregioncolumn(self, coveragedf: pd.DataFrame) -> pd.DataFrame:
        shipcountrycol = None
        possiblecols = ['SUPP_SHP_COUNTRY']
        
        for col in possiblecols:
            if col in coveragedf.columns:
                shipcountrycol = col
                break
        
        if shipcountrycol:
            coveragedf['Region'] = coveragedf[shipcountrycol].apply(self.determineregion)
        else:
            coveragedf['Region'] = 'Unknown'
        return coveragedf
    
    def addinitialstock(self, coveragedf: pd.DataFrame, currentinventory: pd.DataFrame) -> pd.DataFrame:
        if currentinventory.empty:
            coveragedf['Initial_Stock'] = 0
            return coveragedf
        
        inventorycols = ['PART_NO', 'BEGINNING_INVENTORY_TODAY', 'INVENTORY_YARD_TODAY']
        availablecols = [col for col in inventorycols if col in currentinventory.columns]
        
        inventory_value_cols = [col for col in availablecols if col != 'PART_NO']
        
        if len(availablecols) >= 2 and len(inventory_value_cols) > 0:
            agg_dict = {col: 'sum' for col in inventory_value_cols}
            
            inventorysubset = currentinventory[availablecols].groupby('PART_NO').agg(agg_dict).reset_index()
            
            inventorysubset['Initial_Stock'] = (
                inventorysubset.get('BEGINNING_INVENTORY_TODAY', 0) + 
                inventorysubset.get('INVENTORY_YARD_TODAY', 0)
            )
            
            coveragedf = coveragedf.merge(
                inventorysubset[['PART_NO', 'Initial_Stock']], 
                on='PART_NO', 
                how='left'
            )
        else:
            print("DEBUG: Not enough inventory columns found, setting Initial_Stock to 0")
            coveragedf['Initial_Stock'] = 0
        
        coveragedf['Initial_Stock'] = coveragedf['Initial_Stock'].fillna(0)
        return coveragedf
    
    def adddailyprojections(self, coveragedf: pd.DataFrame, datadict: Dict[str, pd.DataFrame], daysforward: int) -> pd.DataFrame:
        consumptiondata = self.combineconsumptiondata(
            datadict['req_split_1'],
            datadict['req_split_2'], 
            datadict['req_split_3']
        )
        
        coveragedf['Initial_Stock'] = pd.to_numeric(coveragedf['Initial_Stock'], errors='coerce').fillna(0)
        
        today = datetime.now().date()
        
        for dayoffset in range(daysforward):
            date = today + timedelta(days=dayoffset)
            datestr = date.strftime('%Y_%m_%d')
            colname = f'Day_{dayoffset:03d}_{datestr}'
            
            if dayoffset == 0:
                coveragedf[colname] = coveragedf['Initial_Stock']
            else:
                # Initialize with zeros first
                coveragedf[colname] = 0.0
        
        for dayoffset in range(1, daysforward):
            date = today + timedelta(days=dayoffset)
            datestr = date.strftime('%Y_%m_%d')
            colname = f'Day_{dayoffset:03d}_{datestr}'
            
            prevdate = date - timedelta(days=1)
            prevdatestr = prevdate.strftime('%Y_%m_%d')
            prevcol = f'Day_{dayoffset-1:03d}_{prevdatestr}'
            
            dailyconsumption = consumptiondata.get(date, pd.Series(dtype=float))
            
            def safe_calculate(row):
                try:
                    if prevcol not in row.index:
                        return 0
                    
                    prev_value = float(row[prevcol]) if pd.notna(row[prevcol]) else 0
                    
                    part_no = row['PART_NO']
                    consumption = float(dailyconsumption.get(part_no, 0)) if part_no in dailyconsumption else 0
                    
                    remaining = max(0, prev_value - consumption)
                    return remaining
                    
                except:
                    return 0
            
            try:
                coveragedf[colname] = coveragedf.apply(safe_calculate, axis=1)
                
            except:
                coveragedf[colname] = 0
        
        if 'Initial_Stock' in coveragedf.columns:
            coveragedf = coveragedf.drop('Initial_Stock', axis=1)
        
        return coveragedf
    
    def analyzeindivpart(self, partnumber: str, datadict: Dict[str, pd.DataFrame]) -> Tuple[Dict, List[Dict]]:
        partinfo = self.findpartinfo(partnumber, datadict)
        if not partinfo:
            return None, []
        
        transactions = self.generateparttransactions(partnumber, datadict, partinfo['Initial Stock'])
        return partinfo, transactions
    
    def findpartinfo(self, partnumber: str, datadict: Dict[str, pd.DataFrame]) -> Dict:
        masterdata = datadict['master_data']
        currentinventory = datadict['current_inventory']
        
        partcol = None
        for possiblecol in ['PART', 'PART_NO', 'PART_NUMBER', 'ARTNR']:
            if possiblecol in masterdata.columns:
                partcol = possiblecol
                break
            
        if not partcol:
            return None
        
        partrow = masterdata[masterdata[partcol].astype(str).str.upper() == partnumber]
        if partrow.empty:
            return None
        
        partrow = partrow.iloc[0]
        initialstock = 0
        if not currentinventory.empty and 'PART_NO' in currentinventory.columns:
            invrow = currentinventory[currentinventory['PART_NO'].astype(str).str.upper() == partnumber]
            if not invrow.empty:
                beginv = invrow['BEGINNING_INVENTORY_TODAY'].fillna(0).sum() 
                yardinv = invrow['INVENTORY_YARD_TODAY'].fillna(0).sum()
                initialstock = int(beginv + yardinv)
                
        partinfo = {
            'Part Number': partnumber,
            'Part Description': partrow.get('PART_DESC', ''),
            'Supplier Name': partrow.get('SUPP_NAME', ''),
            'MFG Code': partrow.get('SUPP_MFG', ''),
            'SHP Code': partrow.get('SUPP_SHP', ''),
            'Initial Stock': int(initialstock),
            'Unit Load Qty': partrow.get('UNIT_LOAD_QTY', 0),
            'Multi Unit Load': partrow.get('MUL', 0),
            'Piece Price': partrow.get('PRICE', 0),
            'Safety Days': partrow.get('SAFETY', 0),
            'Safety Stock': partrow.get('STOCK', 0)}
        
        return partinfo
        
    def generateparttransactions(self, partnumber: str, datadict: Dict[str, pd.DataFrame], initialstock: int) -> List[Dict]:
        transactions = []
        today = datetime.now().date()
        
        transactions.append({
            'Date': today.strftime('%m/%d/%Y'),
            'Transaction Type': 'Stock',
            'Receipt/Reqmt': '',
            'Available QTY': initialstock,
            'ASN': '',
            '_running_qty': initialstock
        })
        
        consumptiondata = self.buildpartconsumption(partnumber, datadict)
        receiptdata = self.buildpartreceipts(partnumber, datadict)
        currentqty = initialstock
        
        for dayoffset in range(1, 91):
            date = today + timedelta(days=dayoffset)
            datestr = date.strftime('%m/%d/%Y')
            
            if date in receiptdata:
                for receipt in receiptdata[date]:
                    currentqty += receipt['quantity']
                    transactions.append({
                        'Date': datestr,
                        'Transaction Type': 'GR',
                        'Receipt/Reqmt': f"+{receipt['quantity']:,}",
                        'Available QTY': currentqty,
                        'ASN': '',  # Will add ASN logic later
                        '_running_qty': currentqty
                    })

            if date in consumptiondata:
                consumptionqty = int(consumptiondata[date])
                if consumptionqty > 0:
                    currentqty = max(0, currentqty - consumptionqty)
                    transactions.append({
                        'Date': datestr,
                        'Transaction Type': 'Req',
                        'Receipt/Reqmt': f"-{consumptionqty:,}",
                        'Available QTY': currentqty,
                        'ASN': '',
                        '_running_qty': currentqty
                    })

        return transactions
    
    def getinitialstock(self, partnumber, currentinventory):
        if currentinventory.empty or 'PART_NO' not in currentinventory.columns:
            return 0
        
        partrows = currentinventory[currentinventory['PART_NO'].astype(str).str.upper() == partnumber.upper()]
        if partrows.empty:
            return 0
        
        totalstock = 0
        if 'BEGINNING_INVENTORY_TODAY' in partrows.columns:
            totalstock += partrows['BEGINNING_INVENTORY_TODAY'].fillna(0).sum()
        if 'INVENTORY_YARD_TODAY' in partrows.columns:
            totalstock += partrows['INVENTORY_YARD_TODAY'].fillna(0).sum()
            
        return int(totalstock)
        
    def buildpartconsumption(self, partnumber, datadict):
        consumptionbydate = {}
        
        for reqkey in ['req_split_1', 'req_split_2', 'req_split_3']:
            reqdf = datadict[reqkey]
            
            if reqdf.empty or 'ARTNR' not in reqdf.columns:
                continue
            
            partreqs = reqdf[reqdf['ARTNR'].astype(str).str.upper() == partnumber.upper()]
            
            if partreqs.empty:
                continue
            
            datecol = self.findcolumn(partreqs, ['PRODDAG', 'DATE', 'PROD_DATE'])
            qtycol = self.findcolumn(partreqs, ['ARTAN', 'QTY', 'QUANTITY'])
            
            if not datecol or not qtycol:
                continue
            
            for _, row in partreqs.iterrows():
                try:
                    if pd.isna(row[datecol]):
                        continue
                    
                    date = pd.to_datetime(row[datecol]).date()
                    quantity = float(row[qtycol]) if not pd.isna(row[qtycol]) else 0
                    
                    if date not in consumptionbydate:
                        consumptionbydate[date] = 0
                        
                    consumptionbydate[date] += quantity
                    
                except:
                    continue
        
        return consumptionbydate
    
    def buildpartreceipts(self, partnumber, datadict):
        receiptbydate = {}
        gtbrdf = datadict['gtbr_data']
        
        if gtbrdf.empty:
            return receiptbydate
        
        partcol = self.findcolumn(gtbrdf, ['PART_NO', 'PART', 'ARTNR', 'PART_NUMBER'])
        if not partcol:
            return receiptbydate
        
        partreceipts = gtbrdf[gtbrdf[partcol].astype(str).str.upper() == partnumber.upper()]
        
        if partreceipts.empty:
            print(f"DEBUG: No GTBR records found for part {partnumber}")
            return receiptbydate
        
        datecol = self.findcolumn(partreceipts, [
            'DELIVERY_DATE', 'RECEIPT_DATE', 'DATE', 'GR_DATE', 'PROD_DATE',
            'ANK_TID_TIDIGAST', 'ANK_TID_SENAST',
            'FRAKT_TID_TIDIGAST', 'FRAKT_TID_SENAST'
        ])
        
        qtycol = self.findcolumn(partreceipts, ['QTY', 'QUANTITY', 'RECEIPT_QTY', 'ARTAN'])
        
        if not datecol or not qtycol:
            return receiptbydate
        
        for _, row in partreceipts.iterrows():
            try:
                if pd.isna(row[datecol]):
                    continue
                
                date = pd.to_datetime(row[datecol]).date()
                quantity = float(row[qtycol]) if not pd.isna(row[qtycol]) else 0
                
                if quantity <= 0:
                    continue
                
                if date not in receiptbydate:
                    receiptbydate[date] = []
                    
                receiptbydate[date].append({
                    'quantity': int(quantity),
                    'asn': row.get('ASN', ''),
                    'po': row.get('PO', '')
                })
                
            except:
                continue
        
        return receiptbydate
            
    def findcolumn(self, df, possiblenames):
        for name in possiblenames:
            if name in df.columns:
                return name
        return None
    
    def combineconsumptiondata(self, req1: pd.DataFrame, req2: pd.DataFrame, req3: pd.DataFrame) -> Dict:
        consumptionbydate = {}
        total_rows_processed = 0
        
        for i, reqdf in enumerate([req1, req2, req3], 1):
            if reqdf.empty or 'ARTNR' not in reqdf.columns:
                continue
            
            datecol = self.findcolumn(reqdf, ['PRODDAG', 'DATE', 'PROD_DATE'])
            qtycol = self.findcolumn(reqdf, ['ARTAN', 'QTY', 'QUANTITY'])
            
            if not datecol or not qtycol:
                continue
            
            rows_processed = 0
            
            for _, row in reqdf.iterrows():
                try:
                    if pd.isna(row[datecol]):
                        continue
                    
                    date = pd.to_datetime(row[datecol]).date()
                    partno = row['ARTNR']
                    quantity = float(row[qtycol]) if not pd.isna(row[qtycol]) else 0
                    
                    if date not in consumptionbydate:
                        consumptionbydate[date] = {}
                    
                    consumptionbydate[date][partno] = consumptionbydate[date].get(partno, 0) + quantity
                    rows_processed += 1
                    
                except:
                    continue
        
            total_rows_processed += rows_processed
        
        for date in consumptionbydate:
            consumptionbydate[date] = pd.Series(consumptionbydate[date])
        
        return consumptionbydate
    
    def exporttocsv(self, coveragedf: pd.DataFrame, filename: str) -> bool:
        try:
            coveragedf.to_csv(filename, index=False)
            return True
        except Exception as e:
            print(f"Export error: {e}")
            return False