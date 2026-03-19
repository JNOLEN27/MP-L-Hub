import numpy as np
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
                'splunk_data': self.import_manager.loaddata("splunk_receiving_data"),
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
 
        sort_cols = [c for c in ['Day Alert', 'Part Number'] if c in coveragedf.columns]
        if sort_cols:
            coveragedf = coveragedf.sort_values(sort_cols).reset_index(drop=True)
        return coveragedf
 
    def loadcoveragecomments(self) -> Dict[str, str]:
        commentsfile = self.import_manager.importsdir.parent / "coveragecomments.json"
        if commentsfile.exists():
            try:
                with open(commentsfile, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading comments: {e}")
        return {}
 
    def savecoveragecomments(self, comments: Dict[str, str]):
        commentsfile = self.import_manager.importsdir.parent / "coveragecomments.json"
        try:
            commentsfile.parent.mkdir(parents=True, exist_ok=True)
            with open(commentsfile, 'w') as f:
                json.dump(comments, f, indent=2)
        except Exception as e:
            print(f"Error saving comments: {e}")
 
    def addcoveragecomments(self, coveragedf: pd.DataFrame) -> pd.DataFrame:
        comments = self.loadcoveragecomments()
        coveragedf['Comments'] = coveragedf['PART_NO'].astype(str).map(comments).fillna('')
        return coveragedf
 
    def adddaysuntilzerocolumn(self, coveragedf: pd.DataFrame) -> pd.DataFrame:
        if coveragedf.empty:
            return coveragedf
 
        dailycols = sorted([col for col in coveragedf.columns if col.startswith('Day_')])
        if not dailycols:
            coveragedf['Day Alert'] = 999
            return coveragedf
 
        arr = coveragedf[dailycols].values.astype(float)
        arr = np.where(np.isnan(arr), np.inf, arr)
 
        mask = arr <= 0
        has_zero = mask.any(axis=1)
        day_alert = np.where(has_zero, mask.argmax(axis=1), 999)
        coveragedf['Day Alert'] = np.minimum(day_alert, 999).astype(int)
 
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
            'Comments': 'Comments',
            'SAFETY': 'Safety Days',
            'STOCK': 'Safety Stock',
        }
 
        for col in [c for c in coveragedf.columns if c.startswith('Day_')]:
            try:
                parts = col.split('_')
                if len(parts) >= 5:
                    datestr = f"{parts[2]}-{parts[3]}-{parts[4]}"
                    columnmapping[col] = datetime.strptime(datestr, '%Y-%m-%d').strftime('%m/%d')
            except Exception:
                try:
                    columnmapping[col] = f"Day {int(parts[1])}" if len(parts) > 1 else col
                except Exception:
                    columnmapping[col] = col
 
        return coveragedf.rename(columns=columnmapping)
 
    def reordercolumns(self, coveragedf: pd.DataFrame) -> pd.DataFrame:
        if coveragedf.empty:
            return coveragedf
 
        preferredorder = [
            'Part Number', 'Part Description', 'MFG Code', 'Supplier Name',
            'SHP Code', 'SHP Country', 'Region', 'Safety Days', 'Safety Stock',
            'Unit Load Qty', 'Price', 'SCC Name', 'Day Alert', 'Comments',
        ]
 
        dateparsed: Dict[str, datetime] = {}
        for col in coveragedf.columns:
            try:
                dateparsed[col] = datetime.strptime(col, '%m/%d')
            except Exception:
                pass
 
        datecolumns = sorted(dateparsed, key=dateparsed.__getitem__)
 
        finalorder = [col for col in preferredorder if col in coveragedf.columns]
        finalorder.extend(datecolumns)
        placed = set(finalorder)
        finalorder.extend(col for col in coveragedf.columns if col not in placed)
 
        return coveragedf[finalorder]
 
    def buildcoverageanalysis(self, datadict: Dict[str, pd.DataFrame], daysforward: int = 40) -> pd.DataFrame:
        uniqueparts = self.getpartswithconsumption(
            datadict['req_split_1'],
            datadict['req_split_2'],
            datadict['req_split_3'],
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
 
        mastercols = [
            'PART', 'PART_DESC', 'PRICE', 'UNIT_LOAD_QTY', 'SAFETY', 'STOCK',
            'MUL', 'SUPP_MFG', 'SUPP_NAME', 'SUPP_SHP', 'SUPP_SHP_COUNTRY', 'SCC_NAME',
        ]
        availablecols = [col for col in mastercols if col in masterdata.columns]
 
        part_col = next(
            (c for c in ['PART', 'PART_NO', 'PART_NUMBER', 'ARTNR'] if c in masterdata.columns),
            None,
        )
 
        if availablecols and part_col:
            merge_cols = [part_col] + [col for col in availablecols if col != part_col]
            mastersubset = masterdata[merge_cols].drop_duplicates(subset=[part_col])
 
            coveragedf = coveragedf.merge(
                mastersubset,
                left_on='PART_NO',
                right_on=part_col,
                how='left',
            )
            if part_col in coveragedf.columns and part_col != 'PART_NO':
                coveragedf = coveragedf.drop(part_col, axis=1)
 
        return coveragedf
 
    def determineregion(self, country) -> str:
        if pd.isna(country) or not country:
            return "No Country Found"
 
        country = str(country).upper().strip()
 
        if country == 'USA':
            return 'USA'
        if country in ('MEXICO', 'CANADA'):
            return 'MEXI'
        if country in (
            'AUSTRIA', 'BELGIUM', 'BULGARIA', 'CZECH REPUBLIC', 'DENMARK', 'FRANCE',
            'GERMANY', 'HUNGARY', 'IRELAND', 'ITALY', 'LITHUANIA', 'MOROCCO',
            'NETHERLANDS', 'NORWAY', 'POLAND', 'PORTUGAL', 'ROMANIA', 'SLOVAK REPUBLIC',
            'SLOVENIA', 'SPAIN', 'SWEDEN', 'SWITZERLAND', 'TUNISIA', 'TURKEY',
            'UKRAINE', 'UNITED KINGDOM',
        ):
            return 'EMEA'
        if country in ('CHINA', 'SOUTH KOREA', 'THAILAND', 'VIETNAM'):
            return 'APAC'
 
        return 'Country Not Mapped. Reach out to Admin'
 
    def addregioncolumn(self, coveragedf: pd.DataFrame) -> pd.DataFrame:
        if 'SUPP_SHP_COUNTRY' not in coveragedf.columns:
            coveragedf['Region'] = 'Unknown'
            return coveragedf
 
        unique_countries = coveragedf['SUPP_SHP_COUNTRY'].unique()
        mapping = {c: self.determineregion(c) for c in unique_countries}
        coveragedf['Region'] = coveragedf['SUPP_SHP_COUNTRY'].map(mapping)
        return coveragedf
 
    def addinitialstock(self, coveragedf: pd.DataFrame, currentinventory: pd.DataFrame) -> pd.DataFrame:
        if currentinventory.empty:
            coveragedf['Initial_Stock'] = 0
            return coveragedf
 
        inventorycols = ['PART_NO', 'BEGINNING_INVENTORY_TODAY', 'INVENTORY_YARD_TODAY', 'INVENTORY_PORT_TODAY']
        availablecols = [col for col in inventorycols if col in currentinventory.columns]
        inventory_value_cols = [col for col in availablecols if col != 'PART_NO']
 
        if len(availablecols) >= 2 and inventory_value_cols:
            inventorysubset = currentinventory[availablecols].copy()
 
            for col in inventory_value_cols:
                inventorysubset[col] = pd.to_numeric(
                    inventorysubset[col].astype(str).str.replace(',', '', regex=False),
                    errors='coerce',
                ).fillna(0)
 
            aggdict = {col: 'sum' for col in inventory_value_cols}
            inventorysubset = inventorysubset.groupby('PART_NO').agg(aggdict).reset_index()
 
            inventorysubset['Initial_Stock'] = (
                inventorysubset.get('BEGINNING_INVENTORY_TODAY', 0)
                + inventorysubset.get('INVENTORY_YARD_TODAY', 0)
                + inventorysubset.get('INVENTORY_PORT_TODAY', 0)
            )
 
            coveragedf = coveragedf.merge(
                inventorysubset[['PART_NO', 'Initial_Stock']],
                on='PART_NO',
                how='left',
            )
        else:
            coveragedf['Initial_Stock'] = 0
 
        coveragedf['Initial_Stock'] = coveragedf['Initial_Stock'].fillna(0)
        return coveragedf
 
    def adddailyprojections(self, coveragedf: pd.DataFrame, datadict: Dict[str, pd.DataFrame], daysforward: int) -> pd.DataFrame:
        consumptiondata = self.combineconsumptiondata(
            datadict['req_split_1'],
            datadict['req_split_2'],
            datadict['req_split_3']
        )
        receiptdata = self.parsesplunkreceivingdata(datadict.get('splunk_data', pd.DataFrame()))
 
        coveragedf['Initial_Stock'] = pd.to_numeric(coveragedf['Initial_Stock'], errors='coerce').fillna(0)
 
        today = datetime.now().date()
 
        day0_receipts = receiptdata.get(today, pd.Series(dtype=float))
        day0_consumption = consumptiondata.get(today, pd.Series(dtype=float))
 
        def calc_day0(row):
            try:
                part_no = row['PART_NO']
                part_no_upper = str(part_no).upper().strip()
                receipts = float(day0_receipts.get(part_no_upper, 0)) if not day0_receipts.empty else 0
                consumption = float(day0_consumption.get(part_no, 0)) if not day0_consumption.empty else 0
                return float(row['Initial_Stock']) + receipts - consumption
            except:
                return float(row['Initial_Stock'])
 
        day0_col = f'Day_000_{today.strftime("%Y_%m_%d")}'
        coveragedf[day0_col] = coveragedf.apply(calc_day0, axis=1)
 
        last_col = day0_col

        for dayoffset in range(1, daysforward):
            date = today + timedelta(days=dayoffset)
        
            daily_receipts = receiptdata.get(date, pd.Series(dtype=float))
            daily_consumption = consumptiondata.get(date, pd.Series(dtype=float))
        
            if date.weekday() >= 5:
                continue
        
            datestr = date.strftime('%Y_%m_%d')
            colname = f'Day_{dayoffset:03d}_{datestr}'
            prevcol = last_col
 
            def safe_calculate(row, prevcol=prevcol, daily_receipts=daily_receipts, daily_consumption=daily_consumption):
                try:
                    if prevcol not in row.index:
                        return 0
                    prev_value = float(row[prevcol]) if pd.notna(row[prevcol]) else 0
                    part_no = row['PART_NO']
                    part_no_upper = str(part_no).upper().strip()
                    receipts = float(daily_receipts.get(part_no_upper, 0)) if not daily_receipts.empty else 0
                    consumption = float(daily_consumption.get(part_no, 0)) if not daily_consumption.empty else 0
                    return prev_value + receipts - consumption
                except:
                    return 0
 
            try:
                coveragedf[colname] = coveragedf.apply(safe_calculate, axis=1)
                last_col = colname
            except:
                coveragedf[colname] = 0
                last_col = colname
 
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
 
        partcol = next(
            (c for c in ['PART', 'PART_NO', 'PART_NUMBER', 'ARTNR'] if c in masterdata.columns),
            None,
        )
        if not partcol:
            return None
 
        partrow = masterdata[masterdata[partcol].astype(str).str.upper() == partnumber]
        if partrow.empty:
            return None
 
        partrow = partrow.iloc[0]
        initialstock = 0
        if not currentinventory.empty and 'PART_NO' in currentinventory.columns:
            invrow = currentinventory[
                currentinventory['PART_NO'].astype(str).str.upper() == partnumber
            ]
            if not invrow.empty:
                beginv = pd.to_numeric(
                    invrow['BEGINNING_INVENTORY_TODAY'].astype(str).str.replace(',', '', regex=False),
                    errors='coerce',
                ).fillna(0).sum()
                yardinv = pd.to_numeric(
                    invrow['INVENTORY_YARD_TODAY'].astype(str).str.replace(',', '', regex=False),
                    errors='coerce',
                ).fillna(0).sum()
                initialstock = int(beginv + yardinv)
 
        return {
            'Part Number': partnumber,
            'Part Description': partrow.get('PART_DESC', ''),
            'Supplier Name': partrow.get('SUPP_NAME', ''),
            'MFG Code': partrow.get('SUPP_MFG', ''),
            'SHP Code': partrow.get('SUPP_SHP', ''),
            'Initial Stock': int(initialstock),
            'Unit Load Qty': self.safefloat(partrow.get('UNIT_LOAD_QTY', 0)),
            'Multi Unit Load': self.safefloat(partrow.get('MUL', 0)),
            'Piece Price': self.safefloat(partrow.get('PRICE', 0)),
            'Safety Days': self.safefloat(partrow.get('SAFETY', 0)),
            'Safety Stock': self.safefloat(partrow.get('STOCK', 0)),
        }
 
    def generateparttransactions(self, partnumber: str, datadict: Dict[str, pd.DataFrame], initialstock: int) -> List[Dict]:
        transactions = []
        today = datetime.now().date()
 
        transactions.append({
            'Date': today.strftime('%m/%d/%Y'),
            'Transaction Type': 'Stock',
            'Receipt/Reqmt': '',
            'Available QTY': initialstock,
            'ASN': '',
            '_running_qty': initialstock,
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
                        'ASN': '',
                        '_running_qty': currentqty,
                    })
 
            if date in consumptiondata:
                consumptionqty = int(consumptiondata[date])
                if consumptionqty > 0:
                    currentqty = currentqty - consumptionqty
                    transactions.append({
                        'Date': datestr,
                        'Transaction Type': 'Req',
                        'Receipt/Reqmt': f"-{consumptionqty:,}",
                        'Available QTY': currentqty,
                        'ASN': '',
                        '_running_qty': currentqty,
                    })
 
        return transactions
 
    def getinitialstock(self, partnumber, currentinventory):
        if currentinventory.empty or 'PART_NO' not in currentinventory.columns:
            return 0
 
        partrows = currentinventory[
            currentinventory['PART_NO'].astype(str).str.upper() == partnumber.upper()
        ]
        if partrows.empty:
            return 0
 
        totalstock = 0
        if 'BEGINNING_INVENTORY_TODAY' in partrows.columns:
            totalstock += partrows['BEGINNING_INVENTORY_TODAY'].fillna(0).sum()
        if 'INVENTORY_YARD_TODAY' in partrows.columns:
            totalstock += partrows['INVENTORY_YARD_TODAY'].fillna(0).sum()
        if 'INVENTORY_PORT_TODAY' in partrows.columns:
            totalstock += partrows['INVENTORY_PORT_TODAY'].fillna(0).sum()
 
        return int(totalstock)
 
    def buildpartconsumption(self, partnumber, datadict):
        consumptionbydate = {}
        target = partnumber.upper()
 
        for reqkey in ['req_split_1', 'req_split_2', 'req_split_3']:
            reqdf = datadict[reqkey]
            if reqdf.empty or 'ARTNR' not in reqdf.columns:
                continue
 
            datecol = self.findcolumn(reqdf, ['PRODDAG', 'DATE', 'PROD_DATE'])
            qtycol = self.findcolumn(reqdf, ['ARTAN', 'QTY', 'QUANTITY'])
            if not datecol or not qtycol:
                continue
 
            try:
                partreqs = reqdf[reqdf['ARTNR'].astype(str).str.upper() == target]
                if partreqs.empty:
                    continue
 
                sub = pd.DataFrame({
                    'date': pd.to_datetime(partreqs[datecol], errors='coerce').dt.date,
                    'qty': pd.to_numeric(partreqs[qtycol], errors='coerce').fillna(0),
                })
                sub = sub.dropna(subset=['date'])
 
                for date, qty in sub.groupby('date')['qty'].sum().items():
                    consumptionbydate[date] = consumptionbydate.get(date, 0) + qty
 
            except Exception:
                continue
 
        return consumptionbydate
 
    def buildpartreceipts(self, partnumber, datadict):
        receiptbydate = {}
        splunkdf = datadict.get('splunk_data', pd.DataFrame())
 
        if splunkdf.empty or 'Part Number' not in splunkdf.columns:
            return receiptbydate
 
        target = partnumber.upper()
        partreceipts = splunkdf[
            splunkdf['Part Number'].astype(str).str.upper().str.strip() == target
        ]
        if partreceipts.empty:
            return receiptbydate
 
        date_col = 'Load Delivery Date Final'
        if date_col not in partreceipts.columns:
            return receiptbydate
 
        qty_col = self.findcolumn(splunkdf, [
            'Quantity', 'QTY', 'Qty', 'QUANTITY', 'ARTAN',
            'Load Quantity', 'TO Quantity', 'Receipt Qty', 'Receipt QTY',
            'Pieces', 'Units',
        ])
 
        try:
            def _shift(d):
                if not pd.notna(d): return d
                d = d + timedelta(days=1)
                if d.weekday() == 5: d += timedelta(days=2)
                elif d.weekday() == 6: d += timedelta(days=1)
                return d
            df = partreceipts.copy()
            df['_date'] = pd.to_datetime(df[date_col], errors='coerce').dt.date.apply(_shift)
            df['_qty'] = (pd.to_numeric(df[qty_col], errors='coerce').fillna(0) if qty_col else pd.Series(0, index=df.index))
            df = df[df['_qty'] > 0].dropna(subset=['_date'])
 
            for _, row in df.iterrows():
                receiptbydate.setdefault(row['_date'], []).append({
                    'quantity': int(row['_qty']),
                    'asn': str(row.get('TO Number', '')),
                    'po': '',
                })
        except Exception:
            pass
 
        return receiptbydate
 
    def findcolumn(self, df, possiblenames):
        for name in possiblenames:
            if name in df.columns:
                return name
        return None
 
    def safefloat(self, value, default=0.0):
        try:
            if pd.isna(value):
                return default
        except (TypeError, ValueError):
            pass
        try:
            return float(str(value).replace(',', ''))
        except (ValueError, TypeError):
            return default
 
    def parsesplunkreceivingdata(self, splunkdf: pd.DataFrame) -> Dict:
        if splunkdf.empty:
            return {}
 
        part_col = 'Part Number'
        date_col = 'Load Delivery Date Final'
 
        if part_col not in splunkdf.columns or date_col not in splunkdf.columns:
            print(f"[Splunk] Missing required columns. Available: {splunkdf.columns.tolist()}")
            return {}
 
        qty_col = self.findcolumn(splunkdf, [
            'Quantity', 'QTY', 'Qty', 'QUANTITY', 'ARTAN',
            'Load Quantity', 'TO Quantity', 'Receipt Qty', 'Receipt QTY',
            'Pieces', 'Units',
        ])
 
        if qty_col is None:
            print(f"[Splunk] WARNING: No quantity column found. Available columns: {splunkdf.columns.tolist()}")
        else:
            print(f"[Splunk] Using quantity column: '{qty_col}'")
 
        print(f"[Splunk] Raw dataframe shape: {splunkdf.shape}")
        print(f"[Splunk] Sample dates: {splunkdf[date_col].head(3).tolist()}")
        print(f"[Splunk] Sample parts: {splunkdf[part_col].head(3).tolist()}")
        if qty_col:
            print(f"[Splunk] Sample qtys: {splunkdf[qty_col].head(3).tolist()}")

        try:
            def shift_date(d):
                if not pd.notna(d):
                    return d
                d = d + timedelta(days=1)
                if d.weekday() == 5:
                    d += timedelta(days=2)
                elif d.weekday() == 6:
                    d += timedelta(days=1)
                return d

            df = splunkdf[[part_col, date_col]].copy()
            df['_part'] = df[part_col].astype(str).str.upper().str.strip()
            df['_date'] = pd.to_datetime(df[date_col], errors='coerce').dt.date.apply(shift_date)

            if qty_col:
                qty_series = splunkdf[qty_col].astype(str).str.replace(',', '', regex=False)
                df['_qty'] = pd.to_numeric(qty_series, errors='coerce').fillna(0)
            else:
                df['_qty'] = 0

            print(f"[Splunk] Rows before filter: {len(df)}, non-zero qty: {(df['_qty'] > 0).sum()}, valid dates: {df['_date'].notna().sum()}")

            df = df[df['_qty'] > 0].dropna(subset=['_date'])
            print(f"[Splunk] Rows after filter: {len(df)}")

            grouped = df.groupby(['_date', '_part'])['_qty'].sum()
            print(f"[Splunk] Unique receipt dates found: {len(grouped.index.get_level_values(0).unique())}")

            receiptbydate: Dict = {}
            for (date, partno), qty in grouped.items():
                receiptbydate.setdefault(date, {})[partno] = qty

            for date in receiptbydate:
                receiptbydate[date] = pd.Series(receiptbydate[date])

            print(f"[Splunk] First 5 receipt dates: {sorted(receiptbydate.keys())[:5]}")
            return receiptbydate

        except Exception as e:
            print(f"Error parsing Splunk receiving data: {e}")
            import traceback; traceback.print_exc()
            return {}
 
    def combineconsumptiondata(self, req1: pd.DataFrame, req2: pd.DataFrame, req3: pd.DataFrame) -> Dict:
        all_dfs = []
 
        for reqdf in [req1, req2, req3]:
            if reqdf.empty or 'ARTNR' not in reqdf.columns:
                continue
 
            datecol = self.findcolumn(reqdf, ['PRODDAG', 'DATE', 'PROD_DATE'])
            qtycol = self.findcolumn(reqdf, ['ARTAN', 'QTY', 'QUANTITY'])
            if not datecol or not qtycol:
                continue
 
            try:
                subset = reqdf[['ARTNR', datecol, qtycol]].copy()
                subset.columns = ['ARTNR', '_date_raw', '_qty']
                subset['_qty'] = pd.to_numeric(subset['_qty'], errors='coerce').fillna(0)
                subset['_date'] = pd.to_datetime(subset['_date_raw'], errors='coerce').dt.date
                subset = subset.dropna(subset=['_date'])
                all_dfs.append(subset[['ARTNR', '_date', '_qty']])
            except Exception:
                continue
 
        if not all_dfs:
            return {}
 
        combined = pd.concat(all_dfs, ignore_index=True)
        grouped = combined.groupby(['_date', 'ARTNR'])['_qty'].sum()
 
        consumptionbydate: Dict = {}
        for (date, partno), qty in grouped.items():
            consumptionbydate.setdefault(date, {})[partno] = qty

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
