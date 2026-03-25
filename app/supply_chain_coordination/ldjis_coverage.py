import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import Dict, List, Tuple


class LDJISCoverageEngine:
    N_FIXED = 6
    FIXED_COLS = [
        'Part Number',
        'Part Description',
        'SHP Code',
        'Moves away from COL',
        'Last Covered Mix',
        'Starting COL Mix',
    ]

    def __init__(self, import_manager):
        self.import_manager = import_manager

    def generateworkingdays(self, start_date, n_days: int) -> List[datetime]:
        days = []
        current = start_date
        while len(days) < n_days:
            if current.weekday() < 5:
                days.append(current)
            current += timedelta(days=1)
        return days

    def buildldjiscoveragedata(self, datadict: Dict) -> Tuple[bool, str, pd.DataFrame]:
        try:
            inventory = datadict.get('current_inventory', pd.DataFrame())
            master = datadict.get('master_data', pd.DataFrame())

            if inventory.empty or master.empty:
                return False, "Current Inventory and Master Data are required.", pd.DataFrame()

            ldjis_suppliers = ['LDJIS', 'LDJ-IS', 'LDJIS-VENDOR']

            if 'SUPPLIER_NAME' in master.columns:
                ldjis_mask = master['SUPPLIER_NAME'].str.contains(
                    'LDJIS|LDJ-IS', case=False, na=False
                )
                ldjis_parts = master[ldjis_mask]
            else:
                return False, "SUPPLIER_NAME column not found in master data.", pd.DataFrame()

            if ldjis_parts.empty:
                return False, "No LDJIS supplier parts found in master data.", pd.DataFrame()

            part_col = 'ARTNR' if 'ARTNR' in ldjis_parts.columns else ldjis_parts.columns[0]
            desc_col = 'PART_DESCRIPTION' if 'PART_DESCRIPTION' in ldjis_parts.columns else None
            shp_col = 'SHP_CODE' if 'SHP_CODE' in ldjis_parts.columns else None

            records = []
            for _, row in ldjis_parts.iterrows():
                part = str(row.get(part_col, ''))
                desc = str(row.get(desc_col, '')) if desc_col else ''
                shp = str(row.get(shp_col, '')) if shp_col else ''

                inv_row = pd.DataFrame()
                if 'ARTNR' in inventory.columns:
                    inv_row = inventory[inventory['ARTNR'] == part]

                current_stock = 0
                if not inv_row.empty:
                    stock_col = next(
                        (c for c in ['CURRENT_STOCK', 'STOCK', 'QUANTITY'] if c in inv_row.columns),
                        None
                    )
                    if stock_col:
                        current_stock = float(inv_row.iloc[0][stock_col])

                records.append({
                    'Part Number': part,
                    'Part Description': desc,
                    'SHP Code': shp,
                    'Moves away from COL': 0,
                    'Last Covered Mix': int(current_stock),
                    'Starting COL Mix': int(current_stock),
                })

            if not records:
                return False, "No LDJIS parts could be matched to inventory.", pd.DataFrame()

            df = pd.DataFrame(records)
            return True, "LDJIS coverage data built successfully.", df

        except Exception as e:
            return False, f"Error building LDJIS data: {str(e)}", pd.DataFrame()
