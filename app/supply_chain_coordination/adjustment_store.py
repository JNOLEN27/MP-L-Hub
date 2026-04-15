import json
import uuid
from datetime import datetime
from app.utils.config import SHAREDNETWORKPATH

COLUMN_MAPPING_DEFAULTS = {
    "inv.part_no":   "PART_NO",
    "inv.beginning": "BEGINNING_INVENTORY_TODAY",
    "inv.yard":      "INVENTORY_YARD_TODAY",
    "inv.port":      "INVENTORY_PORT_TODAY",
    "md.unit_load":       "UNIT_LOAD_QTY",
    "md.safety_days":     "SAFETY",
    "md.safety_stock":    "STOCK",
    "md.multi_unit_load": "MUL",
    "req.part_no": "ARTNR",
    "req.date":    "PRODDAG",
    "req.qty":     "ARTAN",
    "splunk.part_no": "Part Number",
    "splunk.date":    "Load Delivery Date Final",
    "splunk.qty":     "Quantity",
    "gtr.part_no": "ARTNR",
    "gtr.date":    "ANK_TID_SENAST",
    "gtr.qty":     "ARTAN",
}

COLUMN_MAPPING_META = [
    ("Inventory Report",     "inv.part_no",    "Part Number column"),
    ("Inventory Report",     "inv.beginning",  "Beginning Inventory Today"),
    ("Inventory Report",     "inv.yard",       "Inventory Yard Today"),
    ("Inventory Report",     "inv.port",       "Inventory Port Today"),
    ("Master Data",          "md.unit_load",       "Unit Load Qty"),
    ("Master Data",          "md.safety_days",     "Safety Days"),
    ("Master Data",          "md.safety_stock",    "Safety Stock"),
    ("Master Data",          "md.multi_unit_load", "Multi-Unit Load"),
    ("Requirement Splits",   "req.part_no", "Part Number column"),
    ("Requirement Splits",   "req.date",    "Production Date column"),
    ("Requirement Splits",   "req.qty",     "Quantity column"),
    ("Splunk Receiving",     "splunk.part_no", "Part Number column"),
    ("Splunk Receiving",     "splunk.date",    "Delivery Date column"),
    ("Splunk Receiving",     "splunk.qty",     "Quantity column"),
    ("Goods to be Received", "gtr.part_no", "Part Number column"),
    ("Goods to be Received", "gtr.date",    "Arrival Date column"),
    ("Goods to be Received", "gtr.qty",     "Quantity column"),
]


class AdjustmentStore:
    """Persistence layer for column mappings, inventory overrides, and delivery
    adjustments.  All three files live on the shared network path so every user
    picks up the same overrides on their next Generate Coverage run."""

    COLUMN_MAPPING_FILE      = SHAREDNETWORKPATH / "column_mapping.json"
    INVENTORY_OVERRIDES_FILE = SHAREDNETWORKPATH / "inventory_overrides.json"
    DELIVERY_ADJUSTMENTS_FILE = SHAREDNETWORKPATH / "delivery_adjustments.json"

    @staticmethod
    def load_column_mapping() -> dict:
        f = AdjustmentStore.COLUMN_MAPPING_FILE
        if f.exists():
            try:
                with open(f, 'r', encoding='utf-8') as fp:
                    saved = json.load(fp)
                mapping = dict(COLUMN_MAPPING_DEFAULTS)
                mapping.update(saved)
                return mapping
            except Exception as e:
                print(f"AdjustmentStore: error loading column mapping: {e}")
        return dict(COLUMN_MAPPING_DEFAULTS)

    @staticmethod
    def save_column_mapping(mapping: dict):
        f = AdjustmentStore.COLUMN_MAPPING_FILE
        try:
            f.parent.mkdir(parents=True, exist_ok=True)
            with open(f, 'w', encoding='utf-8') as fp:
                json.dump(mapping, fp, indent=2)
        except Exception as e:
            print(f"AdjustmentStore: error saving column mapping: {e}")

    @staticmethod
    def load_inventory_overrides() -> list:
        f = AdjustmentStore.INVENTORY_OVERRIDES_FILE
        if f.exists():
            try:
                with open(f, 'r', encoding='utf-8') as fp:
                    return json.load(fp)
            except Exception as e:
                print(f"AdjustmentStore: error loading inventory overrides: {e}")
        return []

    @staticmethod
    def save_inventory_overrides(records: list):
        f = AdjustmentStore.INVENTORY_OVERRIDES_FILE
        try:
            f.parent.mkdir(parents=True, exist_ok=True)
            with open(f, 'w', encoding='utf-8') as fp:
                json.dump(records, fp, indent=2)
        except Exception as e:
            print(f"AdjustmentStore: error saving inventory overrides: {e}")

    @staticmethod
    def add_inventory_override(part_no: str, adjusted_value: float,
                                reason: str, username: str) -> dict:
        records = AdjustmentStore.load_inventory_overrides()
        for r in records:
            if r['part_no'].upper() == part_no.upper().strip() and r['active']:
                r['active'] = False
        record = {
            'id': str(uuid.uuid4()),
            'part_no': part_no.upper().strip(),
            'adjusted_value': float(adjusted_value),
            'reason': reason.strip(),
            'username': username,
            'timestamp': datetime.now().isoformat(timespec='seconds'),
            'active': True,
        }
        records.append(record)
        AdjustmentStore.save_inventory_overrides(records)
        return record

    @staticmethod
    def deactivate_inventory_override(record_id: str):
        records = AdjustmentStore.load_inventory_overrides()
        for r in records:
            if r['id'] == record_id:
                r['active'] = False
        AdjustmentStore.save_inventory_overrides(records)

    @staticmethod
    def load_delivery_adjustments() -> list:
        f = AdjustmentStore.DELIVERY_ADJUSTMENTS_FILE
        if f.exists():
            try:
                with open(f, 'r', encoding='utf-8') as fp:
                    return json.load(fp)
            except Exception as e:
                print(f"AdjustmentStore: error loading delivery adjustments: {e}")
        return []

    @staticmethod
    def save_delivery_adjustments(records: list):
        f = AdjustmentStore.DELIVERY_ADJUSTMENTS_FILE
        try:
            f.parent.mkdir(parents=True, exist_ok=True)
            with open(f, 'w', encoding='utf-8') as fp:
                json.dump(records, fp, indent=2)
        except Exception as e:
            print(f"AdjustmentStore: error saving delivery adjustments: {e}")

    @staticmethod
    def add_delivery_adjustment(adj_type: str, source: str, part_no: str,
                                 date: str, adjusted_qty: float,
                                 reason: str, username: str,
                                 original_qty=None) -> dict:
        records = AdjustmentStore.load_delivery_adjustments()
        record = {
            'id': str(uuid.uuid4()),
            'type': adj_type,         
            'source': source,         
            'part_no': part_no.upper().strip(),
            'date': date,
            'original_qty': original_qty,
            'adjusted_qty': float(adjusted_qty),
            'reason': reason.strip(),
            'username': username,
            'timestamp': datetime.now().isoformat(timespec='seconds'),
            'active': True,
        }
        records.append(record)
        AdjustmentStore.save_delivery_adjustments(records)
        return record

    @staticmethod
    def deactivate_delivery_adjustment(record_id: str):
        records = AdjustmentStore.load_delivery_adjustments()
        for r in records:
            if r['id'] == record_id:
                r['active'] = False
        AdjustmentStore.save_delivery_adjustments(records)
