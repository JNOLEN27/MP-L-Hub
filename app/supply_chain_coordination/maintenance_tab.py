"""
Maintenance tab for Supply Chain Coordination — admin/power-user only.

Three sub-tabs:
  1. Column Mapping   — remap what column names the coverage engine reads
  2. Inventory Adjustments — override Initial_Stock for a part (quality fallouts, etc.)
  3. Delivery Adjustments  — edit delivery quantities or inject expedites
"""
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QPushButton,
    QLineEdit, QTextEdit, QTableWidget, QTableWidgetItem, QTabWidget,
    QScrollArea, QDoubleSpinBox, QComboBox, QDateEdit, QHeaderView,
    QMessageBox, QSplitter, QFrame, QSizePolicy, QAbstractItemView,
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont, QColor

from app.supply_chain_coordination.adjustment_store import (
    AdjustmentStore, COLUMN_MAPPING_DEFAULTS, COLUMN_MAPPING_META,
)


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

_HDR_STYLE = (
    "QLabel { background-color: #d0e8f8; color: #1a3a6b; font-weight: bold;"
    " padding: 4px 8px; border-radius: 3px; }"
)
_BTN_PRIMARY = (
    "QPushButton { background-color: #156082; color: white; padding: 6px 14px;"
    " border: none; border-radius: 4px; font-weight: bold; }"
    "QPushButton:hover { background-color: #1a7aaa; }"
    "QPushButton:disabled { background-color: #aaa; }"
)
_BTN_DANGER = (
    "QPushButton { background-color: #c0392b; color: white; padding: 4px 10px;"
    " border: none; border-radius: 4px; }"
    "QPushButton:hover { background-color: #e74c3c; }"
)
_TABLE_STYLE = (
    "QTableWidget { gridline-color: #d0d0d0; background-color: white; }"
    "QTableWidget::item { padding: 4px; }"
    "QHeaderView::section { background-color: #f0f0f0; padding: 6px;"
    " border: 1px solid #d0d0d0; font-weight: bold; }"
)


def _make_table(headers: list) -> QTableWidget:
    t = QTableWidget(0, len(headers))
    t.setHorizontalHeaderLabels(headers)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers)
    t.setSelectionBehavior(QAbstractItemView.SelectRows)
    t.verticalHeader().setVisible(False)
    t.horizontalHeader().setStretchLastSection(True)
    t.setStyleSheet(_TABLE_STYLE)
    t.setSortingEnabled(True)
    return t


def _noneditable_item(text: str) -> QTableWidgetItem:
    item = QTableWidgetItem(str(text) if text is not None else "")
    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
    return item


# ---------------------------------------------------------------------------
# Sub-tab 1 — Column Mapping
# ---------------------------------------------------------------------------

class ColumnMappingTab(QWidget):
    """UI for remapping logical field keys to actual column names in imports."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._line_edits = {}   # logical_key → QLineEdit
        self._build_ui()
        self._load_saved()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        # ── info banner ──────────────────────────────────────────────────
        info = QLabel(
            "Changes saved here apply to ALL users on their next Generate Coverage run."
            "  Only modify a value if the source file's column name has changed."
        )
        info.setWordWrap(True)
        info.setStyleSheet(
            "QLabel { background-color: #fff8dc; border: 1px solid #e0c060;"
            " padding: 6px 10px; border-radius: 3px; }"
        )
        root.addWidget(info)

        # ── buttons ──────────────────────────────────────────────────────
        btnrow = QHBoxLayout()
        self._save_btn = QPushButton("Save Mapping")
        self._save_btn.setStyleSheet(_BTN_PRIMARY)
        self._save_btn.clicked.connect(self._save)
        btnrow.addWidget(self._save_btn)

        reset_btn = QPushButton("Reset to Defaults")
        reset_btn.setStyleSheet(_BTN_DANGER)
        reset_btn.clicked.connect(self._reset)
        btnrow.addWidget(reset_btn)

        self._status_label = QLabel("")
        self._status_label.setStyleSheet("color: #156082; font-weight: bold;")
        btnrow.addWidget(self._status_label)
        btnrow.addStretch()
        root.addLayout(btnrow)

        # ── scrollable grid ──────────────────────────────────────────────
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        grid = QGridLayout(container)
        grid.setContentsMargins(4, 4, 4, 4)
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(4)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(2, 1)

        # header row
        for col, txt in enumerate(["Group / Field", "Current Column Name", "Default"]):
            lbl = QLabel(txt)
            lbl.setFont(QFont("Arial", 9, QFont.Bold))
            lbl.setStyleSheet("color: #555;")
            grid.addWidget(lbl, 0, col)

        row = 1
        current_group = None
        for (group, key, field_label) in COLUMN_MAPPING_META:
            if group != current_group:
                current_group = group
                sep = QLabel(group)
                sep.setStyleSheet(_HDR_STYLE)
                sep.setFont(QFont("Arial", 9, QFont.Bold))
                grid.addWidget(sep, row, 0, 1, 3)
                row += 1

            field_lbl = QLabel(f"  {field_label}")
            field_lbl.setStyleSheet("color: #333;")
            grid.addWidget(field_lbl, row, 0)

            le = QLineEdit(COLUMN_MAPPING_DEFAULTS[key])
            le.setFixedHeight(26)
            le.setStyleSheet("font-family: monospace; font-size: 11px;")
            self._line_edits[key] = le
            grid.addWidget(le, row, 1)

            default_lbl = QLabel(COLUMN_MAPPING_DEFAULTS[key])
            default_lbl.setStyleSheet("color: #888; font-size: 10px; font-family: monospace;")
            grid.addWidget(default_lbl, row, 2)

            row += 1

        grid.setRowStretch(row, 1)
        scroll.setWidget(container)
        root.addWidget(scroll)

    def _load_saved(self):
        mapping = AdjustmentStore.load_column_mapping()
        for key, le in self._line_edits.items():
            le.setText(mapping.get(key, COLUMN_MAPPING_DEFAULTS[key]))

    def _save(self):
        mapping = {key: le.text().strip() for key, le in self._line_edits.items()
                   if le.text().strip() and le.text().strip() != COLUMN_MAPPING_DEFAULTS[key]}
        AdjustmentStore.save_column_mapping(mapping)
        self._status_label.setText("Saved ✓")
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(3000, lambda: self._status_label.setText(""))

    def _reset(self):
        reply = QMessageBox.question(
            self, "Reset to Defaults",
            "This will clear all column name overrides and revert to system defaults.\n"
            "Are you sure?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        AdjustmentStore.save_column_mapping({})
        for key, le in self._line_edits.items():
            le.setText(COLUMN_MAPPING_DEFAULTS[key])
        self._status_label.setText("Reset to defaults ✓")
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(3000, lambda: self._status_label.setText(""))


# ---------------------------------------------------------------------------
# Sub-tab 2 — Inventory Adjustments
# ---------------------------------------------------------------------------

class InventoryAdjustmentsTab(QWidget):

    def __init__(self, import_manager, userdata, parent=None):
        super().__init__(parent)
        self._im = import_manager
        self._username = userdata.get('username', 'unknown')
        self._build_ui()
        self._refresh_history()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(10)

        # ── banner ───────────────────────────────────────────────────────
        info = QLabel(
            "Override a part's calculated initial inventory (e.g. after a quality fallout). "
            "The override replaces the sum of Beginning + Yard + Port inventory for that part. "
            "Deactivate to revert to the calculated value."
        )
        info.setWordWrap(True)
        info.setStyleSheet(
            "QLabel { background-color: #fff8dc; border: 1px solid #e0c060;"
            " padding: 6px 10px; border-radius: 3px; }"
        )
        root.addWidget(info)

        splitter = QSplitter(Qt.Vertical)

        # ── add override form ────────────────────────────────────────────
        form_widget = QWidget()
        form_widget.setMaximumHeight(200)
        form_layout = QVBoxLayout(form_widget)
        form_layout.setContentsMargins(0, 0, 0, 0)
        form_layout.setSpacing(6)

        hdr = QLabel("Add / Update Inventory Override")
        hdr.setStyleSheet(_HDR_STYLE)
        form_layout.addWidget(hdr)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Part Number:"))
        self._part_input = QLineEdit()
        self._part_input.setPlaceholderText("e.g. ABC12345")
        self._part_input.setFixedWidth(160)
        row1.addWidget(self._part_input)

        lookup_btn = QPushButton("Lookup Current Inventory")
        lookup_btn.setStyleSheet(_BTN_PRIMARY)
        lookup_btn.clicked.connect(self._lookup)
        row1.addWidget(lookup_btn)

        self._calc_label = QLabel("Calculated: —")
        self._calc_label.setStyleSheet("color: #156082; font-weight: bold;")
        row1.addWidget(self._calc_label)
        row1.addStretch()
        form_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("New Value:"))
        self._new_value_spin = QDoubleSpinBox()
        self._new_value_spin.setRange(0, 9_999_999)
        self._new_value_spin.setDecimals(0)
        self._new_value_spin.setFixedWidth(120)
        row2.addWidget(self._new_value_spin)

        row2.addWidget(QLabel("Reason:"))
        self._inv_reason = QLineEdit()
        self._inv_reason.setPlaceholderText("Required — e.g. Quality fallout, 50% scrap")
        self._inv_reason.setMinimumWidth(300)
        row2.addWidget(self._inv_reason)

        submit_btn = QPushButton("Submit Override")
        submit_btn.setStyleSheet(_BTN_PRIMARY)
        submit_btn.clicked.connect(self._submit_override)
        row2.addWidget(submit_btn)
        row2.addStretch()
        form_layout.addLayout(row2)

        splitter.addWidget(form_widget)

        # ── history table ────────────────────────────────────────────────
        hist_widget = QWidget()
        hist_layout = QVBoxLayout(hist_widget)
        hist_layout.setContentsMargins(0, 0, 0, 0)
        hist_layout.setSpacing(4)

        hist_hdr = QHBoxLayout()
        hist_hdr.addWidget(QLabel("Adjustment History"))
        refresh_btn = QPushButton("Refresh")
        refresh_btn.setFixedWidth(70)
        refresh_btn.clicked.connect(self._refresh_history)
        hist_hdr.addWidget(refresh_btn)
        hist_hdr.addStretch()
        hist_layout.addLayout(hist_hdr)

        self._hist_table = _make_table([
            "Active", "Part No", "Adjusted Value",
            "Reason", "User", "Timestamp", "Deactivate",
        ])
        self._hist_table.setSortingEnabled(False)
        hist_layout.addWidget(self._hist_table)
        splitter.addWidget(hist_widget)

        splitter.setSizes([180, 400])
        root.addWidget(splitter)

    def _lookup(self):
        part_no = self._part_input.text().strip().upper()
        if not part_no:
            return
        try:
            mapping = AdjustmentStore.load_column_mapping()
            inv_df = self._im.loaddata("current_inventory_report")
            if inv_df.empty:
                self._calc_label.setText("Calculated: (no inventory data loaded)")
                return
            col_part  = mapping.get("inv.part_no",  "PART_NO")
            col_begin = mapping.get("inv.beginning", "BEGINNING_INVENTORY_TODAY")
            col_yard  = mapping.get("inv.yard",  "INVENTORY_YARD_TODAY")
            col_port  = mapping.get("inv.port",  "INVENTORY_PORT_TODAY")

            if col_part not in inv_df.columns:
                self._calc_label.setText(f"Calculated: (column '{col_part}' not found)")
                return

            partrows = inv_df[inv_df[col_part].astype(str).str.upper() == part_no]
            if partrows.empty:
                self._calc_label.setText("Calculated: part not found in inventory data")
                return

            def _sum_col(col):
                if col in partrows.columns:
                    import pandas as pd
                    return pd.to_numeric(
                        partrows[col].astype(str).str.replace(',', '', regex=False),
                        errors='coerce',
                    ).fillna(0).sum()
                return 0

            total = _sum_col(col_begin) + _sum_col(col_yard) + _sum_col(col_port)
            self._calc_label.setText(f"Calculated: {int(total):,}")
            self._new_value_spin.setValue(total)
        except Exception as e:
            self._calc_label.setText(f"Error: {e}")

    def _submit_override(self):
        part_no = self._part_input.text().strip().upper()
        reason  = self._inv_reason.text().strip()
        if not part_no:
            QMessageBox.warning(self, "Missing Field", "Please enter a Part Number.")
            return
        if not reason:
            QMessageBox.warning(self, "Missing Field", "Reason is required.")
            return
        new_val = self._new_value_spin.value()
        AdjustmentStore.add_inventory_override(part_no, new_val, reason, self._username)
        self._part_input.clear()
        self._inv_reason.clear()
        self._calc_label.setText("Calculated: —")
        self._new_value_spin.setValue(0)
        self._refresh_history()

    def _refresh_history(self):
        records = AdjustmentStore.load_inventory_overrides()
        records_sorted = sorted(records, key=lambda r: r['timestamp'], reverse=True)
        self._hist_table.setRowCount(len(records_sorted))
        for row, r in enumerate(records_sorted):
            active = r.get('active', False)
            status_item = _noneditable_item("✓ Active" if active else "Inactive")
            if active:
                status_item.setForeground(QColor(0, 120, 0))
                status_item.setBackground(QColor(230, 255, 230))
            else:
                status_item.setForeground(QColor(150, 150, 150))
            self._hist_table.setItem(row, 0, status_item)
            self._hist_table.setItem(row, 1, _noneditable_item(r.get('part_no', '')))
            self._hist_table.setItem(row, 2, _noneditable_item(f"{r.get('adjusted_value', 0):,.0f}"))
            self._hist_table.setItem(row, 3, _noneditable_item(r.get('reason', '')))
            self._hist_table.setItem(row, 4, _noneditable_item(r.get('username', '')))
            self._hist_table.setItem(row, 5, _noneditable_item(r.get('timestamp', '')))
            if active:
                deact_btn = QPushButton("Deactivate")
                deact_btn.setStyleSheet(_BTN_DANGER)
                record_id = r['id']
                deact_btn.clicked.connect(lambda _, rid=record_id: self._deactivate(rid))
                self._hist_table.setCellWidget(row, 6, deact_btn)
            else:
                self._hist_table.setItem(row, 6, _noneditable_item(""))
        self._hist_table.resizeColumnsToContents()
        self._hist_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)

    def _deactivate(self, record_id: str):
        AdjustmentStore.deactivate_inventory_override(record_id)
        self._refresh_history()


# ---------------------------------------------------------------------------
# Sub-tab 3 — Delivery Adjustments
# ---------------------------------------------------------------------------

class DeliveryAdjustmentsTab(QWidget):

    def __init__(self, import_manager, userdata, parent=None):
        super().__init__(parent)
        self._im = import_manager
        self._username = userdata.get('username', 'unknown')
        self._search_results = []   # list of dicts for the current search
        self._build_ui()
        self._refresh_history()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(10)

        # ── banner ───────────────────────────────────────────────────────
        info = QLabel(
            "Edit existing delivery quantities (e.g. 800 planned, 300 shipped) or "
            "add expedite deliveries that don't appear in the import reports. "
            "All changes are applied before coverage is recalculated."
        )
        info.setWordWrap(True)
        info.setStyleSheet(
            "QLabel { background-color: #fff8dc; border: 1px solid #e0c060;"
            " padding: 6px 10px; border-radius: 3px; }"
        )
        root.addWidget(info)

        top_splitter = QSplitter(Qt.Horizontal)

        # ── LEFT: Edit existing ──────────────────────────────────────────
        edit_widget = QWidget()
        edit_layout = QVBoxLayout(edit_widget)
        edit_layout.setContentsMargins(0, 0, 0, 0)
        edit_layout.setSpacing(6)

        edit_hdr = QLabel("Edit Existing Delivery")
        edit_hdr.setStyleSheet(_HDR_STYLE)
        edit_layout.addWidget(edit_hdr)

        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Part No:"))
        self._edit_part = QLineEdit()
        self._edit_part.setPlaceholderText("e.g. ABC12345")
        self._edit_part.setFixedWidth(140)
        search_row.addWidget(self._edit_part)

        search_row.addWidget(QLabel("Source:"))
        self._edit_source = QComboBox()
        self._edit_source.addItems(["Splunk Receiving", "Goods to be Received"])
        self._edit_source.setFixedWidth(160)
        search_row.addWidget(self._edit_source)

        search_btn = QPushButton("Search")
        search_btn.setStyleSheet(_BTN_PRIMARY)
        search_btn.clicked.connect(self._search_deliveries)
        search_row.addWidget(search_btn)
        search_row.addStretch()
        edit_layout.addLayout(search_row)

        self._search_table = _make_table(["Part No", "Date", "Qty", "Source detail"])
        self._search_table.setFixedHeight(150)
        self._search_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._search_table.itemSelectionChanged.connect(self._on_search_selection)
        edit_layout.addWidget(self._search_table)

        save_row = QHBoxLayout()
        save_row.addWidget(QLabel("New Qty:"))
        self._edit_qty = QDoubleSpinBox()
        self._edit_qty.setRange(0, 9_999_999)
        self._edit_qty.setDecimals(0)
        self._edit_qty.setFixedWidth(110)
        save_row.addWidget(self._edit_qty)

        save_row.addWidget(QLabel("Reason:"))
        self._edit_reason = QLineEdit()
        self._edit_reason.setPlaceholderText("Required — e.g. Only 300 of 800 shipped")
        self._edit_reason.setMinimumWidth(220)
        save_row.addWidget(self._edit_reason)

        save_edit_btn = QPushButton("Save Edit")
        save_edit_btn.setStyleSheet(_BTN_PRIMARY)
        save_edit_btn.clicked.connect(self._save_edit)
        save_row.addWidget(save_edit_btn)
        save_row.addStretch()
        edit_layout.addLayout(save_row)
        edit_layout.addStretch()
        top_splitter.addWidget(edit_widget)

        # ── RIGHT: Add expedite ──────────────────────────────────────────
        add_widget = QWidget()
        add_layout = QVBoxLayout(add_widget)
        add_layout.setContentsMargins(0, 0, 0, 0)
        add_layout.setSpacing(6)

        add_hdr = QLabel("Add Expedite / New Delivery")
        add_hdr.setStyleSheet(_HDR_STYLE)
        add_layout.addWidget(add_hdr)

        form_grid = QGridLayout()
        form_grid.setHorizontalSpacing(8)
        form_grid.setVerticalSpacing(6)

        form_grid.addWidget(QLabel("Part No:"), 0, 0)
        self._add_part = QLineEdit()
        self._add_part.setPlaceholderText("e.g. ABC12345")
        form_grid.addWidget(self._add_part, 0, 1)

        form_grid.addWidget(QLabel("Delivery Date:"), 1, 0)
        self._add_date = QDateEdit(QDate.currentDate())
        self._add_date.setCalendarPopup(True)
        self._add_date.setDisplayFormat("yyyy-MM-dd")
        form_grid.addWidget(self._add_date, 1, 1)

        form_grid.addWidget(QLabel("Quantity:"), 2, 0)
        self._add_qty = QDoubleSpinBox()
        self._add_qty.setRange(1, 9_999_999)
        self._add_qty.setDecimals(0)
        form_grid.addWidget(self._add_qty, 2, 1)

        form_grid.addWidget(QLabel("Source:"), 3, 0)
        self._add_source = QComboBox()
        self._add_source.addItems(["Splunk Receiving", "Goods to be Received"])
        form_grid.addWidget(self._add_source, 3, 1)

        form_grid.addWidget(QLabel("Reason:"), 4, 0)
        self._add_reason = QLineEdit()
        self._add_reason.setPlaceholderText("Required — e.g. Air freight expedite")
        form_grid.addWidget(self._add_reason, 4, 1)

        add_layout.addLayout(form_grid)

        add_btn = QPushButton("Add Expedite")
        add_btn.setStyleSheet(_BTN_PRIMARY)
        add_btn.clicked.connect(self._add_expedite)
        add_layout.addWidget(add_btn, alignment=Qt.AlignLeft)
        add_layout.addStretch()
        top_splitter.addWidget(add_widget)

        top_splitter.setSizes([500, 360])
        root.addWidget(top_splitter)

        # ── history table ────────────────────────────────────────────────
        hist_hdr = QHBoxLayout()
        hist_lbl = QLabel("Adjustment History")
        hist_lbl.setStyleSheet(_HDR_STYLE)
        hist_hdr.addWidget(hist_lbl)
        refresh_btn = QPushButton("Refresh")
        refresh_btn.setFixedWidth(70)
        refresh_btn.clicked.connect(self._refresh_history)
        hist_hdr.addWidget(refresh_btn)
        hist_hdr.addStretch()
        root.addLayout(hist_hdr)

        self._hist_table = _make_table([
            "Active", "Type", "Source", "Part No", "Date",
            "Original Qty", "Adjusted Qty", "Reason", "User", "Timestamp", "Deactivate",
        ])
        self._hist_table.setSortingEnabled(False)
        root.addWidget(self._hist_table)

    # ── search & edit helpers ────────────────────────────────────────────

    def _source_key(self, combo: QComboBox) -> str:
        return "splunk" if combo.currentIndex() == 0 else "goods_to_be_received"

    def _search_deliveries(self):
        part_no = self._edit_part.text().strip().upper()
        if not part_no:
            QMessageBox.warning(self, "Missing Field", "Please enter a Part Number.")
            return
        source_key = self._source_key(self._edit_source)
        category = "splunk_receiving_data" if source_key == "splunk" else "goods_to_be_received"
        try:
            mapping = AdjustmentStore.load_column_mapping()
            df = self._im.loaddata(category)
            if df.empty:
                QMessageBox.information(self, "No Data", f"No data found for {category}.")
                return

            # Resolve column names
            if source_key == "splunk":
                col_part = mapping.get("splunk.part_no", "Part Number")
                col_date = mapping.get("splunk.date",    "Load Delivery Date Final")
                col_qty  = mapping.get("splunk.qty",     "Quantity")
                col_detail = "TO Number"
            else:
                col_part = mapping.get("gtr.part_no", "ARTNR")
                col_date = mapping.get("gtr.date",    "ANK_TID_SENAST")
                col_qty  = mapping.get("gtr.qty",     "ARTAN")
                col_detail = "FRAKT_TID_SENAST"

            if col_part not in df.columns:
                QMessageBox.warning(self, "Column Not Found",
                    f"Part number column '{col_part}' not found in {category}.\n"
                    "Check the Column Mapping tab.")
                return

            matches = df[df[col_part].astype(str).str.upper() == part_no]
            self._search_results = []
            self._search_table.setRowCount(0)

            for _, row in matches.iterrows():
                date_val = str(row.get(col_date, "")) if col_date in df.columns else ""
                qty_val  = row.get(col_qty, 0) if col_qty in df.columns else 0
                detail   = str(row.get(col_detail, "")) if col_detail in df.columns else ""
                entry = {
                    'part_no': part_no,
                    'date': date_val,
                    'qty': qty_val,
                    'source': source_key,
                    'detail': detail,
                }
                self._search_results.append(entry)
                r = self._search_table.rowCount()
                self._search_table.insertRow(r)
                self._search_table.setItem(r, 0, _noneditable_item(part_no))
                self._search_table.setItem(r, 1, _noneditable_item(date_val))
                self._search_table.setItem(r, 2, _noneditable_item(f"{qty_val:,.0f}" if qty_val else qty_val))
                self._search_table.setItem(r, 3, _noneditable_item(detail))

            if not self._search_results:
                QMessageBox.information(self, "Not Found",
                    f"No records found for part '{part_no}' in {category}.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Search failed: {e}")

    def _on_search_selection(self):
        rows = self._search_table.selectionModel().selectedRows()
        if not rows:
            return
        idx = rows[0].row()
        if idx < len(self._search_results):
            qty = self._search_results[idx]['qty']
            try:
                self._edit_qty.setValue(float(qty))
            except (TypeError, ValueError):
                pass

    def _save_edit(self):
        rows = self._search_table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.warning(self, "No Selection", "Select a delivery record from the search results first.")
            return
        reason = self._edit_reason.text().strip()
        if not reason:
            QMessageBox.warning(self, "Missing Field", "Reason is required.")
            return
        idx = rows[0].row()
        if idx >= len(self._search_results):
            return
        rec = self._search_results[idx]
        new_qty = self._edit_qty.value()
        AdjustmentStore.add_delivery_adjustment(
            adj_type='edit',
            source=rec['source'],
            part_no=rec['part_no'],
            date=rec['date'],
            adjusted_qty=new_qty,
            reason=reason,
            username=self._username,
            original_qty=rec['qty'],
        )
        self._edit_reason.clear()
        self._refresh_history()
        QMessageBox.information(self, "Saved",
            f"Edit saved for {rec['part_no']} on {rec['date']}: "
            f"{rec['qty']} → {int(new_qty):,}")

    def _add_expedite(self):
        part_no = self._add_part.text().strip().upper()
        reason  = self._add_reason.text().strip()
        if not part_no:
            QMessageBox.warning(self, "Missing Field", "Please enter a Part Number.")
            return
        if not reason:
            QMessageBox.warning(self, "Missing Field", "Reason is required.")
            return
        date_str = self._add_date.date().toString("yyyy-MM-dd")
        qty = self._add_qty.value()
        source_key = self._source_key(self._add_source)
        AdjustmentStore.add_delivery_adjustment(
            adj_type='add',
            source=source_key,
            part_no=part_no,
            date=date_str,
            adjusted_qty=qty,
            reason=reason,
            username=self._username,
            original_qty=None,
        )
        self._add_part.clear()
        self._add_reason.clear()
        self._add_qty.setValue(1)
        self._refresh_history()
        QMessageBox.information(self, "Added",
            f"Expedite added: {part_no}, {int(qty):,} pcs on {date_str}")

    def _refresh_history(self):
        records = AdjustmentStore.load_delivery_adjustments()
        records_sorted = sorted(records, key=lambda r: r['timestamp'], reverse=True)
        self._hist_table.setRowCount(len(records_sorted))
        for row, r in enumerate(records_sorted):
            active = r.get('active', False)
            status_item = _noneditable_item("✓ Active" if active else "Inactive")
            if active:
                status_item.setForeground(QColor(0, 120, 0))
                status_item.setBackground(QColor(230, 255, 230))
            else:
                status_item.setForeground(QColor(150, 150, 150))
            self._hist_table.setItem(row, 0, status_item)
            self._hist_table.setItem(row, 1, _noneditable_item(r.get('type', '')))
            self._hist_table.setItem(row, 2, _noneditable_item(r.get('source', '')))
            self._hist_table.setItem(row, 3, _noneditable_item(r.get('part_no', '')))
            self._hist_table.setItem(row, 4, _noneditable_item(r.get('date', '')))
            orig = r.get('original_qty')
            self._hist_table.setItem(row, 5, _noneditable_item(
                f"{orig:,.0f}" if orig is not None else "—"))
            self._hist_table.setItem(row, 6, _noneditable_item(
                f"{r.get('adjusted_qty', 0):,.0f}"))
            self._hist_table.setItem(row, 7, _noneditable_item(r.get('reason', '')))
            self._hist_table.setItem(row, 8, _noneditable_item(r.get('username', '')))
            self._hist_table.setItem(row, 9, _noneditable_item(r.get('timestamp', '')))
            if active:
                deact_btn = QPushButton("Deactivate")
                deact_btn.setStyleSheet(_BTN_DANGER)
                record_id = r['id']
                deact_btn.clicked.connect(lambda _, rid=record_id: self._deactivate(rid))
                self._hist_table.setCellWidget(row, 10, deact_btn)
            else:
                self._hist_table.setItem(row, 10, _noneditable_item(""))
        self._hist_table.resizeColumnsToContents()
        self._hist_table.horizontalHeader().setSectionResizeMode(7, QHeaderView.Stretch)

    def _deactivate(self, record_id: str):
        AdjustmentStore.deactivate_delivery_adjustment(record_id)
        self._refresh_history()


# ---------------------------------------------------------------------------
# MaintenanceTab — container shown in the SCC window
# ---------------------------------------------------------------------------

class MaintenanceTab(QWidget):

    def __init__(self, import_manager, userdata, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        tabs = QTabWidget()

        tabs.addTab(ColumnMappingTab(),                           "Column Mapping")
        tabs.addTab(InventoryAdjustmentsTab(import_manager, userdata), "Inventory Adjustments")
        tabs.addTab(DeliveryAdjustmentsTab(import_manager, userdata),  "Delivery Adjustments")

        layout.addWidget(tabs)
