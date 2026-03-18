import torch
import torch.nn as nn
import torch.optim as optim
import pandas as pd
import numpy as np
from torch.utils.data import DataLoader, TensorDataset
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler, LabelEncoder
from typing import Dict, List, Tuple, Optional

# ---------------------------------------------------------------------------
# Column name constants — adjust here if upstream data changes column names
# ---------------------------------------------------------------------------
PART_COL        = "PART"            # Part number in master data
COUNTRY_COL     = "SUPP_SHP_COUNTRY"  # Country column in master data
SAFETY_COL      = "SAFETY"          # Target: safety stock value

REQ_PART_COL    = "ARTNR"           # Part number in req split files
REQ_DATE_COL    = "PRODDAG"         # Production date in req split files
REQ_QTY_COL     = "ARTAN"          # Daily quantity in req split files

# Numeric features fed into the model
NUMERIC_FEATURES = [
    "PRICE",
    "UNIT_LOAD_QTY",
    "SHIP_QTY",
    "STOCK",
    "FIXED_PERIOD",
    "TTT_DAYS",
    "avg_daily_usage",  # computed column added before training
]

# Categorical features that will be label-encoded
CATEGORICAL_FEATURES = [
    "Region",           # derived from SUPP_SHP_COUNTRY
    "SCC_NAME",
    "CONSOLIDATOR",
    "MULT_UNIT_LOAD_VALID",
]


# ---------------------------------------------------------------------------
# PyTorch model definition
# ---------------------------------------------------------------------------
class SafetyStockModel(nn.Module):
    def __init__(self, input_size: int, hidden_sizes: List[int], dropout: float):
        super().__init__()
        layers: List[nn.Module] = []
        prev = input_size
        for h in hidden_sizes:
            layers += [nn.Linear(prev, h), nn.ReLU(), nn.Dropout(dropout)]
            prev = h
        layers.append(nn.Linear(prev, 1))
        self.network = nn.Sequential(*layers)

    def forward(self, x: torch.Tensor) -> torch.Tensor:
        return self.network(x).squeeze(-1)


# ---------------------------------------------------------------------------
# Main class
# ---------------------------------------------------------------------------
class InventorybyPurposeNeuralNetwork:
    def __init__(self, import_manager):
        self.import_manager = import_manager
        self.scaler = StandardScaler()
        self.label_encoders: Dict[str, LabelEncoder] = {}
        self.model: Optional[SafetyStockModel] = None
        self.feature_columns: List[str] = []

    # ------------------------------------------------------------------
    # Data loading
    # ------------------------------------------------------------------
    def loadrequireddata(self) -> Tuple[bool, str, Dict[str, pd.DataFrame]]:
        try:
            data = {
                'master_data':  self.import_manager.loaddata("master_data"),
                'req_split_1':  self.import_manager.loaddata("part_requirement_split_1"),
                'req_split_2':  self.import_manager.loaddata("part_requirement_split_2"),
                'req_split_3':  self.import_manager.loaddata("part_requirement_split_3"),
            }

            if data['master_data'].empty:
                return False, "Master Data is required but could not be loaded.", {}

            missing_splits = [k for k in ('req_split_1', 'req_split_2', 'req_split_3')
                              if data[k].empty]
            if missing_splits:
                return False, f"Could not load: {missing_splits}", {}

            return True, "Data loaded successfully.", data

        except Exception as e:
            return False, f"Error loading data: {str(e)}", {}

    # ------------------------------------------------------------------
    # Region mapping
    # ------------------------------------------------------------------
    def determineregion(self, country) -> str:
        if pd.isna(country) or not country:
            return "No Country Found"

        country = str(country).upper().strip()

        if country == 'USA':
            return 'USA'
        if country in ('MEXICO', 'CANADA'):
            return 'MEX'
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

        return 'Country not mapped. Reach out to Admin'

    # ------------------------------------------------------------------
    # Daily usage calculation
    # ------------------------------------------------------------------
    def calculatedailyusage(
        self,
        masterdata: pd.DataFrame,
        reqsplits: List[pd.DataFrame],
    ) -> pd.Series:
        """
        Compute the average daily usage per part using a region-specific
        forward-looking date window:

          USA / MEX  → weeks 1–4   (today + 0  … today + 27 days)
          EMEA       → weeks 5–9   (today + 28 … today + 62 days)
          APAC       → weeks 10–14 (today + 63 … today + 97 days)

        The req splits are in long format: each row is one
        (ARTNR, PRODDAG, SKIFT) with a quantity in ARTAN.
        Daily usage = sum of all shifts per day, averaged over the window.
        """
        today = pd.Timestamp.today().normalize()

        region_windows = {
            'USA':  (today,                          today + pd.Timedelta(days=27)),
            'MEX':  (today,                          today + pd.Timedelta(days=27)),
            'EMEA': (today + pd.Timedelta(days=28),  today + pd.Timedelta(days=62)),
            'APAC': (today + pd.Timedelta(days=63),  today + pd.Timedelta(days=97)),
        }

        # Combine all req splits and parse production dates
        combined = pd.concat(reqsplits, ignore_index=True)
        combined[REQ_DATE_COL] = pd.to_datetime(combined[REQ_DATE_COL], errors='coerce')
        combined[REQ_QTY_COL]  = pd.to_numeric(combined[REQ_QTY_COL],  errors='coerce').fillna(0)

        # Sum across shifts → one row per (part, date)
        daily = (
            combined
            .groupby([REQ_PART_COL, REQ_DATE_COL], as_index=False)[REQ_QTY_COL]
            .sum()
        )

        # Build a part → region lookup from master data
        region_lookup = dict(
            zip(masterdata[PART_COL],
                masterdata[COUNTRY_COL].apply(self.determineregion))
        )

        results = []
        for part in masterdata[PART_COL]:
            region = region_lookup.get(part, 'No Country Found')
            window = region_windows.get(region)

            if window is None:
                results.append(0.0)
                continue

            start, end = window
            partrows = daily[
                (daily[REQ_PART_COL] == part) &
                (daily[REQ_DATE_COL] >= start) &
                (daily[REQ_DATE_COL] <= end)
            ]

            results.append(float(partrows[REQ_QTY_COL].mean()) if not partrows.empty else 0.0)

        return pd.Series(results, index=masterdata.index, name='avg_daily_usage')

    # ------------------------------------------------------------------
    # Feature preparation
    # ------------------------------------------------------------------
    def preparefeatures(
        self,
        data: Dict[str, pd.DataFrame],
    ) -> Tuple[bool, str, Optional[np.ndarray], Optional[np.ndarray]]:
        try:
            masterdata = data['master_data'].copy()
            reqsplits  = [data['req_split_1'], data['req_split_2'], data['req_split_3']]

            # Derived columns
            masterdata['Region']          = masterdata[COUNTRY_COL].apply(self.determineregion)
            masterdata['avg_daily_usage'] = self.calculatedailyusage(masterdata, reqsplits)

            # Drop rows where the target is missing
            masterdata = masterdata.dropna(subset=[SAFETY_COL])
            if masterdata.empty:
                return False, f"No rows with a valid '{SAFETY_COL}' value.", None, None

            # Encode categorical features
            for col in CATEGORICAL_FEATURES:
                if col in masterdata.columns:
                    le = LabelEncoder()
                    masterdata[col] = le.fit_transform(masterdata[col].astype(str))
                    self.label_encoders[col] = le

            all_features = NUMERIC_FEATURES + CATEGORICAL_FEATURES
            missing = [f for f in all_features if f not in masterdata.columns]
            if missing:
                return False, f"Missing feature columns: {missing}", None, None

            self.feature_columns = all_features

            X = masterdata[all_features].apply(pd.to_numeric, errors='coerce').fillna(0).values.astype(np.float32)
            y = pd.to_numeric(masterdata[SAFETY_COL], errors='coerce').fillna(0).values.astype(np.float32)

            return True, f"Features prepared: {X.shape[0]} samples, {X.shape[1]} features.", X, y

        except Exception as e:
            return False, f"Error preparing features: {str(e)}", None, None

    # ------------------------------------------------------------------
    # Model construction
    # ------------------------------------------------------------------
    def buildmodel(
        self,
        input_size: int,
        hidden_sizes: List[int] = [128, 64, 32],
        dropout: float = 0.3,
    ) -> SafetyStockModel:
        self.model = SafetyStockModel(input_size, hidden_sizes, dropout)
        return self.model

    # ------------------------------------------------------------------
    # Training
    # ------------------------------------------------------------------
    def train(
        self,
        data: Dict[str, pd.DataFrame],
        epochs: int = 100,
        batch_size: int = 32,
        learning_rate: float = 1e-3,
        test_size: float = 0.2,
        hidden_sizes: List[int] = [128, 64, 32],
        dropout: float = 0.3,
    ) -> Tuple[bool, str, Dict]:
        """
        Full pipeline: prepare features → scale → split → train → evaluate.
        Returns (success, message, metrics_dict).
        """
        success, message, X, y = self.preparefeatures(data)
        if not success:
            return False, message, {}

        X_scaled = self.scaler.fit_transform(X)

        X_train, X_test, y_train, y_test = train_test_split(
            X_scaled, y, test_size=test_size, random_state=42,
        )

        X_train_t = torch.tensor(X_train, dtype=torch.float32)
        y_train_t = torch.tensor(y_train, dtype=torch.float32)
        X_test_t  = torch.tensor(X_test,  dtype=torch.float32)
        y_test_t  = torch.tensor(y_test,  dtype=torch.float32)

        loader   = DataLoader(TensorDataset(X_train_t, y_train_t), batch_size=batch_size, shuffle=True)
        criterion = nn.MSELoss()

        self.buildmodel(X_train.shape[1], hidden_sizes, dropout)
        optimizer = optim.Adam(self.model.parameters(), lr=learning_rate)

        train_losses: List[float] = []
        self.model.train()
        for _ in range(epochs):
            epoch_loss = 0.0
            for X_batch, y_batch in loader:
                optimizer.zero_grad()
                loss = criterion(self.model(X_batch), y_batch)
                loss.backward()
                optimizer.step()
                epoch_loss += loss.item() * len(X_batch)
            train_losses.append(epoch_loss / len(X_train_t))

        self.model.eval()
        with torch.no_grad():
            preds = self.model(X_test_t).numpy()

        mae  = float(np.mean(np.abs(preds - y_test)))
        rmse = float(np.sqrt(np.mean((preds - y_test) ** 2)))

        metrics = {
            'train_losses': train_losses,
            'test_mse':     float(np.mean((preds - y_test) ** 2)),
            'test_mae':     mae,
            'test_rmse':    rmse,
            'n_train':      len(X_train),
            'n_test':       len(X_test),
        }

        return True, f"Training complete — Test RMSE: {rmse:.4f}  MAE: {mae:.4f}", metrics

    # ------------------------------------------------------------------
    # Inference
    # ------------------------------------------------------------------
    def predict(
        self,
        masterdata: pd.DataFrame,
        reqsplits: List[pd.DataFrame],
    ) -> Optional[np.ndarray]:
        """Predict safety stock for new / unseen master data."""
        if self.model is None or not self.feature_columns:
            return None

        df = masterdata.copy()
        df['Region']          = df[COUNTRY_COL].apply(self.determineregion)
        df['avg_daily_usage'] = self.calculatedailyusage(df, reqsplits)

        for col, le in self.label_encoders.items():
            if col in df.columns:
                df[col] = le.transform(df[col].astype(str))

        X = df[self.feature_columns].apply(pd.to_numeric, errors='coerce').fillna(0).values.astype(np.float32)
        X_scaled = self.scaler.transform(X)

        self.model.eval()
        with torch.no_grad():
            return self.model(torch.tensor(X_scaled, dtype=torch.float32)).numpy()

    # ------------------------------------------------------------------
    # Persistence
    # ------------------------------------------------------------------
    def savemodel(self, filepath: str) -> bool:
        if self.model is None:
            return False
        torch.save({
            'model_state':    self.model.state_dict(),
            'scaler':         self.scaler,
            'label_encoders': self.label_encoders,
            'feature_columns': self.feature_columns,
        }, filepath)
        return True

    def loadmodel(
        self,
        filepath: str,
        hidden_sizes: List[int] = [128, 64, 32],
        dropout: float = 0.3,
    ) -> bool:
        try:
            checkpoint = torch.load(filepath)
            self.feature_columns  = checkpoint['feature_columns']
            self.scaler           = checkpoint['scaler']
            self.label_encoders   = checkpoint['label_encoders']
            self.buildmodel(len(self.feature_columns), hidden_sizes, dropout)
            self.model.load_state_dict(checkpoint['model_state'])
            return True
        except Exception:
            return False
