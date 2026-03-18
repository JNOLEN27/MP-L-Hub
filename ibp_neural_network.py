import torch
import torch.nn as nn
import torch.optim as optim
import pandas as pd
import numpy as np
from torch.utils.data import DataLoader, TensorDataset
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler, LabelEncoder
from typing import Dict, List, Tuple, Optional

partcol      = "PART"
countrycol   = "SUPP_SHP_COUNTRY"
sscol        = "STOCK"
sdcol        = "SAFETY"   # target: safety stock

reqpartcol = "ARTNR"
reqdatecol = "PRODDAG"
reqqtycol  = "ARTAN"

numericfeatures = [
    "PRICE",
    "UNIT_LOAD_QTY",
    "MULT_UNIT_LOAD_VALID",
    "SHIP_QTY",
    "STOCK",
    # "SAFETY" removed — it is the target, including it here causes data leakage
    "FIXED_PERIOD",
    "TTT_DAYS",
    "avgdailyusage",
]

categoricalfeatures = [
    "Region",
    "SCC_NAME",
    "CONSOLIDATOR",
    "SUPP_SHP_COUNTRY",
    "SUPP_SHP",
]


class SafetyStockModel(nn.Module):
    def __init__(self, inputsize: int, hiddensizes: List[int], dropout: float):
        super().__init__()
        layers: List[nn.Module] = []
        prev = inputsize
        for h in hiddensizes:
            layers += [nn.Linear(prev, h), nn.ReLU(), nn.Dropout(dropout)]
            prev = h
        layers.append(nn.Linear(prev, 1))
        self.network = nn.Sequential(*layers)

    def forward(self, x: torch.Tensor) -> torch.Tensor:
        return self.network(x).squeeze(-1)


class InventorybyPurposeNeuralNetwork:
    def __init__(self, import_manager):
        self.import_manager = import_manager
        self.scaler = StandardScaler()
        self.labelencoders: Dict[str, LabelEncoder] = {}  # fix 1: : not =
        self.model: Optional[SafetyStockModel] = None
        self.featurecolumns: List[str] = []
        self.lastpredictions: Optional[pd.DataFrame] = None

    def loadrequireddata(self) -> Tuple[bool, str, Dict[str, pd.DataFrame]]:
        try:
            data = {
                'master_data': self.import_manager.loaddata("master_data"),
                'req_split_1': self.import_manager.loaddata("part_requirement_split_1"),
                'req_split_2': self.import_manager.loaddata("part_requirement_split_2"),
                'req_split_3': self.import_manager.loaddata("part_requirement_split_3"),
            }

            if data['master_data'].empty:
                return False, "Master Data is required but could not be loaded.", data

            missingsplits = [k for k in ('req_split_1', 'req_split_2', 'req_split_3') if data[k].empty]
            if missingsplits:
                return False, f"Could not load: {missingsplits}", {}

            return True, "Data loaded successfully", data

        except Exception as e:
            return False, f"Error loading data: {str(e)}", {}

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
            return "EMEA"
        if country in ('CHINA', 'SOUTH KOREA', 'THAILAND', 'VIETNAM'):
            return "APAC"

        return 'Country not mapped. Reach out to Admin'

    def calculatedailyusage(self, masterdata: pd.DataFrame, reqsplits: List[pd.DataFrame]) -> pd.Series:
        today = pd.Timestamp.today().normalize()
        regionwindows = {
            'USA':  (today,                          today + pd.Timedelta(days=27)),
            'MEX':  (today,                          today + pd.Timedelta(days=27)),
            'EMEA': (today + pd.Timedelta(days=28),  today + pd.Timedelta(days=62)),
            'APAC': (today + pd.Timedelta(days=63),  today + pd.Timedelta(days=97)),
        }

        combined = pd.concat(reqsplits, ignore_index=True)  # fix 3: ignore_index
        combined[reqdatecol] = pd.to_datetime(combined[reqdatecol], errors='coerce')
        combined[reqqtycol]  = pd.to_numeric(combined[reqqtycol],  errors='coerce').fillna(0)

        daily = combined.groupby([reqpartcol, reqdatecol], as_index=False)[reqqtycol].sum()
        regionlookup = dict(zip(masterdata[partcol], masterdata[countrycol].apply(self.determineregion)))

        results = []
        for part in masterdata[partcol]:
            region = regionlookup.get(part, 'No Country Found')
            window = regionwindows.get(region)  # fix 2: regionwindows not regionlookup

            if window is None:
                results.append(0.0)
                continue

            start, end = window
            partrows = daily[
                (daily[reqpartcol] == part) &
                (daily[reqdatecol] >= start) &
                (daily[reqdatecol] <= end)
            ]

            results.append(float(partrows[reqqtycol].mean()) if not partrows.empty else 0.0)

        return pd.Series(results, index=masterdata.index, name='avgdailyusage')

    def preparefeatures(self, data: Dict[str, pd.DataFrame]) -> Tuple[bool, str, Optional[np.ndarray], Optional[np.ndarray]]:
        try:
            masterdata = data['master_data'].copy()
            reqsplits  = [data['req_split_1'], data['req_split_2'], data['req_split_3']]
            masterdata['Region']       = masterdata[countrycol].apply(self.determineregion)
            masterdata['avgdailyusage'] = self.calculatedailyusage(masterdata, reqsplits)
            masterdata = masterdata.dropna(subset=[sdcol])  # fix 8: target is sdcol (SAFETY)

            if masterdata.empty:
                return False, f"No rows with a valid {sdcol} value.", None, None

            for col in categoricalfeatures:
                if col in masterdata.columns:
                    le = LabelEncoder()
                    masterdata[col] = le.fit_transform(masterdata[col].astype(str))
                    self.labelencoders[col] = le  # fix 4: le not len

            allfeatures = numericfeatures + categoricalfeatures
            missing = [f for f in allfeatures if f not in masterdata.columns]

            if missing:
                return False, f"Missing feature columns: {missing}", None, None

            self.featurecolumns = allfeatures

            x = masterdata[allfeatures].apply(pd.to_numeric, errors='coerce').fillna(0).values.astype(np.float32)
            y = pd.to_numeric(masterdata[sdcol], errors='coerce').fillna(0).values.astype(np.float32)  # fix 8
            partids = masterdata[partcol].values if partcol in masterdata.columns else np.arange(len(x))

            return True, f"Features prepared: {x.shape[0]} samples, {x.shape[1]} features.", x, y, partids

        except Exception as e:
            return False, f"Error preparing features: {str(e)}", None, None, None

    def buildmodel(self, inputsize: int, hiddensizes: List[int] = [128, 64, 32], dropout: float = 0.3) -> SafetyStockModel:
        self.model = SafetyStockModel(inputsize, hiddensizes, dropout)
        return self.model

    def train(self, data: Dict[str, pd.DataFrame], epochs: int = 100, batchsize: int = 32, learningrate: float = 1e-3, testsize: float = 0.2, hiddensizes: List[int] = [128, 64, 32], dropout: float = 0.3) -> Tuple[bool, str, Dict]:
        success, message, x, y, partids = self.preparefeatures(data)

        if not success:
            return False, message, {}

        xscaled = self.scaler.fit_transform(x)
        indices = np.arange(len(x))
        xtrain, xtest, ytrain, ytest, idx_train, idx_test = train_test_split(
            xscaled, y, indices, test_size=testsize, random_state=42
        )

        xtraint = torch.tensor(xtrain, dtype=torch.float32)
        ytraint = torch.tensor(ytrain, dtype=torch.float32)
        xtestt  = torch.tensor(xtest,  dtype=torch.float32)
        ytestt  = torch.tensor(ytest,  dtype=torch.float32)

        loader    = DataLoader(TensorDataset(xtraint, ytraint), batch_size=batchsize, shuffle=True)  # fix 7: batch_size
        criterion = nn.MSELoss()

        self.buildmodel(xtrain.shape[1], hiddensizes, dropout)
        optimizer = optim.Adam(self.model.parameters(), lr=learningrate)

        trainlosses: List[float] = []
        self.model.train()

        for _ in range(epochs):
            epochloss = 0.0
            for xbatch, ybatch in loader:
                optimizer.zero_grad()
                loss = criterion(self.model(xbatch), ybatch)
                loss.backward()
                optimizer.step()
                epochloss += loss.item() * len(xbatch)
            trainlosses.append(epochloss / len(xtraint))

        self.model.eval()
        with torch.no_grad():
            xallt  = torch.tensor(xscaled, dtype=torch.float32)
            allpreds = self.model(xallt).numpy()
            preds    = allpreds[idx_test]

        mae  = float(np.mean(np.abs(preds - ytest)))
        rmse = float(np.sqrt(np.mean((preds - ytest) ** 2)))

        split_labels = np.full(len(x), 'train', dtype=object)
        split_labels[idx_test] = 'test'

        self.lastpredictions = pd.DataFrame({
            partcol:      partids,
            'Predicted':  np.round(allpreds, 2),
            'Actual':     y,
            'Error':      np.round(allpreds - y, 2),
            'AbsError':   np.round(np.abs(allpreds - y), 2),
            'Split':      split_labels,
        })

        unseen = data['master_data'][data['master_data'][sdcol].isna()].copy()
        if not unseen.empty:
            reqsplits   = [data['req_split_1'], data['req_split_2'], data['req_split_3']]
            unseenpreds = self.predict(unseen, reqsplits)
            if unseenpreds is not None:
                unseenids = unseen[partcol].values if partcol in unseen.columns else np.arange(len(unseen))
                unseendf  = pd.DataFrame({
                    partcol:     unseenids,
                    'Predicted': np.round(unseenpreds, 2),
                    'Actual':    np.nan,
                    'Error':     np.nan,
                    'AbsError':  np.nan,
                    'Split':     'inference',
                })
                self.lastpredictions = pd.concat([self.lastpredictions, unseendf], ignore_index=True)

        metrics = {
            'trainlosses': trainlosses,
            'testmse':     float(np.mean((preds - ytest) ** 2)),
            'testmae':     mae,
            'testrmse':    rmse,
            'ntrain':      len(xtrain),
            'ntest':       len(xtest),
        }

        return True, f"Training complete - Test RMSE: {rmse:.4f} MAE: {mae:.4f}", metrics

    def predict(self, masterdata: pd.DataFrame, reqsplits: List[pd.DataFrame]) -> Optional[np.ndarray]:
        if self.model is None or not self.featurecolumns:
            return None

        df = masterdata.copy()
        df['Region']        = df[countrycol].apply(self.determineregion)
        df['avgdailyusage'] = self.calculatedailyusage(df, reqsplits)

        for col, le in self.labelencoders.items():
            if col in df.columns:
                df[col] = le.transform(df[col].astype(str))

        x = df[self.featurecolumns].apply(pd.to_numeric, errors='coerce').fillna(0).values.astype(np.float32)
        xscaled = self.scaler.transform(x)

        self.model.eval()
        with torch.no_grad():
            return self.model(torch.tensor(xscaled, dtype=torch.float32)).numpy()

    def exportpredictions(self, filepath: str) -> Tuple[bool, str]:
        if self.lastpredictions is None:
            return False, "No predictions available. Run train() first."
        try:
            if filepath.endswith('.xlsx'):
                self.lastpredictions.to_excel(filepath, index_label='Part Index')
            else:
                self.lastpredictions.to_csv(filepath, index_label='Part Index')
            return True, f"Predictions exported to {filepath}"
        except Exception as e:
            return False, f"Export failed: {str(e)}"

    def savemodel(self, filepath: str) -> bool:
        if self.model is None:
            return False
        torch.save({
            'modelstate':     self.model.state_dict(),
            'scaler':         self.scaler,
            'labelencoders':  self.labelencoders,
            'featurecolumns': self.featurecolumns,
        }, filepath)
        return True

    def loadmodel(self, filepath: str, hiddensizes: List[int] = [128, 64, 32], dropout: float = 0.3) -> bool:
        try:
            checkpoint = torch.load(filepath, weights_only=False)
            self.featurecolumns  = checkpoint['featurecolumns']
            self.scaler          = checkpoint['scaler']
            self.labelencoders   = checkpoint['labelencoders']
            self.buildmodel(len(self.featurecolumns), hiddensizes, dropout)
            self.model.load_state_dict(checkpoint['modelstate'])
            return True
        except Exception:
            return False
