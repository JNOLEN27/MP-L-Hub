"""
Quick smoke test for InventorybyPurposeNeuralNetwork.
Run from the repo root:  python test_nn.py
No real imported files needed — uses synthetic data matching the real schema.
"""
import pandas as pd
import numpy as np
from ibp_neural_network import InventorybyPurposeNeuralNetwork

# ---------------------------------------------------------------------------
# 1. Build synthetic data that matches the real column schemas
# ---------------------------------------------------------------------------
np.random.seed(42)
N = 200  # number of parts

COUNTRIES = ['USA', 'MEXICO', 'GERMANY', 'FRANCE', 'CHINA', 'SOUTH KOREA']
SCC_NAMES = ['SCC_A', 'SCC_B', 'SCC_C']
CONSOLIDATORS = ['CON_1', 'CON_2', 'CON_3']

master_data = pd.DataFrame({
    'PART':               [f'PART_{i:04d}' for i in range(N)],
    'PART_DESC':          [f'Description {i}' for i in range(N)],
    'SUPP_NAME':          [f'Supplier_{i % 10}' for i in range(N)],
    'SUPP_SHP':           [f'SHIP_{i % 5}' for i in range(N)],
    'SUPP_SHP_COUNTRY':   np.random.choice(COUNTRIES, N),
    'SCC_NAME':           np.random.choice(SCC_NAMES, N),
    'PRICE':              np.random.uniform(1, 500, N).round(2),
    'UNIT_LOAD_QTY':      np.random.choice([1, 5, 10, 20, 50], N).astype(float),
    'MULT_UNIT_LOAD_VALID': np.random.choice([0, 1], N).astype(float),
    'SHIP_QTY':           np.random.choice([10, 20, 50, 100], N).astype(float),
    'STOCK':              np.random.randint(0, 500, N).astype(float),
    'SAFETY':             np.random.randint(5, 150, N).astype(float),   # target
    'FIXED_PERIOD':       np.random.choice([1, 7, 14, 30], N).astype(float),
    'TTT_DAYS':           np.random.randint(1, 60, N).astype(float),
    'CONSOLIDATOR':       np.random.choice(CONSOLIDATORS, N),
})

# Req splits: long format (ARTNR, PRODDAG, SKIFT, ARTAN)
# Cover ~120 days from today so all region windows have data
today = pd.Timestamp.today().normalize()
dates = pd.date_range(today, periods=120, freq='D')

def make_req_split(parts, dates, seed):
    rng = np.random.default_rng(seed)
    rows = []
    for part in parts:
        for date in dates:
            for shift in ['A', 'B']:
                rows.append({
                    'ARTNR':            part,
                    'PART_DESCRIPTION': f'Desc {part}',
                    'PRODDAG':          date,
                    'WEEK':             date.isocalendar()[1],
                    'SKIFT':            shift,
                    'ARTAN':            float(rng.integers(0, 30)),
                })
    return pd.DataFrame(rows)

parts = master_data['PART'].tolist()
req1 = make_req_split(parts[:70],  dates, seed=1)
req2 = make_req_split(parts[70:140], dates, seed=2)
req3 = make_req_split(parts[140:],  dates, seed=3)

# ---------------------------------------------------------------------------
# 2. Build a mock import_manager that returns the synthetic data
# ---------------------------------------------------------------------------
class MockImportManager:
    def __init__(self, master, r1, r2, r3):
        self._data = {
            'master_data':              master,
            'part_requirement_split_1': r1,
            'part_requirement_split_2': r2,
            'part_requirement_split_3': r3,
        }
    def loaddata(self, category):
        return self._data.get(category, pd.DataFrame())

# ---------------------------------------------------------------------------
# 3. Run the pipeline
# ---------------------------------------------------------------------------
print("=" * 60)
print("IBP Neural Network — smoke test")
print("=" * 60)

nn = InventorybyPurposeNeuralNetwork(MockImportManager(master_data, req1, req2, req3))

# Load data
ok, msg, data = nn.loadrequireddata()
print(f"\n[loadrequireddata]  ok={ok}  msg={msg}")
assert ok, "Data loading failed"

# Test region mapping
print("\n[determineregion] spot checks:")
for country, expected in [('USA', 'USA'), ('MEXICO', 'MEX'), ('GERMANY', 'EMEA'), ('CHINA', 'APAC'), ('MARS', 'Country not mapped. Reach out to Admin')]:
    result = nn.determineregion(country)
    status = "✓" if result == expected else "✗"
    print(f"  {status}  {country:10s} → {result}")

# Test daily usage calculation
print("\n[calculatedailyusage] sample of computed values:")
usage = nn.calculatedailyusage(master_data.head(10), [req1, req2, req3])
for part, val in zip(master_data['PART'].head(10), usage):
    country = master_data.loc[master_data['PART'] == part, 'SUPP_SHP_COUNTRY'].values[0]
    region  = nn.determineregion(country)
    print(f"  {part}  country={country:12s}  region={region:4s}  avg_daily={val:.2f}")

# Train
print("\n[train] training model (50 epochs)...")
ok, msg, metrics = nn.train(data, epochs=50, batchsize=32)
print(f"  ok={ok}  msg={msg}")
if ok:
    print(f"  Train samples : {metrics['ntrain']}")
    print(f"  Test  samples : {metrics['ntest']}")
    print(f"  Test RMSE     : {metrics['testrmse']:.4f}")
    print(f"  Test MAE      : {metrics['testmae']:.4f}")
    print(f"  Final train loss: {metrics['trainlosses'][-1]:.6f}")

# Predict on a few rows
print("\n[predict] predictions vs actuals (first 5 test parts):")
preds = nn.predict(master_data.head(5), [req1, req2, req3])
if preds is not None:
    for i, (pred, actual) in enumerate(zip(preds, master_data['SAFETY'].head(5))):
        print(f"  Part {i}  predicted={pred:.1f}  actual={actual:.1f}")

# Save / load round-trip
import tempfile, os
print("\n[savemodel / loadmodel] round-trip test...")
with tempfile.NamedTemporaryFile(suffix='.pt', delete=False) as f:
    tmppath = f.name
try:
    saved = nn.savemodel(tmppath)
    print(f"  save: {saved}")

    nn2 = InventorybyPurposeNeuralNetwork(MockImportManager(master_data, req1, req2, req3))
    loaded = nn2.loadmodel(tmppath)
    print(f"  load: {loaded}")

    preds2 = nn2.predict(master_data.head(5), [req1, req2, req3])
    match = np.allclose(preds, preds2, atol=1e-4)
    print(f"  predictions match after reload: {match}")
finally:
    os.unlink(tmppath)

print("\n" + "=" * 60)
print("Smoke test complete.")
print("=" * 60)
