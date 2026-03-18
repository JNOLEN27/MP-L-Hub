"""
Real-data validation test for InventorybyPurposeNeuralNetwork.
Run from the repo root:  python test_nn.py
Uses the live imported files via DataImportManager (80/20 train-test split).
"""
import pandas as pd
import numpy as np
from ibp_neural_network import InventorybyPurposeNeuralNetwork
from import_manager import DataImportManager

# ---------------------------------------------------------------------------
# 1. Run the pipeline
# ---------------------------------------------------------------------------
print("=" * 60)
print("IBP Neural Network — real data validation")
print("=" * 60)

nn = InventorybyPurposeNeuralNetwork(DataImportManager())

ok, msg, data = nn.loadrequireddata()
print(f"\n[loadrequireddata]  ok={ok}  msg={msg}")
assert ok, "Data loading failed"

print(f"\n[master_data] {len(data['master_data'])} rows loaded")

print("\n[train] training model (100 epochs, 80/20 split)...")
ok, msg, metrics = nn.train(data, epochs=100, batchsize=32)
print(f"  ok={ok}  msg={msg}")
if not ok:
    raise SystemExit("Training failed.")

print(f"  Train samples : {metrics['ntrain']}")
print(f"  Test  samples : {metrics['ntest']}")
print(f"  Test RMSE     : {metrics['testrmse']:.4f}")
print(f"  Test MAE      : {metrics['testmae']:.4f}")
print(f"  Final train loss: {metrics['trainlosses'][-1]:.6f}")

# ---------------------------------------------------------------------------
# 2. Print test-split results (real part numbers vs model predictions)
# ---------------------------------------------------------------------------
testrows = nn.lastpredictions[nn.lastpredictions['Split'] == 'test'].copy()
testrows = testrows.sort_values('AbsError', ascending=False)

print(f"\n[test split — {len(testrows)} parts]")
print(f"{'PART':<15} {'Actual':>10} {'Predicted':>10} {'Error':>10} {'AbsError':>10}")
print("-" * 57)
for _, row in testrows.iterrows():
    print(f"{str(row['PART']):<15} {row['Actual']:>10.1f} {row['Predicted']:>10.1f} {row['Error']:>10.1f} {row['AbsError']:>10.1f}")

print("\n" + "=" * 60)
print("Validation complete.")
print("=" * 60)
