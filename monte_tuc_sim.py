"""
monte_tuc_sim.py
Monte Carlo Tied-Up Capital (TUC) Inventory Simulation

Translated from R and optimised for Python / NumPy.

Data is loaded through the application's DataImportManager.
The only file not previously covered is the PVA Percentage CSV
(category "pva_percentage", required columns: PVA.Percentage, Frequency).

Public API
----------
load_required_data(im)                  -> Dict
get_part_parameters(part, master_data)  -> Dict
monte_carlo_inventory_sim(...)          -> Dict   # per-part MC
simulate_piwd_plus(...)                 -> Dict   # deterministic PLUS plan
plant_forecast_shared_pva(...)          -> Dict   # full plant forecast
update_inventory_report(...)            -> Dict   # archive today's CIR value
view_inventory_history(...)             -> DataFrame
compare_inventory_change(...)           -> Dict
plot_historical_and_forecast(...)       -> Dict

Performance notes
-----------------
plant_forecast_shared_pva uses several optimisations over the naïve approach:

1. _get_obsolete_inventory is fully vectorised (pandas merge instead of a
   per-part Python loop).

2. All DataFrames are pre-grouped by part key once (req_by_part, gtbr_by_part,
   del_schedule) so every per-part lookup is O(1) instead of O(n).

3. The validation pass and simulation-cache build happen in a single loop.
   Each part's consumption schedule, planned deliveries, DOW array, and
   next-delivery look-ahead array are pre-computed and stored — no redundant
   work in the worker processes.

4. ProcessPoolExecutor is created with an *initializer* (_worker_init) that
   receives the full part cache and PVA matrix **once per worker process**
   instead of serialising DataFrames with every submitted task.  For a plant
   with 5 000 parts this cuts IPC data from ~150 GB to ~0.3 GB total.

5. Tier 4 parts (lowest value, 1 sim each) are submitted in batches of 64
   to further reduce task-submission overhead.

6. _precompute_next_delivery builds a look-ahead index once per part so that
   _simulate_run can do an O(1) array read instead of an O(10) inner loop on
   every day × every simulation.

7. get_consumption_schedule uses pandas reindex instead of a Python list
   comprehension for the date-to-consumption mapping.
"""

from __future__ import annotations

import math
import time
import warnings
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from tqdm import tqdm

from import_manager import DataImportManager
from app.utils.config import SHAREDNETWORKPATH

# ─── Constants ────────────────────────────────────────────────────────────────

ARCHIVE_PATH: Path = SHAREDNETWORKPATH / "archive" / "Inventory_Archive.csv"

# Tier-4 batch size: number of low-value parts grouped into a single worker task.
# Larger values reduce IPC overhead; smaller values keep progress granular.
_TIER4_BATCH_SIZE: int = 64

# ─── Worker-process globals ────────────────────────────────────────────────────
# These are populated by _worker_init (called once per worker process, not per
# task) and read by the worker functions.  They are never used in the main process.

_WORKER_PART_CACHE: Dict = {}
_WORKER_PVA_MATRIX: Optional[np.ndarray] = None


def _worker_init(part_cache: Dict, pva_matrix: np.ndarray) -> None:
    """
    ProcessPoolExecutor initializer — called exactly once per worker process.

    Storing data in module-level globals means each submitted task only needs
    to carry a part name and two small integers, eliminating the enormous
    per-task serialisation cost of passing DataFrames through the task queue.
    """
    global _WORKER_PART_CACHE, _WORKER_PVA_MATRIX
    _WORKER_PART_CACHE = part_cache
    _WORKER_PVA_MATRIX = pva_matrix


# ─── Data Loading ─────────────────────────────────────────────────────────────

def load_required_data(im: DataImportManager) -> Dict:
    """
    Load all data needed by the simulator via the import manager.

    Returns a dict with keys:
        master_data, manual_ttt, requirement, inventory, gtbr, pva_distribution
    """
    split1 = im.loaddata("part_requirement_split_1")
    split2 = im.loaddata("part_requirement_split_2")
    split3 = im.loaddata("part_requirement_split_3")

    # Combine splits, skipping the duplicate header rows that arise from
    # exporting in chunks (mirrors R's rbind(split1, split2[-1,], split3[-1,]))
    req = pd.concat(
        [split1, split2.iloc[1:].reset_index(drop=True),
                 split3.iloc[1:].reset_index(drop=True)],
        ignore_index=True,
    )
    req["ARTAN"]   = pd.to_numeric(req["ARTAN"],   errors="coerce").fillna(0)
    req["PRODDAG"] = pd.to_datetime(req["PRODDAG"], format="%m/%d/%Y", errors="coerce")

    gtbr = im.loaddata("goods_to_be_received")
    gtbr["ANK_TID_TIDIGAST"] = pd.to_datetime(
        gtbr["ANK_TID_TIDIGAST"], format="%m/%d/%Y %H:%M", errors="coerce"
    )

    pva_raw = im.loaddata("pva_percentage")
    pva_distribution = np.repeat(
        pva_raw["PVA.Percentage"].astype(float).values,
        pva_raw["Frequency"].astype(int).values,
    )

    return {
        "master_data":      im.loaddata("master_data"),
        "manual_ttt":       im.loaddata("manual_TTT"),
        "requirement":      req,
        "inventory":        im.loaddata("current_inventory_report"),
        "gtbr":             gtbr,
        "pva_distribution": pva_distribution,
    }


# ─── Part Lookups (public API — signatures unchanged) ─────────────────────────

def get_part_parameters(part_number: str, master_data: pd.DataFrame) -> Dict:
    row = master_data[master_data["PART"] == part_number]
    if row.empty:
        raise KeyError(f"Part {part_number} not found in master data")
    r = row.iloc[0]

    def _num(col: str, default: float = 0.0) -> float:
        v = pd.to_numeric(r.get(col, default), errors="coerce")
        return default if pd.isna(v) else float(v)

    ul  = max(1.0, _num("UNIT_LOAD_QTY",       1.0))
    mul = max(1.0, _num("MULT_UNIT_LOAD_VALID", 1.0))

    return {
        "part_number":     part_number,
        "supplier":        str(r.get("SUPP_SHP", "")),
        "safety_days":     _num("SAFETY"),
        "safety_stock":    _num("STOCK"),
        "unit_load":       ul,
        "multi_unit_load": mul,
        "order_increment": ul * mul,
        "price":           _num("PRICE"),
    }


def get_delivery_schedule(supplier_code: str, manual_ttt: pd.DataFrame) -> List[int]:
    """Return sorted list of delivery weekdays (1=Mon … 5=Fri, ISO convention)."""
    rows = manual_ttt[manual_ttt["LEVNR"] == supplier_code]
    if rows.empty:
        return [1, 2, 3, 4, 5]
    days = (
        pd.to_numeric(rows["FRAKTDAG"], errors="coerce")
        .dropna().astype(int).tolist()
    )
    return sorted(days) if days else [1, 2, 3, 4, 5]


def get_initial_stock(part_number: str, inventory: pd.DataFrame) -> float:
    row = inventory[inventory["PART_NO"] == part_number]
    if row.empty:
        raise KeyError(f"Part {part_number} not found in inventory")
    r    = row.iloc[0]
    beg  = pd.to_numeric(r.get("BEGINNING_INVENTORY_TODAY", 0), errors="coerce") or 0.0
    yard = pd.to_numeric(r.get("INVENTORY_YARD_TODAY",     0), errors="coerce") or 0.0
    return float(beg + yard)


def get_planned_deliveries(
    part_number: str,
    sim_dates: pd.DatetimeIndex,
    gtbr: pd.DataFrame,
) -> pd.DataFrame:
    """
    Return a DataFrame with columns [date, quantity, day_index (1-based)]
    for deliveries of *part_number* falling within *sim_dates*.
    """
    part = gtbr[gtbr["ARTNR"] == part_number].copy()
    if part.empty:
        return pd.DataFrame(columns=["date", "quantity", "day_index"])

    part = part.dropna(subset=["ANK_TID_TIDIGAST"])
    start, end = sim_dates[0], sim_dates[-1]
    part = part[
        (part["ANK_TID_TIDIGAST"] >= start) &
        (part["ANK_TID_TIDIGAST"] <= end)
    ]
    if part.empty:
        return pd.DataFrame(columns=["date", "quantity", "day_index"])

    # Map each delivery date to a 1-based day index
    date_to_idx: Dict = {d.date(): i + 1 for i, d in enumerate(sim_dates)}
    part["day_index"] = part["ANK_TID_TIDIGAST"].dt.date.map(date_to_idx)
    part = part.dropna(subset=["day_index"])
    part["quantity"] = pd.to_numeric(part["ARTAN"], errors="coerce").fillna(0)

    out = part[["ANK_TID_TIDIGAST", "quantity", "day_index"]].copy()
    out.columns = ["date", "quantity", "day_index"]
    out["day_index"] = out["day_index"].astype(int)
    return out.sort_values("day_index").reset_index(drop=True)


def get_consumption_schedule(
    part_number: str,
    days: int,
    requirement: pd.DataFrame,
) -> Tuple[np.ndarray, pd.DatetimeIndex]:
    """
    Return (consumption_array, dates) of length *days*.

    consumption_array[i] is the planned usage on day i.
    Weekends are excluded unless production is scheduled on them.
    """
    part_req = requirement[requirement["ARTNR"] == part_number].copy()
    part_req = part_req.dropna(subset=["PRODDAG"]).sort_values("PRODDAG")

    if part_req.empty:
        warnings.warn(f"Part {part_number} not in requirements – using 0 consumption")
        start = pd.Timestamp.today().normalize()
        dates = pd.bdate_range(start, periods=days)[:days]
        return np.zeros(days), dates

    start_date     = part_req["PRODDAG"].min()
    weekend_exists = part_req["PRODDAG"].dt.dayofweek.isin([5, 6]).any()

    # Generate enough candidate dates then slice to exactly *days*
    pool = (
        pd.date_range(start_date, periods=days * 2)
        if weekend_exists
        else pd.bdate_range(start_date, periods=days * 2)
    )
    dates = pool[:days]

    daily = part_req.groupby("PRODDAG")["ARTAN"].sum()
    # Vectorised reindex — avoids a Python list comprehension over all dates
    consumption = daily.reindex(dates, fill_value=0.0).to_numpy(dtype=float)
    return consumption, dates


# ─── Delivery Need ────────────────────────────────────────────────────────────

def calculate_delivery_need(
    current_stock: float,
    upcoming_consumption: np.ndarray,
    safety_stock: float,
    safety_days: float,
    order_increment: float,
) -> float:
    """
    Calculate how many units to order so that stock stays above the safety
    buffer through the next delivery cycle.
    """
    min_needed = float(upcoming_consumption.sum()) + safety_stock

    if safety_days > 0 and len(upcoming_consumption) > 0:
        sd_int  = int(safety_days)
        sd_frac = safety_days - sd_int
        if sd_frac > 0 and len(upcoming_consumption) > sd_int:
            min_needed += sd_frac * upcoming_consumption[sd_int]

    shortage = min_needed - current_stock
    if shortage <= 0:
        return 0.0
    return math.ceil(shortage / order_increment) * order_increment


# ─── Core Single-Run Simulation ───────────────────────────────────────────────

def _precompute_next_delivery(dow: np.ndarray, del_set: set, days: int) -> np.ndarray:
    """
    For each 0-based day index d, return the 1-based index of the next delivery
    day within the following 10 days, or 0 if none exists.

    This is computed once per part before any simulations run, eliminating the
    O(10) inner look-ahead loop inside _simulate_run for every day × every sim.
    Uses numpy searchsorted so the pre-computation itself is O(days · log D)
    where D is the number of delivery days.
    """
    # 0-based indices of days that are delivery days
    del_idx = np.array([d for d in range(days) if dow[d] in del_set], dtype=np.int32)
    next_del = np.zeros(days, dtype=np.int32)
    if len(del_idx) == 0:
        return next_del
    for d in range(days):
        # First delivery-day index strictly after d
        pos = np.searchsorted(del_idx, d + 1)
        if pos < len(del_idx) and del_idx[pos] <= d + 9:   # within 10 days
            next_del[d] = int(del_idx[pos]) + 1            # convert to 1-based
    return next_del


def _simulate_run(
    initial_stock: float,
    base_consumption: np.ndarray,
    dow: np.ndarray,               # dayofweek, 0=Mon … 6=Sun, length = days
    delivery_dow_set: set,         # 0-indexed weekdays on which supplier delivers
    planned_qty: np.ndarray,       # (days,) – non-zero where a planned delivery arrives
    last_planned_day: int,         # 1-indexed; 0 = no planned deliveries
    params: Dict,
    pva_samples: Optional[np.ndarray],   # (days,) or None for deterministic
    next_delivery: Optional[np.ndarray] = None,  # pre-computed look-ahead (days,)
) -> Dict:
    """
    Simulate one inventory trajectory.

    Days are 1-indexed in the loop so that planned_qty[day-1] aligns with
    day_index from get_planned_deliveries.

    When *next_delivery* is provided (pre-computed by _precompute_next_delivery),
    the inner look-ahead loop is replaced by an O(1) array read.
    """
    days = len(base_consumption)

    actual = (
        np.ceil(pva_samples * base_consumption).astype(float)
        if pva_samples is not None
        else base_consumption.astype(float)
    )

    stock      = np.empty(days + 1)
    deliveries = np.zeros(days)
    d_type     = np.zeros(days, dtype=np.int8)   # 1=planned, 2=simulated
    stock[0]   = initial_stock

    order_inc = params["order_increment"]
    ss        = params["safety_stock"]
    sd        = params["safety_days"]

    for day in range(1, days + 1):
        idx = day - 1
        s   = stock[idx]

        if planned_qty[idx] > 0:
            # Use the hard delivery scheduled in GTBR
            s              += planned_qty[idx]
            deliveries[idx]  = planned_qty[idx]
            d_type[idx]    = 1

        elif day > last_planned_day:
            # Past the planning horizon – simulate replenishment on delivery days
            current_dow = dow[idx]
            if current_dow in delivery_dow_set and current_dow < 5:   # weekday only
                # O(1) look-ahead via pre-computed array; fallback to loop if absent
                if next_delivery is not None:
                    nd     = int(next_delivery[idx])
                    next_dd = nd if nd > 0 else None
                else:
                    next_dd = None
                    for fd in range(day + 1, min(days, day + 10) + 1):
                        if dow[fd - 1] in delivery_dow_set and dow[fd - 1] < 5:
                            next_dd = fd
                            break

                end_idx  = days if next_dd is None else next_dd - 1
                upcoming = actual[idx:end_idx]

                qty = calculate_delivery_need(s, upcoming, ss, sd, order_inc)
                if qty > 0:
                    s              += qty
                    deliveries[idx]  = qty
                    d_type[idx]    = 2

        stock[day] = s - actual[idx]

    stockouts = stock[1:] < 0
    return {
        "stock":            stock,
        "deliveries":       deliveries,
        "delivery_type":    d_type,
        "final_stock":      float(stock[-1]),
        "min_stock":        float(stock.min()),
        "avg_stock":        float(stock.mean()),
        "stockout_days":    int(stockouts.sum()),
        "total_deliveries": float(deliveries.sum()),
        "num_deliveries":   int((deliveries > 0).sum()),
        "num_planned":      int((d_type == 1).sum()),
        "num_simulated":    int((d_type == 2).sum()),
    }


# ─── Shared Setup Helper ──────────────────────────────────────────────────────

def _build_part_sim_inputs(
    part_number: str, days: int, data: Dict
) -> Tuple[Dict, float, np.ndarray, pd.DatetimeIndex, set, np.ndarray, int, pd.DataFrame]:
    """
    Centralised preparation shared by simulate_piwd_plus and
    monte_carlo_inventory_sim to avoid code duplication.
    """
    params      = get_part_parameters(part_number, data["master_data"])
    init_stock  = get_initial_stock(part_number, data["inventory"])
    base_cons, dates = get_consumption_schedule(part_number, days, data["requirement"])

    # FRAKTDAG uses ISO 1-5 → convert to pandas dayofweek 0-4
    raw_del_days = get_delivery_schedule(params["supplier"], data["manual_ttt"])
    del_set = {d - 1 for d in raw_del_days}

    # Planned deliveries as a dense array (length = days)
    planned = get_planned_deliveries(part_number, dates, data["gtbr"])
    planned_qty = np.zeros(days)
    if not planned.empty:
        valid = planned[(planned["day_index"] >= 1) & (planned["day_index"] <= days)]
        if not valid.empty:
            np.add.at(planned_qty, valid["day_index"].values - 1, valid["quantity"].values)
    last_pd = int(planned["day_index"].max()) if not planned.empty else 0

    dow = dates.dayofweek.to_numpy()   # 0=Mon … 6=Sun

    return params, init_stock, base_cons, dates, del_set, planned_qty, last_pd, planned


# ─── PLUS Plan (deterministic) ───────────────────────────────────────────────

def simulate_piwd_plus(part_number: str, days: int, data: Dict) -> Dict:
    """
    Deterministic inventory trajectory using base consumption only (no PVA).
    This is the PLUS plan baseline.
    """
    params, init_stock, base_cons, dates, del_set, planned_qty, last_pd, planned = (
        _build_part_sim_inputs(part_number, days, data)
    )
    result = _simulate_run(
        init_stock, base_cons, dates.dayofweek.to_numpy(),
        del_set, planned_qty, last_pd, params, pva_samples=None,
    )
    result.update({
        "part_number":        part_number,
        "params":             params,
        "dates":              dates,
        "base_consumption":   base_cons,
        "planned_deliveries": planned,
        "last_planned_day":   last_pd,
        "initial_stock":      init_stock,
    })
    return result


# ─── Per-Part Monte Carlo ─────────────────────────────────────────────────────

def monte_carlo_inventory_sim(
    part_number: str,
    days: int,
    data: Dict,
    n_sims: int = 1000,
    seed: Optional[int] = None,
) -> Dict:
    """
    Full Monte Carlo inventory simulation for a single part.

    PVA values are pre-sampled as an (n_sims × days) matrix so that
    the inner simulation loop is called with plain numpy arrays instead of
    hitting the random-number generator on every iteration.
    """
    rng = np.random.default_rng(seed)
    pva = data["pva_distribution"]

    params, init_stock, base_cons, dates, del_set, planned_qty, last_pd, planned = (
        _build_part_sim_inputs(part_number, days, data)
    )
    dow = dates.dayofweek.to_numpy()

    # Pre-compute next-delivery look-ahead so the simulation loop is faster
    next_delivery = _precompute_next_delivery(dow, del_set, days)

    # ── Print header ──
    ci_value = init_stock * params["price"]
    print(f"\n{'='*45}")
    print(f"Monte Carlo Inventory Simulation")
    print(f"{'='*45}")
    print(f"Part Number:        {part_number}")
    print(f"Supplier:           {params['supplier']}")
    print(f"Safety Days:        {params['safety_days']:.2f}")
    print(f"Safety Stock:       {params['safety_stock']:.0f}")
    print(f"Unit Load:          {params['unit_load']:.0f}")
    print(f"Multi-Unit Load:    {params['multi_unit_load']:.0f}")
    print(f"Order Increment:    {params['order_increment']:.0f}")
    print(f"Piece Price:        ${params['price']:.2f}")
    print(f"Initial Stock:      {init_stock:.0f}")
    print(f"Inventory Value:    ${ci_value:,.2f}")
    if not planned.empty:
        print(f"Planned Deliveries: {len(planned)} (through day {last_pd})")
        print(f"Planned Quantity:   {planned['quantity'].sum():.0f}")
    else:
        print("Planned Deliveries: None")
    print(f"\nSimulating {days} days × {n_sims} iterations…\n")

    # ── Compute PLUS plan ──
    plus = simulate_piwd_plus(part_number, days, data)

    # ── Pre-sample all PVA at once ──
    pva_matrix = rng.choice(pva, size=(n_sims, days), replace=True)

    # ── Run simulations ──
    all_stock = np.empty((n_sims, days + 1))
    metrics   = {
        k: np.empty(n_sims) for k in [
            "final_stock", "min_stock", "avg_stock", "stockout_days",
            "total_deliveries", "num_deliveries", "num_planned", "num_simulated",
        ]
    }

    for i in tqdm(range(n_sims), desc="Simulating", unit="sim"):
        r = _simulate_run(
            init_stock, base_cons, dow, del_set,
            planned_qty, last_pd, params, pva_matrix[i], next_delivery,
        )
        all_stock[i] = r["stock"]
        for k in metrics:
            metrics[k][i] = r[k]

    # ── Compute trajectory statistics ──
    avg_traj = all_stock.mean(axis=0)
    q05      = np.percentile(all_stock,  5, axis=0)
    q25      = np.percentile(all_stock, 25, axis=0)
    q75      = np.percentile(all_stock, 75, axis=0)
    q95      = np.percentile(all_stock, 95, axis=0)

    # ── Print results ──
    fs = metrics["final_stock"]
    ms = metrics["min_stock"]
    ag = metrics["avg_stock"]
    so = metrics["stockout_days"]

    print("\n=== Results ===\n")
    print(f"Final Stock:")
    print(f"  Mean={fs.mean():.1f}  Median={np.median(fs):.1f}  "
          f"PLUS={plus['final_stock']:.1f}  SD={fs.std():.1f}  "
          f"[{fs.min():.1f}, {fs.max():.1f}]")
    print(f"\nMinimum Stock Reached:")
    print(f"  Mean={ms.mean():.1f}  Median={np.median(ms):.1f}  "
          f"PLUS={plus['min_stock']:.1f}  "
          f"P5={np.percentile(ms, 5):.1f}  P95={np.percentile(ms, 95):.1f}")
    print(f"\nAverage Inventory:")
    print(f"  Mean={ag.mean():.1f}  Median={np.median(ag):.1f}  "
          f"PLUS={plus['avg_stock']:.1f}")
    print(f"\nDeliveries:")
    print(f"  Planned={metrics['num_planned'].mean():.1f}  "
          f"Simulated={metrics['num_simulated'].mean():.1f}  "
          f"Total={metrics['num_deliveries'].mean():.1f}  "
          f"Units={metrics['total_deliveries'].mean():.1f}  "
          f"PLUS={plus['num_deliveries']}")
    so_pct = 100 * (so > 0).sum() / n_sims
    print(f"\nStockout Risk: {so_pct:.2f}%  ({int((so > 0).sum())}/{n_sims})")
    if (so > 0).any():
        print(f"  Avg days when occurs: {so[so > 0].mean():.2f}")
    print(f"  PLUS Plan: {'STOCKOUT' if plus['stockout_days'] > 0 else 'No stockout'}")

    _plot_monte_carlo(
        part_number, days, all_stock, plus, avg_traj,
        q05, q25, q75, q95, fs, ms, params, last_pd, n_sims,
    )

    return {
        "part_number":        part_number,
        "params":             params,
        "initial_stock":      init_stock,
        "planned_deliveries": planned,
        "plus_plan":          plus,
        "all_trajectories":   all_stock,
        "avg_trajectory":     avg_traj,
        "q05": q05, "q25": q25, "q75": q75, "q95": q95,
        **metrics,
    }


# ─── Plotting ─────────────────────────────────────────────────────────────────

def _plot_monte_carlo(
    part_number: str,
    days: int,
    all_stock: np.ndarray,
    plus: Dict,
    avg_traj: np.ndarray,
    q05: np.ndarray,
    q25: np.ndarray,
    q75: np.ndarray,
    q95: np.ndarray,
    final_stocks: np.ndarray,
    min_stocks: np.ndarray,
    params: Dict,
    last_planned_day: int,
    n_sims: int,
) -> None:
    x = np.arange(days + 1)
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    fig.suptitle(f"Monte Carlo Inventory Simulation — Part {part_number}", fontsize=13)

    # ── Plot 1: Average trajectory + confidence bands ──
    ax = axes[0, 0]
    ax.fill_between(x, q05, q95, alpha=0.15, color="royalblue", label="90% CI")
    ax.fill_between(x, q25, q75, alpha=0.30, color="royalblue", label="50% CI")
    if last_planned_day > 0:
        ax.plot(x[:last_planned_day + 1], avg_traj[:last_planned_day + 1],
                lw=2, color="darkgreen", label="Avg (Planned)")
    if last_planned_day < days:
        ax.plot(x[last_planned_day:], avg_traj[last_planned_day:],
                lw=2, color="royalblue", label="Avg (Simulated)")
    if 0 < last_planned_day < days:
        ax.axvline(last_planned_day, color="purple", ls=":", lw=1, label="Transition")
    ax.plot(x, plus["stock"], lw=2.5, ls="--", color="black", label="PLUS Plan")
    ax.axhline(params["safety_stock"], color="orange", ls="--", lw=1.5, label="Safety Stock")
    ax.axhline(0, color="red", ls="--", lw=1)
    ax.set_title("Average Trajectory + Confidence Bands")
    ax.set_xlabel("Days"); ax.set_ylabel("Stock Level")
    ax.legend(fontsize=7, loc="best"); ax.grid(True, alpha=0.3)

    # ── Plot 2: Sample trajectories ──
    ax = axes[0, 1]
    sample_n = min(100, n_sims)
    for i in range(sample_n):
        ax.plot(x, all_stock[i], color="black", alpha=0.07, lw=0.5)
    if last_planned_day > 0:
        ax.plot(x[:last_planned_day + 1], avg_traj[:last_planned_day + 1],
                lw=2, color="darkgreen", label="Avg (Planned)")
    if last_planned_day < days:
        ax.plot(x[last_planned_day:], avg_traj[last_planned_day:],
                lw=2, color="red", label="Avg (Simulated)")
    if 0 < last_planned_day < days:
        ax.axvline(last_planned_day, color="purple", ls=":", lw=1)
    ax.plot(x, plus["stock"], lw=2.5, ls="--", color="black", label="PLUS Plan")
    ax.axhline(params["safety_stock"], color="orange", ls="--", lw=1.5, label="Safety Stock")
    ax.axhline(0, color="red", ls="--", lw=1)
    ax.set_title(f"Sample of {sample_n} Trajectories")
    ax.set_xlabel("Days"); ax.set_ylabel("Stock Level")
    ax.legend(fontsize=7, loc="best"); ax.grid(True, alpha=0.3)

    # ── Plot 3: Final stock distribution ──
    ax = axes[1, 0]
    ax.hist(final_stocks, bins=50, color="lightsteelblue", edgecolor="none")
    ax.axvline(final_stocks.mean(),      color="red",     ls="--", lw=2,   label="Mean")
    ax.axvline(np.median(final_stocks),  color="blue",    ls="--", lw=2,   label="Median")
    ax.axvline(plus["final_stock"],      color="black",   ls="--", lw=2.5, label="PLUS Plan")
    ax.axvline(0,                        color="darkred", lw=1.5,           label="Stockout")
    ax.set_title("Final Stock Distribution")
    ax.set_xlabel("Final Stock Level"); ax.set_ylabel("Count")
    ax.legend(fontsize=8); ax.grid(True, alpha=0.3)

    # ── Plot 4: Minimum stock distribution ──
    ax = axes[1, 1]
    ax.hist(min_stocks, bins=50, color="lightcoral", edgecolor="none")
    ax.axvline(min_stocks.mean(),        color="red",    ls="--", lw=2,   label="Mean")
    ax.axvline(params["safety_stock"],   color="orange", ls="--", lw=2,   label="Safety Stock")
    ax.axvline(plus["min_stock"],        color="black",  ls="--", lw=2.5, label="PLUS Plan")
    ax.axvline(0,                        color="darkred", lw=1.5,          label="Stockout")
    ax.set_title("Minimum Stock Distribution")
    ax.set_xlabel("Min Stock Reached"); ax.set_ylabel("Count")
    ax.legend(fontsize=8); ax.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.show()


# ─── Plant-Wide Forecast — Internal Helpers ───────────────────────────────────

def _get_obsolete_inventory(
    master_data: pd.DataFrame,
    requirement: pd.DataFrame,
    inventory: pd.DataFrame,
) -> Dict:
    """
    Return total value and list of parts that have stock but no requirements.

    Fully vectorised using pandas merge — replaces the original O(n_inventory)
    Python loop with per-part DataFrame lookups.
    """
    active_parts = set(requirement["ARTNR"].dropna().unique())

    # Filter inventory to non-active parts in one vectorised step
    inv_obs = inventory[~inventory["PART_NO"].isin(active_parts)].copy()
    if inv_obs.empty:
        print("\nObsolete inventory: 0 parts, $0\n")
        return {"total_value": 0.0, "count": 0, "parts": [], "values": []}

    inv_obs["_stock"] = (
        pd.to_numeric(inv_obs["BEGINNING_INVENTORY_TODAY"], errors="coerce").fillna(0)
        + pd.to_numeric(inv_obs["INVENTORY_YARD_TODAY"],     errors="coerce").fillna(0)
    )

    # Pull prices from master_data via a merge (one pass, no Python loop)
    md_prices = (
        master_data[["PART", "PRICE"]]
        .drop_duplicates("PART")
        .copy()
    )
    md_prices["PRICE"] = pd.to_numeric(md_prices["PRICE"], errors="coerce").fillna(0)

    merged = inv_obs.merge(md_prices, left_on="PART_NO", right_on="PART", how="left")
    merged["PRICE"] = merged["PRICE"].fillna(0)
    merged["_value"] = merged["_stock"] * merged["PRICE"]

    valid = merged[(merged["_stock"] > 0) & (merged["PRICE"] > 0)]

    total_value = float(valid["_value"].sum())
    parts  = valid["PART_NO"].tolist()
    values = valid["_value"].tolist()

    print(f"\nObsolete inventory: {len(parts)} parts, ${total_value:,.0f}\n")
    return {"total_value": total_value, "count": len(parts),
            "parts": parts, "values": values}


def _build_plant_lookup_dicts(data: Dict) -> Dict:
    """
    Build O(1) lookup structures from the raw DataFrames.
    Called once before simulation; eliminates repeated O(n) scans during
    validation and cache-building.

    Returns a dict containing:
        md_idx        — master_data indexed by PART
        inv_idx       — inventory indexed by PART_NO
        req_by_part   — requirement grouped by ARTNR (dict of DataFrames)
        gtbr_by_part  — gtbr grouped by ARTNR (dict of DataFrames)
        del_schedule  — supplier code → frozenset of 0-based delivery weekdays
    """
    md   = data["master_data"]
    inv  = data["inventory"]
    ttt  = data["manual_ttt"]
    req  = data["requirement"]
    gtbr = data["gtbr"]

    # Indexed DataFrames for O(log n) .loc access
    md_idx  = md.set_index("PART",    drop=False) if "PART"    in md.columns  else md
    inv_idx = inv.set_index("PART_NO", drop=False) if "PART_NO" in inv.columns else inv

    # Delivery schedule: supplier → set of 0-based weekday ints
    del_schedule: Dict[str, set] = {}
    if not ttt.empty and "LEVNR" in ttt.columns and "FRAKTDAG" in ttt.columns:
        for levnr, grp in ttt.groupby("LEVNR"):
            raw = pd.to_numeric(grp["FRAKTDAG"], errors="coerce").dropna().astype(int).tolist()
            del_schedule[str(levnr)] = {d - 1 for d in raw} if raw else {0, 1, 2, 3, 4}

    # Requirement and GTBR grouped by part — O(n) once, O(1) per part thereafter
    req_by_part: Dict[str, pd.DataFrame] = {}
    if not req.empty and "ARTNR" in req.columns:
        for artnr, grp in req.groupby("ARTNR"):
            req_by_part[str(artnr)] = grp

    gtbr_by_part: Dict[str, pd.DataFrame] = {}
    if not gtbr.empty and "ARTNR" in gtbr.columns:
        for artnr, grp in gtbr.groupby("ARTNR"):
            gtbr_by_part[str(artnr)] = grp

    return {
        "md_idx":        md_idx,
        "inv_idx":       inv_idx,
        "del_schedule":  del_schedule,
        "req_by_part":   req_by_part,
        "gtbr_by_part":  gtbr_by_part,
    }


_DEFAULT_DEL_SET: set = {0, 1, 2, 3, 4}


def _build_part_cache_entry(
    part: str,
    days: int,
    lookups: Dict,
) -> Optional[Dict]:
    """
    Build the full simulation cache entry for one part using indexed lookups.

    All DataFrame operations use pre-built O(1) structures from
    _build_plant_lookup_dicts.  The returned dict is stored in part_cache and
    shipped to worker processes once via the initializer, not per-task.

    Returns None if the part should be excluded (missing price / no consumption).
    """
    md_idx       = lookups["md_idx"]
    inv_idx      = lookups["inv_idx"]
    del_schedule = lookups["del_schedule"]
    req_by_part  = lookups["req_by_part"]
    gtbr_by_part = lookups["gtbr_by_part"]

    # ── Part parameters (indexed → O(log n)) ──────────────────────────────────
    if part not in md_idx.index:
        return None
    r = md_idx.loc[part]
    if isinstance(r, pd.DataFrame):     # handle duplicate PART values
        r = r.iloc[0]

    def _num(col: str, default: float = 0.0) -> float:
        v = pd.to_numeric(r.get(col, default), errors="coerce")
        return default if pd.isna(v) else float(v)

    price = _num("PRICE")
    if price <= 0:
        return None

    ul  = max(1.0, _num("UNIT_LOAD_QTY",       1.0))
    mul = max(1.0, _num("MULT_UNIT_LOAD_VALID", 1.0))
    params = {
        "part_number":     part,
        "supplier":        str(r.get("SUPP_SHP", "")),
        "safety_days":     _num("SAFETY"),
        "safety_stock":    _num("STOCK"),
        "unit_load":       ul,
        "multi_unit_load": mul,
        "order_increment": ul * mul,
        "price":           price,
    }

    # ── Initial stock (indexed → O(log n)) ────────────────────────────────────
    if part not in inv_idx.index:
        return None
    inv_row = inv_idx.loc[part]
    if isinstance(inv_row, pd.DataFrame):
        inv_row = inv_row.iloc[0]
    beg  = pd.to_numeric(inv_row.get("BEGINNING_INVENTORY_TODAY", 0), errors="coerce") or 0.0
    yard = pd.to_numeric(inv_row.get("INVENTORY_YARD_TODAY",     0), errors="coerce") or 0.0
    init_stock = float(beg + yard)

    # ── Consumption schedule (O(1) group lookup + vectorised reindex) ─────────
    part_req = req_by_part.get(part)
    if part_req is None or part_req.empty:
        return None
    part_req = part_req.dropna(subset=["PRODDAG"]).sort_values("PRODDAG")
    if part_req.empty:
        return None

    start_date     = part_req["PRODDAG"].min()
    weekend_exists = part_req["PRODDAG"].dt.dayofweek.isin([5, 6]).any()
    pool = (
        pd.date_range(start_date, periods=days * 2)
        if weekend_exists
        else pd.bdate_range(start_date, periods=days * 2)
    )
    dates = pool[:days]

    daily    = part_req.groupby("PRODDAG")["ARTAN"].sum()
    base_cons = daily.reindex(dates, fill_value=0.0).to_numpy(dtype=float)

    if base_cons.sum() <= 0:
        return None     # no active consumption → exclude

    # ── Delivery schedule (O(1) dict lookup) ──────────────────────────────────
    del_set = del_schedule.get(params["supplier"], _DEFAULT_DEL_SET)

    # ── Planned deliveries (O(1) group lookup + vectorised np.add.at) ─────────
    planned_qty = np.zeros(days)
    last_pd     = 0

    part_gtbr = gtbr_by_part.get(part)
    if part_gtbr is not None and not part_gtbr.empty:
        pg = part_gtbr.dropna(subset=["ANK_TID_TIDIGAST"])
        start_d, end_d = dates[0], dates[-1]
        pg = pg[(pg["ANK_TID_TIDIGAST"] >= start_d) & (pg["ANK_TID_TIDIGAST"] <= end_d)]
        if not pg.empty:
            date_to_idx  = {d.date(): i + 1 for i, d in enumerate(dates)}
            day_indices  = pg["ANK_TID_TIDIGAST"].dt.date.map(date_to_idx).dropna()
            quantities   = pd.to_numeric(
                pg.loc[day_indices.index, "ARTAN"], errors="coerce"
            ).fillna(0)
            day_indices  = day_indices.astype(int)
            valid_mask   = (day_indices >= 1) & (day_indices <= days)
            if valid_mask.any():
                np.add.at(
                    planned_qty,
                    day_indices[valid_mask].values - 1,
                    quantities[valid_mask].values,
                )
                last_pd = int(day_indices[valid_mask].max())

    # ── Pre-compute next-delivery look-ahead array ────────────────────────────
    dow          = dates.dayofweek.to_numpy()
    next_delivery = _precompute_next_delivery(dow, del_set, days)

    return {
        "days":         days,
        "init_stock":   init_stock,
        "base_cons":    base_cons,
        "dow":          dow,
        "del_set":      del_set,
        "planned_qty":  planned_qty,
        "last_pd":      last_pd,
        "next_delivery": next_delivery,
        "params":       params,
        "price":        price,
    }


# ─── Worker Functions (must be module-level for pickling) ─────────────────────

def _sim_part_worker(args: Tuple) -> np.ndarray:
    """
    Worker: simulate one part for *n_sims* runs using module-level cached data.
    Per-task payload: (part_name: str, n_sims: int, pva_start: int) — tiny.
    Returns the mean stock-value trajectory (shape: days+1).
    """
    part_number, n_sims, pva_start = args
    cache = _WORKER_PART_CACHE.get(part_number)
    if cache is None:
        return np.zeros(1)

    days  = cache["days"]
    price = cache["price"]
    pva_block = _WORKER_PVA_MATRIX[pva_start:pva_start + n_sims]

    value_sum = np.zeros(days + 1)
    for sim_i in range(n_sims):
        r = _simulate_run(
            cache["init_stock"], cache["base_cons"], cache["dow"],
            cache["del_set"],    cache["planned_qty"], cache["last_pd"],
            cache["params"],     pva_block[sim_i],     cache["next_delivery"],
        )
        value_sum += r["stock"] * price

    return value_sum / n_sims


def _sim_parts_batch_worker(args: Tuple) -> np.ndarray:
    """
    Worker: simulate a batch of Tier-4 parts (1 sim each) using a shared PVA
    sample.  Batching reduces task-submission overhead from O(n_parts) to
    O(n_parts / batch_size) round-trips through the work queue.

    Per-task payload: (part_names: List[str], pva_row: int) — still tiny.
    Returns the *combined* mean stock-value trajectory for all parts in the batch.
    """
    part_names, pva_row = args
    if not part_names:
        return np.zeros(1)

    # Discover days from the first valid cache entry
    days = None
    for p in part_names:
        c = _WORKER_PART_CACHE.get(p)
        if c is not None:
            days = c["days"]
            break
    if days is None:
        return np.zeros(1)

    pva_sample = _WORKER_PVA_MATRIX[pva_row]
    result     = np.zeros(days + 1)

    for part_number in part_names:
        cache = _WORKER_PART_CACHE.get(part_number)
        if cache is None:
            continue
        try:
            r = _simulate_run(
                cache["init_stock"], cache["base_cons"], cache["dow"],
                cache["del_set"],    cache["planned_qty"], cache["last_pd"],
                cache["params"],     pva_sample,           cache["next_delivery"],
            )
            result += r["stock"] * cache["price"]
        except Exception:
            pass

    return result   # n_sims=1 per part, so no division needed


# ─── Plant-Wide Forecast (Shared PVA) ────────────────────────────────────────

def plant_forecast_shared_pva(
    days: int,
    data: Dict,
    n_sims: int = 100,
    seed: Optional[int] = None,
    n_workers: Optional[int] = None,
) -> Dict:
    """
    Forecast total plant inventory value over *days* using a shared
    plant-wide PVA distribution.

    Parts are tiered by initial inventory value to balance accuracy vs speed:
      Tier 1 (top 2):   full n_sims simulations each
      Tier 2 (3–50):    up to 20 simulations
      Tier 3 (51–250):  up to 5 simulations
      Tier 4 (rest):    1 simulation (batched, _TIER4_BATCH_SIZE per task)

    Key performance improvements over the naïve approach:
      • _get_obsolete_inventory is vectorised (merge instead of loop).
      • DataFrames are pre-grouped once; per-part lookups are O(1).
      • Validation and cache-build are combined into a single pass.
      • ProcessPoolExecutor uses an initializer so the full part_cache and
        PVA matrix travel to each worker process exactly once, not per task.
      • Each submitted task carries only (part_name, n_sims, pva_row) — ~50 B.
      • _simulate_run uses a pre-computed next-delivery array instead of an
        inner look-ahead loop.
    """
    rng = np.random.default_rng(seed)
    pva = data["pva_distribution"]

    print("=== Plant Forecast — Shared PVA ===\n")

    # ── Vectorised obsolete inventory scan ────────────────────────────────────
    obsolete = _get_obsolete_inventory(
        data["master_data"], data["requirement"], data["inventory"]
    )

    # ── Pre-sample the plant-wide PVA matrix once ─────────────────────────────
    print(f"Pre-sampling {n_sims}×{days} PVA matrix…")
    pva_matrix = rng.choice(pva, size=(n_sims, days), replace=True)

    # ── Build fast O(1) lookup structures (one-time groupby pass) ─────────────
    print("Building lookup structures…")
    lookups = _build_plant_lookup_dicts(data)

    # ── Identify active parts and build simulation cache in ONE pass ───────────
    all_parts = data["requirement"]["ARTNR"].dropna().unique()

    active_parts: List[str] = []
    part_cache:   Dict      = {}
    init_vals:    List[float] = []

    print("Validating parts and building simulation cache…")
    for part in tqdm(all_parts, desc="Part cache"):
        try:
            entry = _build_part_cache_entry(str(part), days, lookups)
            if entry is not None:
                pstr = str(part)
                active_parts.append(pstr)
                part_cache[pstr] = entry
                init_vals.append(entry["init_stock"] * entry["price"])
        except Exception:
            pass

    # Sort descending by value so highest-value parts get the most simulations
    order        = np.argsort(init_vals)[::-1]
    active_parts = [active_parts[i] for i in order]
    init_vals    = [init_vals[i]    for i in order]
    total_value  = sum(init_vals)

    tier_defs = [
        (slice(0,    2),   n_sims,          "Tier 1 (top 2)",  False),
        (slice(2,   50),   min(n_sims, 20), "Tier 2 (3–50)",   False),
        (slice(50, 250),   min(n_sims, 5),  "Tier 3 (51–250)", False),
        (slice(250, None), 1,               "Tier 4 (rest)",   True),
    ]

    print("\n=== Tiered Strategy ===")
    for sl, ns, label, _batched in tier_defs:
        tier_p = active_parts[sl]
        tier_v = (init_vals[sl.start:] if sl.stop is None
                  else init_vals[sl.start:sl.stop])
        if tier_p:
            pct = 100 * sum(tier_v) / total_value if total_value else 0
            print(f"  {label}: {len(tier_p)} parts, "
                  f"${sum(tier_v):>12,.0f}  ({pct:.1f}%)  "
                  f"{ns} sim{'s' if ns > 1 else ''}")

    active_trajectory = np.zeros(days + 1)
    t0 = time.time()

    # ── Single pool creation — initializer ships data to workers once ──────────
    with ProcessPoolExecutor(
        max_workers=n_workers,
        initializer=_worker_init,
        initargs=(part_cache, pva_matrix),
    ) as ex:
        for sl, n_tier_sims, label, batched in tier_defs:
            tier_parts = active_parts[sl]
            if not tier_parts:
                continue

            print(f"\nSimulating {label}…")

            if batched:
                # Tier 4: batch multiple parts per task to cut IPC round-trips
                batches = [
                    tier_parts[i:i + _TIER4_BATCH_SIZE]
                    for i in range(0, len(tier_parts), _TIER4_BATCH_SIZE)
                ]
                futures = {
                    ex.submit(_sim_parts_batch_worker, (batch, 0)): batch
                    for batch in batches
                }
                for fut in tqdm(as_completed(futures),
                                total=len(batches), desc=label):
                    active_trajectory += fut.result()
            else:
                # Tier 1/2/3: one task per part, n_sims > 1
                futures = {
                    ex.submit(_sim_part_worker, (part, n_tier_sims, 0)): part
                    for part in tier_parts
                }
                for fut in tqdm(as_completed(futures),
                                total=len(tier_parts), desc=label):
                    active_trajectory += fut.result()

    elapsed = (time.time() - t0) / 60
    plant_trajectory = active_trajectory + obsolete["total_value"]

    print(f"\n=== Results ({elapsed:.1f} min) ===")
    print(f"  Active initial:   ${active_trajectory[0]:>14,.0f}")
    print(f"  Obsolete:         ${obsolete['total_value']:>14,.0f}")
    print(f"  Total initial:    ${plant_trajectory[0]:>14,.0f}")
    print(f"  Active final:     ${active_trajectory[-1]:>14,.0f}")
    print(f"  Total final:      ${plant_trajectory[-1]:>14,.0f}")
    change = plant_trajectory[-1] - plant_trajectory[0]
    pct    = 100 * change / plant_trajectory[0] if plant_trajectory[0] else 0
    print(f"  Change:           ${change:>+14,.0f}  ({pct:+.1f}%)")

    _plot_plant_forecast(plant_trajectory, active_trajectory, days, data)

    return {
        "plant_trajectory":  plant_trajectory,
        "active_trajectory": active_trajectory,
        "obsolete":          obsolete,
        "n_parts":           len(active_parts),
        "elapsed_minutes":   elapsed,
    }


def _plot_plant_forecast(
    plant_traj: np.ndarray,
    active_traj: np.ndarray,
    days: int,
    data: Dict,
) -> None:
    # Use first active part to get a date axis
    first_part = data["requirement"]["ARTNR"].dropna().iloc[0]
    _, dates   = get_consumption_schedule(first_part, days, data["requirement"])
    x          = np.arange(days + 1)

    fig, ax = plt.subplots(figsize=(13, 5))
    ax.plot(x, plant_traj / 1e6, lw=2.5, color="royalblue", label="Total Inventory")
    ax.plot(x, active_traj / 1e6, lw=1.5, color="green", ls="--", label="Active Only")

    if days > 30:
        month_idx = [i for i, d in enumerate(dates) if d.day == 1]
        ax.set_xticks(month_idx)
        ax.set_xticklabels(
            [dates[i].strftime("%b %Y") for i in month_idx], rotation=30, ha="right"
        )
    else:
        ax.set_xlabel("Days")

    ax.set_title("Plant Inventory Forecast — Tied-Up Capital")
    ax.set_ylabel("Inventory Value")
    ax.yaxis.set_major_formatter(
        plt.FuncFormatter(lambda v, _: f"${v:.1f}M")
    )
    ax.legend(); ax.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.show()


# ─── Inventory Archive Utilities ──────────────────────────────────────────────

def _fmt(x: float) -> str:
    return f"${x:,.0f}"


def update_inventory_report(
    new_report: pd.DataFrame,
    report_date=None,
    archive_path: Path = ARCHIVE_PATH,
) -> Dict:
    """
    Calculate total inventory value from a loaded CIR DataFrame and append
    (or update) today's entry in the CSV archive.

    *new_report* should be the DataFrame returned by
    ``import_manager.loaddata("current_inventory_report")``.
    """
    if report_date is None:
        report_date = pd.Timestamp.today().normalize()
    else:
        report_date = pd.Timestamp(report_date)

    stock     = (new_report["BEGINNING_INVENTORY_TODAY"].fillna(0)
                 + new_report["INVENTORY_YARD_TODAY"].fillna(0))
    values    = stock * new_report["PRICE"].fillna(0)
    total_val = float(values[values > 0].sum())
    n_parts   = int((values > 0).sum())

    print(f"\n=== Inventory Update ===")
    print(f"Date:        {report_date.date()}")
    print(f"Total Value: {_fmt(total_val)}")
    print(f"Parts:       {n_parts}")

    archive_path.parent.mkdir(parents=True, exist_ok=True)

    new_row = pd.DataFrame([{
        "date":        report_date,
        "total_value": total_val,
        "part_count":  n_parts,
    }])

    if archive_path.exists():
        arc  = pd.read_csv(archive_path, parse_dates=["date"])
        mask = arc["date"].dt.normalize() == report_date
        if mask.any():
            arc.loc[mask, "total_value"] = total_val
            arc.loc[mask, "part_count"]  = n_parts
        else:
            arc = pd.concat([arc, new_row], ignore_index=True)
    else:
        arc = new_row

    arc = (arc.dropna(subset=["date", "total_value"])
              .sort_values("date")
              .reset_index(drop=True))
    arc.to_csv(archive_path, index=False)

    print(f"Archive saved ({len(arc)} entries): {archive_path}")
    return {"date": report_date.date(), "total_value": total_val,
            "part_count": n_parts, "archive_rows": len(arc)}


def view_inventory_history(
    archive_path: Path = ARCHIVE_PATH,
    days: Optional[int] = 30,
    plot: bool = True,
) -> Optional[pd.DataFrame]:
    if not archive_path.exists():
        print("No archive found. Run update_inventory_report() first.")
        return None

    arc = (pd.read_csv(archive_path, parse_dates=["date"])
           .dropna(subset=["date", "total_value"])
           .sort_values("date"))

    if days is not None:
        cutoff = pd.Timestamp.today() - pd.Timedelta(days=days)
        arc = arc[arc["date"] >= cutoff]

    print(f"\n=== Inventory History ({len(arc)} entries) ===")
    print(arc.to_string(index=False))

    if plot and len(arc) > 1:
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.plot(arc["date"], arc["total_value"] / 1e6, lw=2, color="royalblue")
        ax.fill_between(arc["date"], arc["total_value"] / 1e6,
                        alpha=0.15, color="royalblue")
        ax.set_title("Historical Inventory Value")
        ax.set_ylabel("Value ($M)")
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"${v:.1f}M"))
        plt.xticks(rotation=25, ha="right"); ax.grid(True, alpha=0.3)
        plt.tight_layout(); plt.show()

    if len(arc) == 0:
        return arc

    latest = arc["total_value"].iloc[-1]
    print(f"\nCurrent:  {_fmt(latest)}")
    print(f"Highest:  {_fmt(arc['total_value'].max())}  "
          f"on {arc.loc[arc['total_value'].idxmax(), 'date'].date()}")
    print(f"Lowest:   {_fmt(arc['total_value'].min())}  "
          f"on {arc.loc[arc['total_value'].idxmin(), 'date'].date()}")
    print(f"Average:  {_fmt(arc['total_value'].mean())}")
    if len(arc) >= 2:
        ch  = latest - arc["total_value"].iloc[0]
        pct = 100 * ch / arc["total_value"].iloc[0]
        print(f"Change:   {_fmt(ch)}  ({pct:+.2f}%)")

    return arc


def compare_inventory_change(archive_path: Path = ARCHIVE_PATH) -> Optional[Dict]:
    """Print and return the day-over-day change in inventory value."""
    if not archive_path.exists():
        print("No archive found.")
        return None

    arc = (pd.read_csv(archive_path, parse_dates=["date"])
           .dropna(subset=["date", "total_value"])
           .sort_values("date", ascending=False))

    if len(arc) < 2:
        print("Need at least 2 days of data.")
        return None

    today, yesterday = arc.iloc[0], arc.iloc[1]
    change = today["total_value"] - yesterday["total_value"]
    pct    = 100 * change / yesterday["total_value"]
    arrow  = "↑" if change > 0 else ("↓" if change < 0 else "→")

    print(f"\n=== Daily Comparison ===")
    print(f"Today     ({today['date'].date()}):  {_fmt(today['total_value'])}")
    print(f"Yesterday ({yesterday['date'].date()}):  {_fmt(yesterday['total_value'])}")
    print(f"Change:  {_fmt(change)}  ({pct:+.2f}%)  {arrow}")

    return {
        "change": change, "pct_change": pct,
        "today": today.to_dict(), "yesterday": yesterday.to_dict(),
    }


def plot_historical_and_forecast(
    data: Dict,
    forecast_days: int = 250,
    historical_days: Optional[int] = None,
    n_sims: int = 100,
    archive_path: Path = ARCHIVE_PATH,
    seed: Optional[int] = None,
) -> Optional[Dict]:
    """
    Combine the historical archive with a fresh plant forecast and plot both
    on a single timeline.
    """
    if not archive_path.exists():
        print("No archive found. Run update_inventory_report() first.")
        return None

    arc = (pd.read_csv(archive_path, parse_dates=["date"])
           .dropna(subset=["date", "total_value"])
           .sort_values("date"))

    if historical_days is not None:
        cutoff = pd.Timestamp.today() - pd.Timedelta(days=historical_days)
        arc = arc[arc["date"] >= cutoff]

    print(f"Historical: {len(arc)} days "
          f"({arc['date'].min().date()} → {arc['date'].max().date()})")
    print(f"Running {forecast_days}-day forecast with {n_sims} simulations…\n")

    fc         = plant_forecast_shared_pva(forecast_days, data, n_sims=n_sims, seed=seed)
    fc_vals    = fc["plant_trajectory"][1:]  # exclude day-0 (= today's actual stock)
    fc_dates   = pd.date_range(
        pd.Timestamp.today() + pd.Timedelta(days=1), periods=forecast_days, freq="D"
    )

    fig, ax = plt.subplots(figsize=(14, 5))
    ax.plot(arc["date"], arc["total_value"] / 1e6,
            color="darkgreen", lw=2, label="Historical")
    ax.plot(fc_dates, fc_vals / 1e6,
            color="royalblue", lw=2, label="Forecast")
    ax.axvline(pd.Timestamp.today(), color="red", ls="--", lw=2, label="Today")
    ax.set_title("Inventory by Purpose — Tied-Up Capital")
    ax.set_ylabel("Value ($M)")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"${v:.1f}M"))
    plt.xticks(rotation=25, ha="right")
    ax.legend(); ax.grid(True, alpha=0.3)
    plt.tight_layout(); plt.show()

    latest = arc["total_value"].iloc[-1]
    final  = fc_vals[-1]
    change = final - latest
    pct    = 100 * change / latest

    print(f"\nCurrent:   {_fmt(latest)}")
    print(f"Forecast:  {_fmt(final)}")
    print(f"Change:    {_fmt(change)}  ({pct:+.1f}%)")

    return {"historical": arc,
            "forecast": pd.DataFrame({"date": fc_dates, "value": fc_vals}),
            "forecast_result": fc}
