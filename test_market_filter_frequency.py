import ast
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
from pandas.testing import assert_series_equal


SOURCE = Path(__file__).with_name("simple_etf_backtester_colab.py").read_text()
START = SOURCE.index("DEFAULT_TAX_RATE =")
END = SOURCE.index("# ---------- 16. DISPLAY")
TREE = ast.parse(SOURCE[START:END])
TREE.body = [
    node for node in TREE.body
    if not (
        isinstance(node, ast.Assign)
        and any(isinstance(target, ast.Name) and target.id == "import_output" for target in node.targets)
    )
]
NAMESPACE = globals().copy()
exec(compile(TREE, "backtester_core", "exec"), NAMESPACE)


def price_fixture():
    """Daily prices with several 3-day-SMA crossings inside weeks and months."""
    dates = pd.bdate_range("2024-01-02", "2024-03-29")
    values = [100 + ((position * 7) % 19) - 9 for position in range(len(dates))]
    return pd.DataFrame({"IWB": values, "RISK": 100.0}, index=dates)


def expanded_state(prices, frequency):
    sparse = NAMESPACE["calculate_market_filter_state"](
        prices, "IWB", "Price > SMA", 3, frequency
    )
    return sparse.reindex(prices.index).astype("boolean").ffill().fillna(False).astype(bool), sparse.index


def assert_state_changes_only_on_evaluations(prices, frequency):
    persisted, evaluation_dates = expanded_state(prices, frequency)
    # Before the first weekly/monthly evaluation the implementation intentionally
    # has no state and forces CASH. Audit only changes after an evaluable state exists.
    persisted = persisted.loc[evaluation_dates.min():]
    changed_dates = persisted.index[persisted.ne(persisted.shift()).fillna(True)]
    assert set(changed_dates).issubset(set(evaluation_dates))


def test_weekly_state_changes_only_on_final_trading_day_of_week():
    prices = price_fixture()
    assert_state_changes_only_on_evaluations(prices, "Weekly")
    evaluation_dates = NAMESPACE["make_filter_evaluation_dates"](prices.index, "Weekly")
    for dt in evaluation_dates:
        week_start = dt.normalize() - pd.Timedelta(days=dt.weekday())
        week_dates = prices.index[(prices.index >= week_start) & (prices.index < week_start + pd.Timedelta(days=7))]
        assert dt == week_dates.max()


def test_monthly_state_changes_only_on_final_trading_day_of_month():
    prices = price_fixture()
    assert_state_changes_only_on_evaluations(prices, "Monthly")
    evaluation_dates = NAMESPACE["make_filter_evaluation_dates"](prices.index, "Monthly")
    assert all(dt == prices.loc[str(dt.year) + "-" + f"{dt.month:02d}"].index.max() for dt in evaluation_dates)


def test_daily_filter_matches_pre_frequency_implementation_exactly():
    prices = price_fixture()
    signals = pd.Series("RISK", index=prices.index, dtype="object")
    expected = signals.copy()
    for dt in expected.index:
        sma = NAMESPACE["sma_value"](prices["IWB"], dt, 3)
        price = NAMESPACE["last_price_on_or_before"](prices[["IWB"]], dt)["IWB"]
        if pd.isna(sma) or not (price > sma):
            expected.loc[dt] = "CASH"

    actual = NAMESPACE["apply_market_filter"](
        signals, prices, "IWB", "Price > SMA", 3, "Daily"
    )
    assert_series_equal(actual, expected)


def test_weekly_filter_uses_persisted_not_intervening_raw_state():
    prices = price_fixture()
    persisted, evaluation_dates = expanded_state(prices, "Weekly")
    raw_sma = prices["IWB"].rolling(3, min_periods=3).mean()
    raw = (prices["IWB"] > raw_sma).where(raw_sma.notna())
    intervening_difference = (
        (~prices.index.isin(evaluation_dates)) & raw.notna() & raw.ne(persisted)
    )
    assert intervening_difference.any(), "Fixture must exercise a between-week raw crossover."

    signals = pd.Series("RISK", index=prices.index, dtype="object")
    filtered = NAMESPACE["apply_market_filter"](
        signals, prices, "IWB", "Price > SMA", 3, "Weekly"
    )
    assert_series_equal(filtered.eq("RISK"), persisted, check_names=False)


def test_diagnostic_table_exposes_raw_and_persisted_states_and_trades():
    prices = price_fixture()
    signals = pd.Series("RISK", index=prices.index, dtype="object")
    filtered = NAMESPACE["apply_market_filter"](
        signals, prices, "IWB", "Price > SMA", 3, "Monthly"
    )
    _, trades, holdings = NAMESPACE["backtest"](
        prices, filtered, initial_value=100.0, tax_rate=0.0, tx_cost=0.0, execution_mode="Same close"
    )
    diagnostics = NAMESPACE["build_market_filter_diagnostics"](
        prices, "IWB", "Price > SMA", 3, "Monthly", holdings, trades
    )
    assert list(diagnostics.columns) == [
        "Date", "IWB close", "IWB SMA (3)", "Raw price > SMA",
        "Filter evaluation date", "Persisted filter state", "Current holding", "Trade generated",
    ]
    assert diagnostics["Raw price > SMA"].notna().any()
    assert diagnostics["Trade generated"].str.len().gt(0).any()


if __name__ == "__main__":
    test_weekly_state_changes_only_on_final_trading_day_of_week()
    test_monthly_state_changes_only_on_final_trading_day_of_month()
    test_daily_filter_matches_pre_frequency_implementation_exactly()
    test_weekly_filter_uses_persisted_not_intervening_raw_state()
    test_diagnostic_table_exposes_raw_and_persisted_states_and_trades()
    print("Passed: market-filter frequency and diagnostics regression checks.")
