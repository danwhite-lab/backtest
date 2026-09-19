import ast
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd


SOURCE = Path(__file__).with_name("simple_etf_backtester_colab.py").read_text()
START = SOURCE.index("DEFAULT_TAX_RATE =")
END = SOURCE.index("# ---------- 16. DISPLAY")
TREE = ast.parse(SOURCE[START:END])
TREE.body = [
    node
    for node in TREE.body
    if not (
        isinstance(node, ast.Assign)
        and any(isinstance(target, ast.Name) and target.id == "import_output" for target in node.targets)
    )
]
NAMESPACE = globals().copy()
exec(compile(TREE, "backtester_core", "exec"), NAMESPACE)


def assert_close(actual, expected, tolerance=1e-10):
    assert abs(actual - expected) <= tolerance, (actual, expected)


def test_no_contribution_metrics_are_unchanged():
    dates = pd.bdate_range("2022-01-03", periods=500)
    equity = pd.Series(100_000 * np.exp(np.arange(len(dates)) * 0.0004), index=dates)
    holdings = pd.Series("QQQ", index=dates)
    metrics = NAMESPACE["calculate_metrics"](equity, pd.DataFrame(), holdings, initial_value=100_000)

    years = (dates[-1] - dates[0]).days / 365.25
    expected_cagr = (equity.iloc[-1] / equity.iloc[0]) ** (1 / years) - 1
    expected_drawdown = (equity / equity.cummax() - 1).min()
    assert_close(metrics["CAGR"], expected_cagr)
    assert_close(metrics["Max Drawdown"], expected_drawdown)
    assert_close(metrics["Periodic Contributions"], 0.0)


def test_deposits_are_not_returns():
    dates = pd.to_datetime(["2021-01-04", "2022-01-04", "2023-01-04"])
    equity = pd.Series([100.0, 210.0, 331.0], index=dates)
    # Each year the existing balance earns 10%, then a $100 end-of-day deposit arrives.
    flows = pd.Series([0.0, 100.0, 100.0], index=dates)
    holdings = pd.Series("ASSET", index=dates)
    metrics = NAMESPACE["calculate_metrics"](
        equity,
        pd.DataFrame(),
        holdings,
        external_flows=flows,
        initial_value=100.0,
    )

    years = (dates[-1] - dates[0]).days / 365.25
    expected_twr = 1.21 ** (1 / years) - 1
    assert_close(metrics["CAGR"], expected_twr)
    assert_close(metrics["Periodic Contributions"], 200.0)
    assert_close(metrics["Total Invested"], 300.0)
    assert_close(metrics["Net Profit"], 31.0)
    assert metrics["Money-Weighted Return"] > 0


def test_constant_price_benchmark_uses_same_cash_flows():
    dates = pd.bdate_range("2023-01-02", periods=6)
    prices = pd.DataFrame({"BENCH": 10.0}, index=dates)
    flows = pd.Series(0.0, index=dates)
    flows.iloc[2] = 100.0
    flows.iloc[4] = 100.0
    benchmark = NAMESPACE["benchmark_buyhold"](
        prices, "BENCH", initial_value=100.0, external_flows=flows
    )
    metrics = NAMESPACE["calculate_metrics"](
        benchmark,
        pd.DataFrame(),
        pd.Series("BENCH", index=dates),
        external_flows=flows,
        initial_value=100.0,
    )

    assert_close(benchmark.iloc[-1], 300.0)
    assert_close(metrics["CAGR"], 0.0)
    assert_close(metrics["Max Drawdown"], 0.0)
    assert_close(metrics["Money-Weighted Return"], 0.0, tolerance=1e-8)


def test_agitq_and_benchmark_share_contribution_schedule():
    dates = pd.bdate_range("2019-01-02", "2024-12-31")
    step = np.arange(len(dates))
    prices = pd.DataFrame(
        {
            "QQQ": 100 * np.exp(step * 0.00025 + np.sin(step / 80) * 0.08),
            "TQQQ": 40 * np.exp(step * 0.00055 + np.sin(step / 80) * 0.22),
            "SPYM": 50 * np.exp(step * 0.00015),
            "BENCH": 80 * np.exp(step * 0.0002),
        },
        index=dates,
    )
    NAMESPACE["aligned_prices"] = lambda names, base_currency: prices[names]
    result = NAMESPACE["run_agitq_test"](
        variant="QQQ 3/161",
        signal_asset="QQQ",
        traded_asset="TQQQ",
        defensive_asset="CASH",
        parking_asset="SPYM",
        short_sma_days=3,
        long_sma_days=161,
        contribution_amount=100.0,
        contribution_frequency="Monthly",
        benchmark_asset="BENCH",
        initial_value=100.0,
        tax_rate=0.0,
        tx_cost=0.0025,
        start_date="2021-01-01",
        end_date="2024-12-31",
    )

    flow_count = int((result["external_flows"] > 0).sum())
    assert flow_count > 40
    assert_close(result["total_contributions"], result["external_flows"].sum())
    assert_close(
        result["benchmark_metrics"]["Periodic Contributions"],
        result["metrics"]["Periodic Contributions"],
    )
    assert result["benchmark_metrics"]["Trades"] == 1 + flow_count
    assert result["metrics"]["CAGR"] < 1.0
    assert result["metrics"]["Total Invested"] == 100.0 + result["total_contributions"]


if __name__ == "__main__":
    test_no_contribution_metrics_are_unchanged()
    test_deposits_are_not_returns()
    test_constant_price_benchmark_uses_same_cash_flows()
    test_agitq_and_benchmark_share_contribution_schedule()
    print(
        "Passed: no-flow regression, cash-flow-adjusted TWR, XIRR, totals, "
        "AGITQ integration, and benchmark contributions."
    )
