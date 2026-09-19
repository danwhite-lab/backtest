import ast
from itertools import product
from pathlib import Path
import re

import numpy as np
import pandas as pd

import test_market_filter_frequency as base


SOURCE = Path(__file__).with_name("simple_etf_backtester_colab.py").read_text()
START = SOURCE.index("# ---------- 18. BATCH TEST ----------")
END = SOURCE.index("# ---------- 18. SIMPLE MAIN UI ----------")
TREE = ast.parse(SOURCE[START:END])
NAMESPACE = base.NAMESPACE.copy()
NAMESPACE["product"] = product
def parse_numeric_filter(expression, ratio=False):
    matched = re.fullmatch(r"(<=|>=|=|<|>)\s*(-?(?:\d+(?:\.\d*)?|\.\d+))\s*(%)?", expression.strip())
    if not matched:
        raise ValueError("Invalid numeric filter")
    operator, raw_value, percent = matched.groups()
    value = float(raw_value)
    return operator, value / 100 if ratio and (percent or abs(value) > 1) else value
NAMESPACE["_parse_numeric_filter"] = parse_numeric_filter
exec(compile(TREE, "batch_core", "exec"), NAMESPACE)


def fake_run_test(**kwargs):
    score = 0.10 if kwargs["strategy"] == "Buy & Hold" else 0.20
    return {
        "metrics": {
            "Start": pd.Timestamp("2020-01-01"),
            "End": pd.Timestamp("2021-01-01"),
            "CAGR": score,
            "Max Drawdown": -score,
            "Volatility": score / 2,
            "Sharpe": score * 10,
            "Final Value": 100_000 * (1 + score),
            "Trades": int(score * 100),
            "Worst Year": -score / 2,
        }
    }


def batch_kwargs():
    return {
        "strategies": ["Buy & Hold", "SMA Trend"],
        "primary_assets": ["IWB"],
        "secondary_assets": ["None"],
        "base_currencies": ["USD"],
        "frequencies": ["Monthly"],
        "lookbacks": [12],
        "sma_days_values": [200],
        "envelope_pcts": [0.05],
        "tax_rates": [0.25],
        "transaction_costs": [0.001],
        "cash_rates": [0.0],
        "initial_values": [100_000],
        "execution_modes": ["Same close"],
        "benchmark_assets": ["None"],
        "filter_assets": ["None"],
        "filter_frequencies": ["Daily"],
    }


def test_batch_count_matches_selected_cartesian_product():
    kwargs = batch_kwargs()
    assert NAMESPACE["batch_combination_count"](**kwargs) == 2


def test_batch_reuses_run_test_and_returns_required_result_columns():
    NAMESPACE["run_test"] = fake_run_test
    results = NAMESPACE["run_batch_test"](**batch_kwargs())
    assert len(results) == 2
    assert list(results.columns) == [
        "Strategy", "Assets", "Parameters", "Test period", "CAGR", "Max drawdown",
        "Volatility", "Sharpe", "Final value", "Trades", "Worst year", "Error",
    ]
    assert set(results["Strategy"]) == {"Buy & Hold", "SMA Trend"}
    assert results["Error"].eq("").all()
    assert results["Parameters"].str.contains("SMA days=200").all()


def test_batch_results_sort_filter_and_highlights():
    NAMESPACE["run_test"] = fake_run_test
    results = NAMESPACE["run_batch_test"](**batch_kwargs())
    filtered = NAMESPACE["filter_and_sort_batch_results"](
        results, sort_by="CAGR", metric="CAGR", expression="> 15%"
    )
    assert filtered["Strategy"].tolist() == ["SMA Trend"]
    highlights = NAMESPACE["batch_result_highlights"](results)
    assert highlights.loc[highlights["Highlight"] == "Highest CAGR", "Strategy"].iloc[0] == "SMA Trend"
    assert highlights.loc[highlights["Highlight"] == "Lowest max drawdown", "Strategy"].iloc[0] == "SMA Trend"
    assert highlights.loc[highlights["Highlight"] == "Highest Sharpe", "Strategy"].iloc[0] == "SMA Trend"


if __name__ == "__main__":
    test_batch_count_matches_selected_cartesian_product()
    test_batch_reuses_run_test_and_returns_required_result_columns()
    test_batch_results_sort_filter_and_highlights()
    print("Passed: Batch Test combinations, engine reuse, table filtering, and highlights.")
