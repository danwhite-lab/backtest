import ast
import json
import os
import re
import sqlite3
import tempfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd


SOURCE = Path(__file__).with_name("simple_etf_backtester_colab.py").read_text()
START = SOURCE.index("DEFAULT_TAX_RATE =")
END = SOURCE.index("# ---------- 18. SIMPLE MAIN UI")
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
exec(compile(TREE, "backtester_history", "exec"), NAMESPACE)


def sample_result(strategy="Dual Momentum", primary="QQQ", cagr=0.12, sharpe=0.8):
    dates = pd.bdate_range("2021-01-04", periods=3)
    metrics = {
        "Start": dates[0],
        "End": dates[-1],
        "CAGR": cagr,
        "Max Drawdown": -0.2,
        "Volatility": 0.18,
        "Sharpe": sharpe,
        "Final Value": 125_000.0,
        "Trades": 9,
        "Worst Year": -0.1,
    }
    return {
        "metrics": metrics,
        "settings": {
            "Strategy": strategy,
            "Primary asset": primary,
            "Secondary asset": "XLE",
            "Benchmark asset": "SPY",
            "SMA days": 200,
            "ROC months": 12,
        },
    }


def test_storage_filters_sort_delete_and_batch_save():
    with tempfile.TemporaryDirectory() as directory:
        NAMESPACE["RESULTS_DB_PATH"] = os.path.join(directory, "history.sqlite")
        first = NAMESPACE["save_backtest_result"](sample_result())
        second = NAMESPACE["save_backtest_result"](
            sample_result(strategy="SMA Trend", primary="SPY", cagr=0.08, sharpe=0.4)
        )
        assert first != second

        filtered = NAMESPACE["load_saved_backtests"](
            asset_text="QQQ",
            parameter_text="200",
            numeric_filters={"cagr": "> 10%", "sharpe": "> 0.6"},
            sort_by="sharpe",
            ascending=False,
        )
        assert filtered["id"].tolist() == [first]
        assert filtered.iloc[0]["strategy"] == "Dual Momentum"
        assert json.loads(filtered.iloc[0]["settings_json"])["SMA days"] == 200

        assert NAMESPACE["_parse_numeric_filter"]("> -30%", ratio=True) == (">", -0.3)
        assert NAMESPACE["delete_saved_backtests"]([second]) == 1
        assert len(NAMESPACE["load_saved_backtests"]()) == 1

        def fake_run_test(**kwargs):
            return sample_result(
                strategy=kwargs["strategy"], primary=kwargs["primary"], cagr=0.15, sharpe=0.9
            )

        NAMESPACE["run_test"] = fake_run_test
        batch = NAMESPACE["batch_momentum_test"](
            primary="QQQ", secondary_options=["XLE"], lookbacks=(10, 12)
        )
        assert len(batch) == 2
        saved = NAMESPACE["load_saved_backtests"](strategy="Dual Momentum")
        assert len(saved) == 3
        assert set(saved["run_type"]) == {"Individual", "Batch"}


if __name__ == "__main__":
    test_storage_filters_sort_delete_and_batch_save()
    print("Passed: SQLite saves, combined filters, sorting, deletion, and batch persistence.")
