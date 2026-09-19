import numpy as np
import pandas as pd
import test_market_filter_frequency as base

NAMESPACE = base.NAMESPACE

def test_gtt_uses_previous_month_yoy_and_ten_month_sma_rules():
    dates = pd.bdate_range("2022-01-03", "2024-12-31")
    prices = pd.Series(100 + np.arange(len(dates)) * .1, index=dates)
    months = pd.date_range("2021-01-31", "2024-11-30", freq="ME")
    rrsfs = pd.Series(np.linspace(100, 150, len(months)), index=months)
    indpro = pd.Series(np.linspace(100, 150, len(months)), index=months)
    signals, audit = NAMESPACE["generate_gtt_signals"](prices, rrsfs, indpro)
    row = audit.dropna(subset=["RRSFS YoY"]).iloc[-1]
    value, yoy = NAMESPACE["gtt_previous_month_yoy"](rrsfs, row["Date"])
    assert row["RRSFS previous-month value"] == value and row["RRSFS YoY"] == yoy
    assert row["Growth regime"] == "Growth Healthy" and row["Allocation"] == "SPY"
    assert audit["10-month SMA"].notna().any()

if __name__ == "__main__":
    test_gtt_uses_previous_month_yoy_and_ten_month_sma_rules()
    print("Passed: GTT prior-month YoY and monthly SMA rules.")
