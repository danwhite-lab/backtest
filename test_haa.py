import numpy as np
import pandas as pd

import test_market_filter_frequency as base


NAMESPACE = base.NAMESPACE


def fixture():
    index = pd.bdate_range("2023-01-02", "2025-03-31")
    assets = ["TIP", "SPY", "BIL", "IEF", "IWM", "VEA", "VWO", "VNQ", "DBC", "TLT"]
    values = {}
    for position, asset in enumerate(assets):
        values[asset] = 100 * np.exp(np.arange(len(index)) * (0.00025 + position * 0.00001))
    return pd.DataFrame(values, index=index)


def test_13612u_is_unweighted_month_return_mean_and_uses_no_future_data():
    prices = fixture()
    dates = NAMESPACE["make_filter_evaluation_dates"](prices.index, "Monthly")
    date = next(dt for dt in dates if dt - pd.DateOffset(months=12) >= prices.index.min())
    momentum = NAMESPACE["haa_13612u"](prices, pd.DatetimeIndex([date])).loc[date, "SPY"]
    now = NAMESPACE["last_price_on_or_before"](prices[["SPY"]], date)["SPY"]
    expected = sum(now / NAMESPACE["last_price_on_or_before"](prices[["SPY"]], date - pd.DateOffset(months=m))["SPY"] - 1 for m in (1, 3, 6, 12)) / 4
    assert abs(momentum - expected) < 1e-12
    changed = prices.copy()
    changed.loc[changed.index > date, "SPY"] *= 100
    assert NAMESPACE["haa_13612u"](changed, pd.DatetimeIndex([date])).loc[date, "SPY"] == momentum


def test_canonical_defensive_and_simple_rules():
    prices = fixture()
    targets, audit, _ = NAMESPACE["generate_haa_targets"](prices, NAMESPACE["HAA_SIMPLE_STRATEGY_NAME"])
    assert np.allclose(targets.sum(axis=1), 1.0)
    for dt, row in audit.iterrows():
        defensive = "BIL" if row["BIL 13612U"] > row["IEF 13612U"] else "IEF"
        assert row["Best defensive asset"] == defensive
        if row["TIP 13612U momentum"] <= 0 or row["SPY 13612U"] <= 0:
            assert targets.loc[row["Date"], defensive] == 1.0


def test_balanced_uses_tip_regime_top_four_and_slot_weights():
    prices = fixture()
    targets, audit, _ = NAMESPACE["generate_haa_targets"](prices, NAMESPACE["HAA_BALANCED_STRATEGY_NAME"])
    assert np.allclose(targets.sum(axis=1), 1.0)
    for _, row in audit.iterrows():
        weights = row["Target weights"]
        if row["TIP 13612U momentum"] <= 0:
            assert weights[row["Best defensive asset"]] == 1.0
        else:
            assert len(row["Offensive ranking"].split(", ")) == 4
            assert all(abs(weight / 0.25 - round(weight / 0.25)) < 1e-10 for weight in weights.values())


if __name__ == "__main__":
    test_13612u_is_unweighted_month_return_mean_and_uses_no_future_data()
    test_canonical_defensive_and_simple_rules()
    test_balanced_uses_tip_regime_top_four_and_slot_weights()
    print("Passed: canonical HAA momentum, defensive, Simple, and Balanced rules.")
