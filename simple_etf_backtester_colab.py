
# ============================================================
# SIMPLE ETF BACKTESTER — COLAB VERSION 1
# ============================================================
# Designed for approximate strategy comparison, not tax/accounting precision.
# Core features:
# - Multiple CSV and Excel uploads
# - Online Yahoo Finance ticker downloads
# - Manual currency selection and automatic date-format detection per dataset
# - Automatic price-column detection with manual override
# - USD or ILS base currency
# - Optional USDILS FX dataset
# - Buy & Hold
# - Absolute Momentum
# - Dual Momentum
# - Relative Momentum
# - Optional external market filter (Price > SMA)
# - SMA Trend
# - SMA Hysteresis Envelope
# - AGITQ / TQQQ Playbook variants with confirmation and overheat routing
# - Daily / Weekly / Monthly signals
# - Same-close or next-available-close execution
# - Approximate 25% realized capital-gains tax
# - Transaction costs
# - User-set annual return while the strategy is in cash
# - Benchmark comparison
# - AGITQ audit export with daily state, events, trades, settings, and metrics
# - Cash-flow-adjusted return, XIRR, max drawdown, volatility, Sharpe, trades, allocation
#
# IMPORTANT:
# This is a ballpark backtester. It prioritizes consistency and transparency.
# Dividends are included only when the chosen price column already includes them
# (for example, Adjusted Close).
# ============================================================

# ---------- 1. INSTALL / IMPORT ----------
import sys, subprocess, pkgutil, io, math, warnings, json, zipfile, tempfile, os, sqlite3, re
from dataclasses import dataclass
from datetime import date, datetime
from itertools import product
from typing import Dict, Optional, List, Tuple

required = ["pandas", "numpy", "matplotlib", "ipywidgets", "openpyxl", "xlrd", "yfinance"]
for pkg in required:
    if pkgutil.find_loader(pkg) is None:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", pkg])

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import ipywidgets as widgets
import yfinance as yf
from IPython.display import display, clear_output, FileLink

warnings.filterwarnings("ignore")

try:
    from google.colab import files
    IN_COLAB = True
except Exception:
    IN_COLAB = False

# ---------- 2. GLOBAL SETTINGS ----------
DEFAULT_TAX_RATE = 0.25
DEFAULT_TRANSACTION_COST = 0.001
DEFAULT_CASH_RATE = 0.0
DEFAULT_BASE_CURRENCY = "USD"

AGITQ_STRATEGY_NAME = "AGITQ / TQQQ Playbook"
HAA_SIMPLE_STRATEGY_NAME = "HAA-Simple (HAA-1 / U1-T1)"
HAA_BALANCED_STRATEGY_NAME = "HAA-Balanced (G8/T4)"
HAA_CANARY = "TIP"
HAA_DEFENSIVE = ["BIL", "IEF"]
HAA_BALANCED_OFFENSIVE = ["SPY", "IWM", "VEA", "VWO", "VNQ", "DBC", "IEF", "TLT"]
AGITQ_PRESETS = {
    "Original 200": {
        "signal_asset": "TQQQ", "traded_asset": "TQQQ",
        "short_sma": None, "long_sma": 200,
        "entry_sma": None, "exit_sma": None,
    },
    "QQQ 3/161": {
        "signal_asset": "QQQ", "traded_asset": "TQQQ",
        "short_sma": 3, "long_sma": 161,
        "entry_sma": None, "exit_sma": None,
    },
    "QLD ISA": {
        "signal_asset": "QQQ", "traded_asset": "QLD",
        "short_sma": 3, "long_sma": 161,
        "entry_sma": None, "exit_sma": None,
    },
    "TQQQ 5/218": {
        "signal_asset": "TQQQ", "traded_asset": "TQQQ",
        "short_sma": 5, "long_sma": 218,
        "entry_sma": None, "exit_sma": None,
    },
    "Hybrid 3/185/161": {
        "signal_asset": "QQQ", "traded_asset": "TQQQ",
        "short_sma": 3, "long_sma": None,
        "entry_sma": 185, "exit_sma": 161,
    },
}

DATE_FORMATS = {
    "Auto detect": None,
    "DD/MM/YYYY": "%d/%m/%Y",
    "MM/DD/YYYY": "%m/%d/%Y",
    "YYYY-MM-DD": "%Y-%m-%d",
    "YYYY/MM/DD": "%Y/%m/%d",
    "DD-MM-YYYY": "%d-%m-%Y",
    "MM-DD-YYYY": "%m-%d-%Y",
}

PRICE_PRIORITY = [
    "adj close", "adjusted close", "adjusted_close", "adj_close",
    "close", "price", "last", "closing index price", "closing_price"
]

IGNORE_HINTS = ["open", "high", "low", "volume", "vol", "change", "%", "pct"]

# ---------- 3. DATA CONTAINER ----------
@dataclass
class AssetData:
    name: str
    raw: pd.DataFrame
    currency: str = "USD"
    date_format_label: str = "Auto detect"
    detected_date_format: Optional[str] = None
    date_col: Optional[str] = None
    price_col: Optional[str] = None
    parsed: Optional[pd.Series] = None

ASSETS: Dict[str, AssetData] = {}
FX_SERIES: Optional[pd.Series] = None   # interpreted as ILS per 1 USD
FX_NAME: Optional[str] = None

# Stored beside the notebook/script by default. Set ETF_BACKTESTER_RESULTS_DB to
# an absolute path (for example, in Google Drive) when a different location is wanted.
RESULTS_DB_PATH = os.environ.get(
    "ETF_BACKTESTER_RESULTS_DB",
    os.path.abspath("simple_etf_backtester_results.sqlite"),
)

# ---------- 4. FILE IMPORT HELPERS ----------
def _read_csv_bytes(content: bytes) -> pd.DataFrame:
    for enc in ["utf-8", "utf-8-sig", "latin1"]:
        try:
            text = content.decode(enc)
            header_row = 0
            for i, line in enumerate(text.splitlines()[:25]):
                first_field = (
                    line.strip().lstrip("\ufeff").split(",", 1)[0]
                    .strip().strip('"').lower()
                )
                if first_field in {"date", "time_period", "datetime", "timestamp"}:
                    header_row = i
                    break
            return pd.read_csv(io.StringIO(text), skiprows=header_row)
        except Exception:
            pass
    raise ValueError("Could not read CSV.")

def _read_upload_bytes(filename: str, content: bytes) -> pd.DataFrame:
    """Read CSV or the first worksheet of an Excel workbook."""
    extension = filename.rsplit(".", 1)[-1].lower()
    if extension in ("xlsx", "xls"):
        engine = "openpyxl" if extension == "xlsx" else "xlrd"
        return pd.read_excel(io.BytesIO(content), engine=engine)
    if extension == "csv":
        return _read_csv_bytes(content)
    raise ValueError(f"Unsupported file type: {filename}. Upload CSV, .xlsx, or .xls files.")

def _normalize_col(c):
    return str(c).strip().lower().replace("\n", " ")

def detect_date_column(df: pd.DataFrame) -> str:
    cols = list(df.columns)
    for c in cols:
        n = _normalize_col(c)
        if "date" in n or "time" in n:
            return c
    return cols[0]

def _is_date_object(value) -> bool:
    return isinstance(value, (date, datetime, pd.Timestamp, np.datetime64))

def detect_date_format(series: pd.Series) -> str:
    """Infer a safe date format without guessing ambiguous day/month dates."""
    values = series.dropna()
    if values.empty:
        raise ValueError("Date column is empty.")

    if pd.api.types.is_datetime64_any_dtype(series):
        return "Excel date"

    object_dates = values.map(_is_date_object)
    if object_dates.all():
        return "Excel date"
    if object_dates.any():
        raise ValueError(
            "Date column mixes text with Excel dates. Use the original CSV; "
            "Excel may have converted ambiguous dates incorrectly."
        )

    text = values.astype(str).str.strip().str.split().str[0]
    if text.str.match(r"^\d{4}-\d{1,2}-\d{1,2}$").all():
        return "YYYY-MM-DD"
    if text.str.match(r"^\d{4}/\d{1,2}/\d{1,2}$").all():
        return "YYYY/MM/DD"

    parts = text.str.extract(r"^(\d{1,2})([/\-])(\d{1,2})\2(\d{4})$")
    if parts.isna().any(axis=None):
        raise ValueError("Could not detect the date format; choose it manually.")

    separators = parts[1].unique()
    if len(separators) != 1:
        raise ValueError("Date column uses mixed separators; choose a manual format.")

    first = pd.to_numeric(parts[0])
    second = pd.to_numeric(parts[2])
    first_is_day = bool((first > 12).any())
    second_is_day = bool((second > 12).any())
    if first_is_day and second_is_day:
        raise ValueError("Date column contains inconsistent day/month ordering.")
    if not first_is_day and not second_is_day:
        raise ValueError(
            "Date format is ambiguous because all days and months are 12 or less; "
            "choose DD/MM/YYYY or MM/DD/YYYY manually."
        )

    separator = separators[0]
    if first_is_day:
        return "DD/MM/YYYY" if separator == "/" else "DD-MM-YYYY"
    return "MM/DD/YYYY" if separator == "/" else "MM-DD-YYYY"

def parse_date_column(series: pd.Series, selected_format: str) -> Tuple[pd.Series, str]:
    values = series.dropna()
    if pd.api.types.is_datetime64_any_dtype(series) or (
        not values.empty and values.map(_is_date_object).all()
    ):
        return pd.to_datetime(series, errors="coerce"), "Excel date"

    object_dates = values.map(_is_date_object)
    if object_dates.any():
        raise ValueError(
            "Date column mixes text with Excel dates. Use the original CSV; "
            "Excel may have converted ambiguous dates incorrectly."
        )

    detected = detect_date_format(series) if selected_format == "Auto detect" else selected_format
    fmt = DATE_FORMATS.get(detected)
    if fmt is None:
        raise ValueError("Choose a supported date format.")

    text = series.astype("string").str.strip().str.split().str[0]
    dates = pd.to_datetime(text, format=fmt, errors="coerce")
    invalid = int(series.notna().sum() - dates.notna().sum())
    if invalid:
        raise ValueError(f"{invalid:,} dates did not match {detected}.")
    return dates, detected

def detect_price_column(df: pd.DataFrame) -> Optional[str]:
    normalized = {_normalize_col(c): c for c in df.columns}
    for target in PRICE_PRIORITY:
        if target in normalized:
            return normalized[target]

    # fuzzy priority
    for c in df.columns:
        n = _normalize_col(c)
        if any(h in n for h in IGNORE_HINTS):
            continue
        if "adjust" in n and "close" in n:
            return c
    for c in df.columns:
        n = _normalize_col(c)
        if any(h in n for h in IGNORE_HINTS):
            continue
        if n in ("close", "price", "last") or "closing" in n:
            return c

    # fallback numeric-like column
    for c in df.columns:
        s = pd.to_numeric(df[c].astype(str).str.replace(",", "").str.replace("%",""), errors="coerce")
        if s.notna().mean() > 0.85 and not any(h in _normalize_col(c) for h in IGNORE_HINTS):
            return c
    return None

def parse_asset(asset: AssetData) -> pd.Series:
    df = asset.raw.copy()
    dcol = asset.date_col or detect_date_column(df)
    pcol = asset.price_col or detect_price_column(df)

    if pcol is None:
        raise ValueError(f"No usable price column found for {asset.name}")

    dates, detected_format = parse_date_column(df[dcol], asset.date_format_label)

    px = (
        df[pcol].astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.strip()
    )
    px = pd.to_numeric(px, errors="coerce")

    out = pd.Series(px.values, index=dates, name=asset.name)
    out = out[~out.index.isna()]
    out = out[~out.index.duplicated(keep="last")]
    out = out.sort_index()
    out = out[out > 0]

    if len(out) < 20:
        raise ValueError(f"{asset.name}: too few valid rows after parsing.")

    asset.date_col = dcol
    asset.price_col = pcol
    asset.detected_date_format = detected_format
    asset.parsed = out
    return out

def download_yahoo_asset(
    ticker: str,
    start_date=None,
    end_date=None,
    currency: str = "USD",
) -> Tuple[str, pd.Series]:
    """Download a split- and dividend-adjusted daily price series from Yahoo Finance."""
    symbol = str(ticker).strip().upper()
    if not symbol:
        raise ValueError("Enter a Yahoo Finance ticker, for example QQQ or SPY.")
    if currency not in ("USD", "ILS"):
        raise ValueError("Currency must be USD or ILS.")

    start = pd.Timestamp(start_date).normalize() if start_date is not None else pd.Timestamp("1900-01-01")
    end = pd.Timestamp(end_date).normalize() if end_date is not None else None
    if end is not None and start > end:
        raise ValueError("Online data start date must be on or before end date.")

    # Yahoo treats end as exclusive, so add one day to include the selected date.
    end_exclusive = end + pd.Timedelta(days=1) if end is not None else None
    history = yf.download(
        symbol,
        start=start.date().isoformat(),
        end=end_exclusive.date().isoformat() if end_exclusive is not None else None,
        auto_adjust=False,
        actions=False,
        progress=False,
        threads=False,
    )
    if history is None or history.empty:
        raise ValueError(f"Yahoo Finance returned no price history for {symbol}.")

    if isinstance(history.columns, pd.MultiIndex):
        history.columns = history.columns.get_level_values(0)
    price_column = "Adj Close" if "Adj Close" in history.columns else "Close"
    prices = pd.to_numeric(history[price_column], errors="coerce")
    if isinstance(prices, pd.DataFrame):
        prices = prices.iloc[:, 0]
    dates = pd.DatetimeIndex(history.index)
    if dates.tz is not None:
        dates = dates.tz_localize(None)

    raw = pd.DataFrame({"Date": dates, "Price": prices.to_numpy()})
    asset = AssetData(
        name=symbol,
        raw=raw,
        currency=currency,
        date_col="Date",
        price_col="Price",
    )
    ASSETS[symbol] = asset
    parsed = parse_asset(asset)
    return symbol, parsed

# ---------- 5. UPLOAD ----------
def upload_csvs():
    global ASSETS
    if not IN_COLAB:
        print("This upload button is designed for Google Colab.")
        return
    uploaded = files.upload()
    for filename, content in uploaded.items():
        stem = filename.rsplit(".", 1)[0]
        df = _read_upload_bytes(filename, content)
        ASSETS[stem] = AssetData(name=stem, raw=df)
    build_import_controls()
    refresh_fx_dropdown()

# ---------- 6. IMPORT CONFIGURATION UI ----------
import_output = widgets.Output()

def build_import_controls():
    with import_output:
        clear_output()
        if not ASSETS:
            print("No datasets uploaded yet.")
            return

        rows = []
        controls = {}

        for name, asset in ASSETS.items():
            detected_date = detect_date_column(asset.raw)
            detected_price = detect_price_column(asset.raw)

            currency = widgets.Dropdown(
                options=["USD", "ILS"],
                value=asset.currency if asset.currency in ["USD","ILS"] else "USD",
                description="Currency:",
                layout=widgets.Layout(width="220px")
            )
            datefmt = widgets.Dropdown(
                options=list(DATE_FORMATS.keys()),
                value=asset.date_format_label,
                description="Date format:",
                layout=widgets.Layout(width="260px")
            )
            datecol = widgets.Dropdown(
                options=list(asset.raw.columns),
                value=detected_date if detected_date in asset.raw.columns else asset.raw.columns[0],
                description="Date col:",
                layout=widgets.Layout(width="300px")
            )
            pricecol = widgets.Dropdown(
                options=list(asset.raw.columns),
                value=detected_price if detected_price in asset.raw.columns else asset.raw.columns[-1],
                description="Price col:",
                layout=widgets.Layout(width="320px")
            )

            controls[name] = (currency, datefmt, datecol, pricecol)
            rows.append(widgets.VBox([
                widgets.HTML(f"<b>{name}</b>"),
                widgets.HBox([currency, datefmt]),
                widgets.HBox([datecol, pricecol]),
            ]))

        save_btn = widgets.Button(description="Save import settings", button_style="success")
        status = widgets.Output()

        def save_settings(_):
            with status:
                clear_output()
                for name, vals in controls.items():
                    currency, datefmt, datecol, pricecol = vals
                    asset = ASSETS[name]
                    asset.currency = currency.value
                    asset.date_format_label = datefmt.value
                    asset.date_col = datecol.value
                    asset.price_col = pricecol.value
                    try:
                        s = parse_asset(asset)
                        div_note = "likely included" if "adj" in _normalize_col(asset.price_col) else "not guaranteed"
                        print(
                            f"{name}: {asset.currency}, date {asset.detected_date_format}, "
                            f"{asset.price_col}, {len(s):,} rows, "
                            f"{s.index.min().date()} → {s.index.max().date()}, "
                            f"dividends {div_note}"
                        )
                    except Exception as e:
                        print(f"{name}: ERROR — {e}")
                refresh_strategy_dropdowns()
                refresh_fx_dropdown()

        save_btn.on_click(save_settings)
        display(widgets.VBox(rows + [save_btn, status]))

# ---------- 7. FX ----------
def get_price_series(name: str, base_currency: str) -> pd.Series:
    if name not in ASSETS:
        raise KeyError(f"Unknown asset: {name}")
    asset = ASSETS[name]
    s = asset.parsed if asset.parsed is not None else parse_asset(asset)

    if asset.currency == base_currency:
        return s.copy()

    if FX_SERIES is None:
        raise ValueError(
            f"{name} is in {asset.currency} but base currency is {base_currency}. "
            "Load an FX series or use a matching base currency."
        )

    idx = s.index.union(FX_SERIES.index).sort_values()
    px = s.reindex(idx).ffill()
    fx = FX_SERIES.reindex(idx).ffill()

    # FX is ILS per 1 USD
    if asset.currency == "ILS" and base_currency == "USD":
        conv = px / fx
    elif asset.currency == "USD" and base_currency == "ILS":
        conv = px * fx
    else:
        raise ValueError("Unsupported currency conversion.")
    conv.name = name
    return conv.dropna()

def set_fx_from_asset(asset_name: str, orientation="ILS per USD"):
    global FX_SERIES, FX_NAME
    if asset_name not in ASSETS:
        raise ValueError("FX asset not found.")
    s = ASSETS[asset_name].parsed if ASSETS[asset_name].parsed is not None else parse_asset(ASSETS[asset_name])
    if orientation == "ILS per USD":
        FX_SERIES = s.copy()
    else:
        FX_SERIES = 1.0 / s
    FX_NAME = asset_name
    print(f"FX series set from {asset_name}. Internal convention: ILS per 1 USD.")

# ---------- 8. DATE ALIGNMENT ----------
def aligned_prices(asset_names: List[str], base_currency: str) -> pd.DataFrame:
    series = [get_price_series(a, base_currency) for a in asset_names]
    idx = series[0].index
    for s in series[1:]:
        idx = idx.union(s.index)
    idx = idx.sort_values()

    df = pd.concat([s.reindex(idx).ffill() for s in series], axis=1)
    df.columns = asset_names
    df = df.dropna(how="any")
    return df

# ---------- 9. SIGNAL DATES ----------
def make_signal_dates(index: pd.DatetimeIndex, frequency: str, weekday="FRI") -> pd.DatetimeIndex:
    s = pd.Series(index=index, data=np.arange(len(index)))

    if frequency == "Daily":
        return index

    if frequency == "Weekly":
        # last available observation in each weekly bucket
        return s.resample(f"W-{weekday}").last().dropna().index

    if frequency == "Monthly":
        return s.resample("ME").last().dropna().index

    raise ValueError("Unknown frequency.")

def last_price_on_or_before(prices: pd.DataFrame, dt: pd.Timestamp) -> pd.Series:
    sub = prices.loc[:dt]
    if sub.empty:
        raise ValueError("No price available before signal date.")
    return sub.iloc[-1]

# ---------- 10. STRATEGY SIGNALS ----------
def roc_on_signal(prices: pd.DataFrame, signal_dates, lookback_months: int) -> pd.DataFrame:
    records = []
    for dt in signal_dates:
        target = dt - pd.DateOffset(months=lookback_months)
        now = last_price_on_or_before(prices, dt)
        past = last_price_on_or_before(prices, target)
        roc = now / past - 1.0
        roc.name = dt
        records.append(roc)
    return pd.DataFrame(records)

def sma_value(series: pd.Series, dt: pd.Timestamp, window_days: int):
    hist = series.loc[:dt]
    if len(hist) < window_days:
        return np.nan
    return hist.iloc[-window_days:].mean()

def generate_signals(
    prices: pd.DataFrame,
    strategy: str,
    primary: str,
    secondary: Optional[str],
    frequency: str,
    lookback_months: int,
    sma_days: int,
    envelope_pct: float,
    weekly_anchor: str = "FRI",
):
    sig_dates = make_signal_dates(prices.index, frequency, weekday=weekly_anchor)
    holdings = pd.Series(index=sig_dates, dtype="object")

    if strategy in ["Absolute Momentum", "Dual Momentum", "Relative Momentum"]:
        needed = lookback_months
        valid_dates = [d for d in sig_dates if d >= prices.index.min() + pd.DateOffset(months=needed)]
        if not valid_dates:
            raise ValueError("Not enough history for selected momentum lookback.")
        rocs = roc_on_signal(prices, valid_dates, lookback_months)

        for dt in valid_dates:
            r = rocs.loc[dt]
            if strategy == "Absolute Momentum":
                holdings.loc[dt] = primary if r[primary] > 0 else "CASH"

            elif strategy == "Dual Momentum":
                if secondary is None:
                    raise ValueError("Dual Momentum requires a second asset.")
                pair = r[[primary, secondary]]
                winner = pair.idxmax()
                holdings.loc[dt] = winner if pair.max() > 0 else "CASH"

            elif strategy == "Relative Momentum":
                cols = list(prices.columns)
                holdings.loc[dt] = r[cols].idxmax()

        return holdings.dropna()

    if strategy == "Buy & Hold":
        holdings[:] = primary
        return holdings.dropna()

    if strategy == "SMA Trend":
        s = prices[primary]
        for dt in sig_dates:
            sma = sma_value(s, dt, sma_days)
            if np.isnan(sma):
                continue
            px = last_price_on_or_before(prices[[primary]], dt)[primary]
            holdings.loc[dt] = primary if px > sma else "CASH"
        return holdings.dropna()

    if strategy == "SMA Hysteresis Envelope":
        s = prices[primary]
        state = "CASH"
        for dt in sig_dates:
            sma = sma_value(s, dt, sma_days)
            if np.isnan(sma):
                continue
            px = last_price_on_or_before(prices[[primary]], dt)[primary]
            upper = sma * (1 + envelope_pct)
            lower = sma * (1 - envelope_pct)

            # hysteresis-style envelope:
            # enter risky asset only above upper band
            # exit only below lower band
            # otherwise retain previous state
            if state == "CASH" and px > upper:
                state = primary
            elif state == primary and px < lower:
                state = "CASH"
            holdings.loc[dt] = state
        return holdings.dropna()

    raise ValueError("Unknown strategy.")

def apply_market_filter(
    signals: pd.Series,
    prices: pd.DataFrame,
    filter_asset: Optional[str],
    rule: str,
    sma_days: int,
    evaluation_frequency: str = "Daily",
) -> pd.Series:
    """Force underlying strategy signals to CASH when an external filter fails."""
    if filter_asset is None or filter_asset == "None":
        return signals.copy()
    if rule != "Price > SMA":
        raise ValueError(f"Unsupported market filter rule: {rule}")
    if filter_asset not in prices.columns:
        raise ValueError(f"Market filter asset is unavailable: {filter_asset}")

    filter_state = calculate_market_filter_state(
        prices=prices,
        filter_asset=filter_asset,
        rule=rule,
        sma_days=sma_days,
        evaluation_frequency=evaluation_frequency,
    )
    filtered = signals.copy()
    for dt in filtered.index:
        latest_state = filter_state.loc[:dt]
        if latest_state.empty or not bool(latest_state.iloc[-1]):
            filtered.loc[dt] = "CASH"
    return filtered

def make_filter_evaluation_dates(
    index: pd.DatetimeIndex,
    evaluation_frequency: str,
) -> pd.DatetimeIndex:
    """Return actual aligned trading dates when the filter may change state."""
    if evaluation_frequency == "Daily":
        return index
    if evaluation_frequency not in ("Weekly", "Monthly"):
        raise ValueError(f"Unknown filter evaluation frequency: {evaluation_frequency}")

    observations = pd.Series(index=index, data=index)
    resample_rule = "W-FRI" if evaluation_frequency == "Weekly" else "ME"
    actual_dates = observations.resample(resample_rule).last().dropna().tolist()
    return pd.DatetimeIndex(actual_dates)

def calculate_market_filter_state(
    prices: pd.DataFrame,
    filter_asset: str,
    rule: str,
    sma_days: int,
    evaluation_frequency: str,
) -> pd.Series:
    """Calculate ON/OFF states without using observations after each date."""
    if rule != "Price > SMA":
        raise ValueError(f"Unsupported market filter rule: {rule}")
    if filter_asset not in prices.columns:
        raise ValueError(f"Market filter asset is unavailable: {filter_asset}")

    filter_series = prices[filter_asset]
    evaluation_dates = make_filter_evaluation_dates(prices.index, evaluation_frequency)
    states = pd.Series(index=evaluation_dates, dtype="bool", name="Market filter ON")
    for dt in evaluation_dates:
        sma = sma_value(filter_series, dt, sma_days)
        px = last_price_on_or_before(prices[[filter_asset]], dt)[filter_asset]
        states.loc[dt] = bool(pd.notna(sma) and px > sma)
    return states

def build_market_filter_diagnostics(
    prices: pd.DataFrame,
    filter_asset: str,
    rule: str,
    sma_days: int,
    evaluation_frequency: str,
    holdings: pd.Series,
    trades: pd.DataFrame,
) -> pd.DataFrame:
    """Return an auditable daily view of the state actually used by the filter.

    ``Raw price > SMA`` is intentionally calculated for every row here only for
    display. Strategy decisions continue to use the sparse, persisted state
    returned by ``calculate_market_filter_state``.
    """
    if rule != "Price > SMA":
        raise ValueError(f"Unsupported market filter rule: {rule}")
    if filter_asset not in prices.columns:
        raise ValueError(f"Market filter asset is unavailable: {filter_asset}")

    filter_series = prices[filter_asset]
    evaluation_dates = make_filter_evaluation_dates(prices.index, evaluation_frequency)
    evaluation_state = calculate_market_filter_state(
        prices=prices,
        filter_asset=filter_asset,
        rule=rule,
        sma_days=sma_days,
        evaluation_frequency=evaluation_frequency,
    )
    sma = filter_series.rolling(sma_days, min_periods=sma_days).mean()
    raw_state = (filter_series > sma).where(sma.notna())
    persisted_state = evaluation_state.reindex(prices.index).astype("boolean").ffill().fillna(False).astype(bool)

    trade_labels = pd.Series("", index=prices.index, dtype="object")
    if not trades.empty:
        labels = (
            trades.assign(_label=trades["action"] + " " + trades["asset"])
            .groupby("date")["_label"]
            .agg("; ".join)
        )
        matching_dates = trade_labels.index.intersection(labels.index)
        trade_labels.loc[matching_dates] = labels.reindex(matching_dates)

    return pd.DataFrame({
        "Date": prices.index,
        f"{filter_asset} close": filter_series.values,
        f"{filter_asset} SMA ({sma_days})": sma.values,
        "Raw price > SMA": raw_state.values,
        "Filter evaluation date": prices.index.isin(evaluation_dates),
        "Persisted filter state": persisted_state.values,
        "Current holding": holdings.reindex(prices.index, fill_value="CASH").values,
        "Trade generated": trade_labels.values,
    })

# ---------- 10B. AGITQ SIGNAL MODULE ----------
def _confirm_agitq_states(raw_states: pd.Series, confirmation_days: int):
    if confirmation_days not in (0, 2, 3):
        raise ValueError("Signal confirmation days must be 0, 2, or 3.")
    required = 1 if confirmation_days == 0 else confirmation_days
    confirmed = "OFF"
    pending = None
    pending_count = 0
    confirmed_values = []
    confirmation_labels = []
    transition_flags = []

    for raw in raw_states:
        transitioned = False
        if raw == confirmed:
            pending = None
            pending_count = 0
            label = f"Confirmed {confirmed}"
        else:
            if raw == pending:
                pending_count += 1
            else:
                pending = raw
                pending_count = 1
            if pending_count >= required:
                confirmed = raw
                pending = None
                pending_count = 0
                transitioned = True
                label = f"Confirmed {confirmed}"
            else:
                label = f"Pending {raw} {pending_count}/{required}"
        confirmed_values.append(confirmed)
        confirmation_labels.append(label)
        transition_flags.append(transitioned)

    return (
        pd.Series(confirmed_values, index=raw_states.index, dtype="object"),
        pd.Series(confirmation_labels, index=raw_states.index, dtype="object"),
        pd.Series(transition_flags, index=raw_states.index, dtype="bool"),
    )

def _hybrid_raw_states(
    short_sma: pd.Series,
    entry_sma: pd.Series,
    exit_sma: pd.Series,
) -> pd.Series:
    state = "OFF"
    output = []
    prev_short = np.nan
    prev_entry = np.nan
    prev_exit = np.nan

    for dt in short_sma.index:
        short = short_sma.loc[dt]
        entry = entry_sma.loc[dt]
        exit_ = exit_sma.loc[dt]
        ready = pd.notna(short) and pd.notna(entry) and pd.notna(exit_)
        prev_ready = pd.notna(prev_short) and pd.notna(prev_entry) and pd.notna(prev_exit)

        if ready:
            if state == "OFF":
                crossed_entry = prev_ready and prev_short <= prev_entry and short > entry
                between_now = entry <= short <= exit_ if exit_ > entry else False
                between_before = (
                    prev_entry <= prev_short <= prev_exit
                    if prev_ready and prev_exit > prev_entry else False
                )
                if exit_ > entry and between_now and not between_before:
                    state = "EXCEPTION"
                elif crossed_entry:
                    state = "NORMAL"
            elif state == "NORMAL":
                if short < exit_:
                    state = "OFF"
            elif state == "EXCEPTION":
                if short < entry:
                    state = "OFF"
                elif short > exit_:
                    state = "NORMAL"

        output.append(state)
        prev_short, prev_entry, prev_exit = short, entry, exit_

    return pd.Series(output, index=short_sma.index, dtype="object", name="raw_hold_state")

def generate_agitq_state(
    prices: pd.DataFrame,
    variant: str,
    signal_asset: str,
    traded_asset: str,
    defensive_asset: str,
    parking_asset: str,
    short_sma_days: Optional[int],
    long_sma_days: Optional[int],
    hybrid_entry_sma_days: Optional[int],
    hybrid_exit_sma_days: Optional[int],
    envelope_pct: float,
    envelope_enabled: bool,
    confirmation_days: int,
) -> pd.DataFrame:
    if variant not in AGITQ_PRESETS:
        raise ValueError(f"Unknown AGITQ variant: {variant}")
    if signal_asset not in prices.columns:
        raise ValueError(f"AGITQ signal asset is unavailable: {signal_asset}")

    signal_price = prices[signal_asset]
    state = pd.DataFrame(index=prices.index)
    state["price"] = signal_price

    if variant == "Original 200":
        if not long_sma_days or long_sma_days < 1:
            raise ValueError("Original 200 requires a positive long SMA.")
        state["long_sma"] = signal_price.rolling(long_sma_days, min_periods=long_sma_days).mean()
        raw = pd.Series("OFF", index=prices.index, dtype="object")
        prior = "OFF"
        for dt in prices.index:
            px, long_value = state.at[dt, "price"], state.at[dt, "long_sma"]
            if pd.notna(long_value):
                if px > long_value:
                    prior = "ON"
                elif px < long_value:
                    prior = "OFF"
            raw.loc[dt] = prior
        reference = state["long_sma"]
        raw_hold = raw.map({"ON": "NORMAL", "OFF": "OFF"})

    elif variant in ("QQQ 3/161", "QLD ISA", "TQQQ 5/218"):
        if not short_sma_days or not long_sma_days:
            raise ValueError(f"{variant} requires positive short and long SMAs.")
        state["short_sma"] = signal_price.rolling(short_sma_days, min_periods=short_sma_days).mean()
        state["long_sma"] = signal_price.rolling(long_sma_days, min_periods=long_sma_days).mean()
        raw = pd.Series("OFF", index=prices.index, dtype="object")
        prior = "OFF"
        for dt in prices.index:
            short_value = state.at[dt, "short_sma"]
            long_value = state.at[dt, "long_sma"]
            if pd.notna(short_value) and pd.notna(long_value):
                if short_value > long_value:
                    prior = "ON"
                elif short_value < long_value:
                    prior = "OFF"
            raw.loc[dt] = prior
        reference = state["long_sma"]
        raw_hold = raw.map({"ON": "NORMAL", "OFF": "OFF"})

    else:
        if not short_sma_days or not hybrid_entry_sma_days or not hybrid_exit_sma_days:
            raise ValueError("Hybrid 3/185/161 requires positive short, entry, and exit SMAs.")
        state["short_sma"] = signal_price.rolling(short_sma_days, min_periods=short_sma_days).mean()
        state["entry_sma"] = signal_price.rolling(
            hybrid_entry_sma_days, min_periods=hybrid_entry_sma_days
        ).mean()
        state["exit_sma"] = signal_price.rolling(
            hybrid_exit_sma_days, min_periods=hybrid_exit_sma_days
        ).mean()
        raw_hold = _hybrid_raw_states(state["short_sma"], state["entry_sma"], state["exit_sma"])
        raw = raw_hold.map(lambda value: "OFF" if value == "OFF" else "ON")
        reference = state["entry_sma"]

    confirmed_hold, confirmation_state, transitioned = _confirm_agitq_states(
        raw_hold, confirmation_days
    )
    state["reference_sma"] = reference
    state["upper_envelope"] = reference * (1 + envelope_pct)
    state["temperature"] = "Unavailable"
    ready = reference.notna()
    state.loc[ready & (signal_price < reference), "temperature"] = "Falling"
    state.loc[
        ready & (signal_price >= reference) & (signal_price <= state["upper_envelope"]),
        "temperature",
    ] = "Focus"
    state.loc[ready & (signal_price > state["upper_envelope"]), "temperature"] = "Overheated"
    state["raw_strategy_signal"] = raw
    state["raw_hold_state"] = raw_hold
    state["confirmed_hold_state"] = confirmed_hold
    state["confirmed_signal"] = confirmed_hold.map(
        lambda value: "OFF" if value == "OFF" else "ON"
    )
    state["confirmation_state"] = confirmation_state
    state["confirmation_completed"] = transitioned

    def allocation_for(row):
        if row["confirmed_signal"] == "OFF":
            return defensive_asset
        if envelope_enabled and row["temperature"] == "Overheated":
            return parking_asset
        return traded_asset

    state["new_capital_target"] = state.apply(allocation_for, axis=1)
    state["final_allocation_action"] = state.apply(
        lambda row: (
            f"OFF → {defensive_asset}"
            if row["confirmed_signal"] == "OFF"
            else f"ON; new capital → {row['new_capital_target']}"
        ),
        axis=1,
    )
    return state

# ---------- 11. EXECUTION DATE ----------
def next_available_date(index: pd.DatetimeIndex, signal_dt: pd.Timestamp, mode: str):
    pos = index.searchsorted(signal_dt)

    if mode == "Same close":
        if pos < len(index) and index[pos] == signal_dt:
            return index[pos]
        if pos == 0:
            return index[0]
        return index[pos - 1]

    # next available close
    if pos < len(index) and index[pos] <= signal_dt:
        pos += 1
    elif pos < len(index) and index[pos] == signal_dt:
        pos += 1

    if pos >= len(index):
        return None
    return index[pos]

# ---------- 11B. CANONICAL HAA SIGNALS ----------
def haa_13612u(prices: pd.DataFrame, decision_dates: pd.DatetimeIndex) -> pd.DataFrame:
    """Unweighted 1/3/6/12-month total-return momentum using month-end observations."""
    output = pd.DataFrame(index=decision_dates, columns=prices.columns, dtype=float)
    for dt in decision_dates:
        now = last_price_on_or_before(prices, dt)
        returns = []
        for months in (1, 3, 6, 12):
            then = last_price_on_or_before(prices, dt - pd.DateOffset(months=months))
            returns.append(now / then - 1.0)
        output.loc[dt] = sum(returns) / 4.0
    return output

def generate_haa_targets(prices: pd.DataFrame, strategy: str):
    """Generate fixed canonical HAA targets, observed only at each month-end."""
    if strategy == HAA_SIMPLE_STRATEGY_NAME:
        offensive = ["SPY"]
    elif strategy == HAA_BALANCED_STRATEGY_NAME:
        offensive = HAA_BALANCED_OFFENSIVE
    else:
        raise ValueError("Unknown HAA strategy.")
    assets = list(dict.fromkeys([HAA_CANARY] + offensive + HAA_DEFENSIVE))
    missing = [asset for asset in assets if asset not in prices.columns]
    if missing:
        raise ValueError("Canonical HAA requires loaded adjusted-price datasets: " + ", ".join(missing))
    dates = make_filter_evaluation_dates(prices.index, "Monthly")
    # Require the full 12-month observation history instead of substituting a
    # trading-day approximation.
    dates = pd.DatetimeIndex([dt for dt in dates if dt - pd.DateOffset(months=12) >= prices.index.min()])
    momentum = haa_13612u(prices[assets], dates)
    targets, audits = [], []
    for dt in dates:
        row = momentum.loc[dt]
        defensive = row[HAA_DEFENSIVE].idxmax()
        risk_on = bool(row[HAA_CANARY] > 0)
        weights = {asset: 0.0 for asset in assets if asset != HAA_CANARY}
        ranking = ""
        if not risk_on:
            weights[defensive] = 1.0
        elif strategy == HAA_SIMPLE_STRATEGY_NAME:
            weights["SPY" if row["SPY"] > 0 else defensive] = 1.0
        else:
            selected = row[HAA_BALANCED_OFFENSIVE].sort_values(ascending=False).head(4)
            ranking = ", ".join(selected.index)
            for asset, value in selected.items():
                weights[asset if value > 0 else defensive] += 0.25
        targets.append(weights)
        audits.append({
            "Date": dt, "TIP 13612U momentum": row[HAA_CANARY],
            "Regime": "Risk On" if risk_on else "Risk Off",
            "Best defensive asset": defensive,
            "Offensive ranking": ranking,
            "Selected holdings": ", ".join(f"{asset} {weight:.0%}" for asset, weight in weights.items() if weight),
            "Target weights": weights, **{f"{asset} 13612U": row[asset] for asset in assets},
        })
    return pd.DataFrame(targets, index=dates).fillna(0.0), pd.DataFrame(audits), momentum

def backtest_weighted_monthly(prices, targets, initial_value=100000.0, tax_rate=DEFAULT_TAX_RATE, tx_cost=DEFAULT_TRANSACTION_COST, execution_mode="Next available close"):
    """Independent weighted portfolio with monthly target-weight rebalancing and one-way costs."""
    index, assets = prices.index, list(targets.columns)
    units, cost_basis = pd.Series(0.0, index=assets), pd.Series(0.0, index=assets)
    cash, trades, equity, holdings = float(initial_value), [], pd.Series(index=index, dtype=float), pd.Series(index=index, dtype="object")
    executions = {}
    for signal_dt, weights in targets.iterrows():
        execution_dt = next_available_date(index, signal_dt, execution_mode)
        if execution_dt is not None:
            executions[execution_dt] = weights
    for dt in index:
        value = cash + float((units * prices.loc[dt, assets]).sum())
        if dt in executions:
            weights = executions[dt]
            desired = weights * value
            current = units * prices.loc[dt, assets]
            # Sell first, then buy. Rebalancing is always evaluated even when holdings persist.
            for asset in assets:
                if current[asset] > desired[asset] + 1e-10:
                    amount = current[asset] - desired[asset]
                    sold_units = amount / prices.at[dt, asset]
                    basis_sold = cost_basis[asset] * sold_units / units[asset] if units[asset] else 0.0
                    units[asset] -= sold_units
                    proceeds = amount * (1 - tx_cost)
                    tax = max(proceeds - basis_sold, 0.0) * tax_rate
                    cash += proceeds - tax
                    cost_basis[asset] -= basis_sold
                    trades.append({"date": dt, "action": "SELL", "asset": asset, "price": prices.at[dt, asset], "tax": tax, "value": proceeds - tax})
            value_after_sells = cash + float((units * prices.loc[dt, assets]).sum())
            desired = weights * value_after_sells
            current = units * prices.loc[dt, assets]
            for asset in assets:
                if desired[asset] > current[asset] + 1e-10:
                    amount = min(desired[asset] - current[asset], cash)
                    if amount > 1e-10:
                        units[asset] += amount * (1 - tx_cost) / prices.at[dt, asset]
                        cash -= amount
                        cost_basis[asset] += amount * (1 - tx_cost)
                        trades.append({"date": dt, "action": "BUY", "asset": asset, "price": prices.at[dt, asset], "tax": 0.0, "value": amount})
        equity.at[dt] = cash + float((units * prices.loc[dt, assets]).sum())
        holdings.at[dt] = ", ".join(asset for asset in assets if units[asset] > 1e-10) or "CASH"
    return equity, pd.DataFrame(trades), holdings

def run_haa_test(strategy, base_currency="USD", initial_value=100000.0, tax_rate=0.25, tx_cost=0.001, execution_mode="Next available close", start_date=None, end_date=None):
    assets = list(dict.fromkeys([HAA_CANARY] + (["SPY"] if strategy == HAA_SIMPLE_STRATEGY_NAME else HAA_BALANCED_OFFENSIVE) + HAA_DEFENSIVE))
    all_prices = aligned_prices(assets, base_currency)
    history = all_prices.loc[:pd.Timestamp(end_date)] if end_date is not None else all_prices
    targets, audit, momentum = generate_haa_targets(history, strategy)
    prices = history.loc[pd.Timestamp(start_date):] if start_date is not None else history
    targets = targets.loc[targets.index <= prices.index.max()]
    equity, trades, holdings = backtest_weighted_monthly(prices, targets, initial_value, tax_rate, tx_cost, execution_mode)
    return {"equity": equity, "trades": trades, "holdings": holdings, "metrics": calculate_metrics(equity, trades, holdings, initial_value=initial_value), "haa_audit": audit, "haa_momentum": momentum, "settings": {"Strategy": strategy, "Execution": execution_mode, "Tax %": tax_rate * 100, "Transaction cost %": tx_cost * 100, "Initial capital": initial_value}, "cash_rate": 0.0, "benchmark": None, "benchmark_metrics": None, "initial_value": initial_value, "total_contributions": 0.0, "external_flows": pd.Series(0.0, index=equity.index)}

# ---------- 12. BACKTEST ----------
def backtest(
    prices: pd.DataFrame,
    signals: pd.Series,
    initial_value: float = 100000.0,
    tax_rate: float = DEFAULT_TAX_RATE,
    tx_cost: float = DEFAULT_TRANSACTION_COST,
    cash_rate: float = DEFAULT_CASH_RATE,
    execution_mode: str = "Next available close",
):
    idx = prices.index
    value = initial_value
    current = "CASH"
    units = 0.0
    cost_basis = 0.0
    cash = initial_value
    trades = []
    equity = pd.Series(index=idx, dtype=float)
    holding_hist = pd.Series(index=idx, dtype="object")

    # translate signals into execution dates
    exec_map = {}
    for sdt, target in signals.items():
        edt = next_available_date(idx, sdt, "Same close" if execution_mode=="Same close" else "Next available close")
        if edt is not None:
            exec_map[edt] = target

    daily_cash_rate = (1 + cash_rate) ** (1/365.25) - 1

    prev_dt = idx[0]
    for dt in idx:
        days = max((dt - prev_dt).days, 0)

        # accrue cash
        if cash > 0 and days > 0:
            cash *= (1 + daily_cash_rate) ** days

        # mark to market before trade
        if current != "CASH":
            value = units * prices.loc[dt, current] + cash
        else:
            value = cash

        # execute signal
        if dt in exec_map:
            target = exec_map[dt]

            if target != current:
                # sell current asset
                if current != "CASH":
                    sell_px = prices.loc[dt, current]
                    proceeds = units * sell_px
                    proceeds *= (1 - tx_cost)

                    gain = proceeds - cost_basis
                    tax = max(gain, 0) * tax_rate
                    proceeds_after_tax = proceeds - tax
                    cash += proceeds_after_tax

                    trades.append({
                        "date": dt, "action": "SELL", "asset": current,
                        "price": sell_px, "tax": tax, "value": cash
                    })

                    units = 0.0
                    cost_basis = 0.0
                    current = "CASH"

                # buy target
                if target != "CASH":
                    buy_px = prices.loc[dt, target]
                    investable = cash * (1 - tx_cost)
                    units = investable / buy_px
                    cost_basis = investable
                    cash = 0.0
                    current = target

                    trades.append({
                        "date": dt, "action": "BUY", "asset": target,
                        "price": buy_px, "tax": 0.0, "value": units * buy_px
                    })

        # final mark to market for day
        if current != "CASH":
            value = units * prices.loc[dt, current] + cash
        else:
            value = cash

        equity.loc[dt] = value
        holding_hist.loc[dt] = current
        prev_dt = dt

    trades_df = pd.DataFrame(trades)
    return equity.dropna(), trades_df, holding_hist.dropna()

def backtest_buy_and_hold_core(
    prices: pd.DataFrame,
    asset: str,
    initial_value: float,
    tax_rate: float = DEFAULT_TAX_RATE,
    tx_cost: float = DEFAULT_TRANSACTION_COST,
    cash_rate: float = DEFAULT_CASH_RATE,
):
    """Run an independent, never-rebalanced Buy & Hold core using normal trade accounting."""
    if asset not in prices.columns:
        raise ValueError(f"Core asset is unavailable: {asset}")
    if initial_value <= 0:
        index = prices.index
        return (
            pd.Series(0.0, index=index),
            pd.DataFrame(columns=["date", "action", "asset", "price", "tax", "value"]),
            pd.Series("CASH", index=index, dtype="object"),
        )
    core_signal = pd.Series({prices.index[0]: asset}, dtype="object")
    return backtest(
        prices=prices[[asset]], signals=core_signal, initial_value=initial_value,
        tax_rate=tax_rate, tx_cost=tx_cost, cash_rate=cash_rate,
        execution_mode="Same close",
    )

def backtest_agitq(
    prices: pd.DataFrame,
    state_history: pd.DataFrame,
    traded_asset: str,
    defensive_asset: str,
    parking_asset: str,
    initial_value: float = 100000.0,
    tax_rate: float = DEFAULT_TAX_RATE,
    tx_cost: float = DEFAULT_TRANSACTION_COST,
    cash_rate: float = DEFAULT_CASH_RATE,
    confirmation_days: int = 0,
    contribution_amount: float = 0.0,
    contribution_frequency: str = "Monthly",
):
    """AGITQ-only multi-position accounting using the existing tax/cost conventions."""
    idx = prices.index
    if idx.empty:
        raise ValueError("No AGITQ prices are available in the selected period.")
    if contribution_amount < 0:
        raise ValueError("Contribution amount cannot be negative.")

    position_assets = [a for a in {traded_asset, defensive_asset, parking_asset} if a != "CASH"]
    missing = [a for a in position_assets if a not in prices.columns]
    if missing:
        raise ValueError(f"Missing AGITQ price data: {', '.join(missing)}")

    positions = {asset: {"units": 0.0, "cost_basis": 0.0} for asset in position_assets}
    cash = float(initial_value)
    trades = []
    events = []
    equity = pd.Series(index=idx, dtype=float)
    external_flows = pd.Series(0.0, index=idx, dtype=float)
    holding_hist = pd.Series(index=idx, dtype="object")

    def portfolio_value(dt):
        return cash + sum(
            position["units"] * prices.loc[dt, asset]
            for asset, position in positions.items()
        )

    def sell_all(asset, dt, reason):
        nonlocal cash
        if asset == "CASH" or asset not in positions or positions[asset]["units"] <= 0:
            return
        position = positions[asset]
        sell_px = prices.loc[dt, asset]
        gross = position["units"] * sell_px
        proceeds = gross * (1 - tx_cost)
        gain = proceeds - position["cost_basis"]
        tax = max(gain, 0) * tax_rate
        sold_units = position["units"]
        cash += proceeds - tax
        position["units"] = 0.0
        position["cost_basis"] = 0.0
        trades.append({
            "date": dt, "action": "SELL", "asset": asset,
            "price": sell_px, "units": sold_units, "tax": tax,
            "value": portfolio_value(dt), "reason": reason,
        })

    def buy_amount(asset, amount, dt, reason):
        nonlocal cash
        amount = min(float(amount), cash)
        if amount <= 0 or asset == "CASH":
            return
        buy_px = prices.loc[dt, asset]
        investable = amount * (1 - tx_cost)
        bought_units = investable / buy_px
        cash -= amount
        positions[asset]["units"] += bought_units
        positions[asset]["cost_basis"] += investable
        trades.append({
            "date": dt, "action": "BUY", "asset": asset,
            "price": buy_px, "units": bought_units, "tax": 0.0,
            "value": portfolio_value(dt), "reason": reason,
        })

    def state_row_on_or_before(dt):
        available = state_history.loc[:dt]
        if available.empty:
            return None
        return available.iloc[-1]

    first_dt = idx[0]
    transition_exec = {}
    previous_hold = None
    for signal_dt, row in state_history.iterrows():
        hold = row["confirmed_hold_state"]
        if signal_dt >= first_dt and previous_hold is not None and hold != previous_hold:
            exec_dt = next_available_date(idx, signal_dt, "Next available close")
            if exec_dt is not None:
                transition_exec.setdefault(exec_dt, []).append((signal_dt, previous_hold, hold, row))
        previous_hold = hold

    contribution_exec = {}
    if contribution_amount > 0:
        contribution_dates = make_filter_evaluation_dates(idx, contribution_frequency)
        for contribution_dt in contribution_dates:
            row = state_row_on_or_before(contribution_dt)
            if row is None:
                continue
            exec_dt = next_available_date(idx, contribution_dt, "Next available close")
            if exec_dt is not None:
                contribution_exec.setdefault(exec_dt, []).append((contribution_dt, row))

    prior_state = state_history.loc[state_history.index < first_dt]
    initial_row = prior_state.iloc[-1] if not prior_state.empty else None
    if initial_row is None or initial_row["confirmed_signal"] == "OFF":
        initial_target = defensive_asset
        initial_reason = "Initial AGITQ defensive allocation"
    else:
        initial_target = initial_row["new_capital_target"]
        initial_reason = "Initial AGITQ BUY allocation"

    daily_cash_rate = (1 + cash_rate) ** (1 / 365.25) - 1
    prev_dt = first_dt
    total_contributions = 0.0

    for dt in idx:
        days = max((dt - prev_dt).days, 0)
        if cash > 0 and days > 0:
            cash *= (1 + daily_cash_rate) ** days

        if dt == first_dt:
            buy_amount(initial_target, cash, dt, initial_reason)
            events.append({
                "date": dt, "event": "INITIAL ALLOCATION", "reason": initial_reason,
                "from_state": "CASH", "to_state": initial_target,
            })

        for signal_dt, old_hold, new_hold, row in transition_exec.get(dt, []):
            old_on = old_hold != "OFF"
            new_on = new_hold != "OFF"
            confirmation_note = " — Confirmation completed" if confirmation_days else ""

            if old_on and not new_on:
                reason = "AGITQ SELL signal" + confirmation_note
                sell_all(traded_asset, dt, reason)
                sell_all(parking_asset, dt, reason)
                buy_amount(defensive_asset, cash, dt, reason)
            elif not old_on and new_on:
                reason = "AGITQ BUY signal" + confirmation_note
                sell_all(defensive_asset, dt, reason)
                buy_amount(row["new_capital_target"], cash, dt, reason)
            else:
                reason = f"AGITQ hold-state change: {old_hold} → {new_hold}" + confirmation_note

            events.append({
                "date": dt, "signal_date": signal_dt, "event": "STATE CHANGE",
                "reason": reason, "from_state": old_hold, "to_state": new_hold,
                "temperature": row["temperature"],
            })

        for contribution_dt, row in contribution_exec.get(dt, []):
            cash += contribution_amount
            total_contributions += contribution_amount
            external_flows.loc[dt] += contribution_amount
            target = row["new_capital_target"]
            if row["confirmed_signal"] == "OFF":
                reason = f"AGITQ OFF: new contribution → {defensive_asset}"
            elif row["temperature"] == "Overheated" and target == parking_asset:
                reason = "Overheated: new contribution → S&P parking asset"
            elif row["temperature"] == "Focus":
                reason = "Focus restored: new contribution → leveraged asset"
            else:
                reason = "AGITQ ON: new contribution → leveraged asset"
            buy_amount(target, contribution_amount, dt, reason)
            events.append({
                "date": dt, "signal_date": contribution_dt, "event": "CONTRIBUTION",
                "reason": reason, "amount": contribution_amount,
                "to_state": target, "temperature": row["temperature"],
            })

        equity.loc[dt] = portfolio_value(dt)
        active = [asset for asset, p in positions.items() if p["units"] > 0]
        if cash > 0.005:
            active.append("CASH")
        holding_hist.loc[dt] = " + ".join(sorted(active)) if active else "CASH"
        prev_dt = dt

    return (
        equity.dropna(),
        pd.DataFrame(trades),
        holding_hist.dropna(),
        pd.DataFrame(events),
        total_contributions,
        external_flows,
    )

# ---------- 13. METRICS ----------
def _cash_flow_adjusted_returns(
    equity: pd.Series,
    external_flows: Optional[pd.Series] = None,
) -> pd.Series:
    """Return series with end-of-day external deposits removed from performance."""
    flows = pd.Series(0.0, index=equity.index, dtype=float)
    if external_flows is not None:
        flows = external_flows.reindex(equity.index, fill_value=0.0).astype(float)
    previous_value = equity.shift(1)
    returns = (equity - flows) / previous_value - 1
    return returns.replace([np.inf, -np.inf], np.nan).dropna()

def _xirr(cash_flows: pd.Series) -> float:
    """Annualized money-weighted return for conventional dated cash flows."""
    flows = cash_flows.groupby(level=0).sum().sort_index()
    flows = flows[flows != 0]
    if len(flows) < 2 or not (flows.lt(0).any() and flows.gt(0).any()):
        return np.nan

    origin = flows.index[0]
    year_fractions = np.array((flows.index - origin).days, dtype=float) / 365.25
    amounts = flows.to_numpy(dtype=float)

    def npv(rate):
        return float(np.sum(amounts / np.power(1.0 + rate, year_fractions)))

    low, high = -0.999999, 1.0
    low_value, high_value = npv(low), npv(high)
    while np.sign(low_value) == np.sign(high_value) and high < 1e9:
        high *= 2.0
        high_value = npv(high)
    if not np.isfinite(low_value) or not np.isfinite(high_value):
        return np.nan
    if np.sign(low_value) == np.sign(high_value):
        return np.nan

    for _ in range(200):
        midpoint = (low + high) / 2.0
        mid_value = npv(midpoint)
        if abs(mid_value) < 1e-10:
            return midpoint
        if np.sign(mid_value) == np.sign(low_value):
            low, low_value = midpoint, mid_value
        else:
            high, high_value = midpoint, mid_value
    return (low + high) / 2.0

def calculate_metrics(
    equity: pd.Series,
    trades: pd.DataFrame,
    holdings: pd.Series,
    external_flows: Optional[pd.Series] = None,
    initial_value: Optional[float] = None,
):
    if len(equity) < 2:
        raise ValueError("Not enough equity data.")

    flows = pd.Series(0.0, index=equity.index, dtype=float)
    if external_flows is not None:
        flows = external_flows.reindex(equity.index, fill_value=0.0).astype(float)

    years = (equity.index[-1] - equity.index[0]).days / 365.25
    daily_ret = _cash_flow_adjusted_returns(equity, flows)
    performance_index = pd.Series(1.0, index=equity.index, dtype=float)
    if not daily_ret.empty:
        performance_index.loc[daily_ret.index] = (1 + daily_ret).cumprod()
        performance_index = performance_index.ffill()
    cagr = performance_index.iloc[-1] ** (1 / years) - 1 if years > 0 else np.nan

    vol = daily_ret.std() * np.sqrt(252)
    sharpe = (daily_ret.mean() / daily_ret.std()) * np.sqrt(252) if daily_ret.std() > 0 else np.nan

    peak = performance_index.cummax()
    dd = performance_index / peak - 1
    max_dd = dd.min()

    annual = performance_index.resample("YE").last().pct_change().dropna()
    worst_year = annual.min() if not annual.empty else np.nan

    starting_capital = float(equity.iloc[0] if initial_value is None else initial_value)
    periodic_contributions = float(flows.sum())
    total_invested = starting_capital + periodic_contributions
    net_profit = float(equity.iloc[-1] - total_invested)
    investor_cash_flows = -flows.copy()
    investor_cash_flows.loc[equity.index[0]] -= starting_capital
    investor_cash_flows.loc[equity.index[-1]] += float(equity.iloc[-1])
    money_weighted_return = _xirr(investor_cash_flows)

    # Every executed BUY or SELL is one trade.
    trade_count = int(len(trades)) if not trades.empty else 0

    alloc = holdings.value_counts(normalize=True).sort_values(ascending=False)

    return {
        "CAGR": cagr,
        "Money-Weighted Return": money_weighted_return,
        "Max Drawdown": max_dd,
        "Volatility": vol,
        "Sharpe": sharpe,
        "Final Value": equity.iloc[-1],
        "Initial Capital": starting_capital,
        "Periodic Contributions": periodic_contributions,
        "Total Invested": total_invested,
        "Net Profit": net_profit,
        "Trades": trade_count,
        "Worst Year": worst_year,
        "Start": equity.index[0],
        "End": equity.index[-1],
        "Allocation": alloc,
        "DrawdownSeries": dd,
        "ReturnSeries": daily_ret,
        "PerformanceIndex": performance_index,
    }

# ---------- 14. BENCHMARK ----------
def benchmark_buyhold(
    prices: pd.DataFrame,
    asset: str,
    initial_value=100000,
    external_flows: Optional[pd.Series] = None,
):
    s = prices[asset].dropna()
    flows = pd.Series(0.0, index=s.index, dtype=float)
    if external_flows is not None:
        flows = external_flows.reindex(s.index, fill_value=0.0).astype(float)
    units = float(initial_value) / s.iloc[0]
    values = pd.Series(index=s.index, dtype=float)
    for dt, price in s.items():
        contribution = flows.loc[dt]
        if contribution > 0:
            units += contribution / price
        values.loc[dt] = units * price
    return values

# ---------- 15. RUN ONE TEST ----------
def run_test(
    strategy,
    primary,
    secondary=None,
    base_currency="USD",
    frequency="Monthly",
    lookback_months=12,
    sma_days=200,
    envelope_pct=0.05,
    tax_rate=0.25,
    tx_cost=0.001,
    cash_rate=0.0,
    execution_mode="Next available close",
    benchmark_asset=None,
    initial_value=100000,
    start_date=None,
    end_date=None,
    market_filter_asset=None,
    market_filter_rule="Price > SMA",
    market_filter_evaluation_frequency="Daily",
    enable_core_strategy_split=False,
    core_asset=None,
    core_allocation_pct=0.0,
):
    core_allocation_pct = float(core_allocation_pct)
    if not 0 <= core_allocation_pct <= 100:
        raise ValueError("Core allocation must be between 0% and 100%.")
    core_enabled = bool(enable_core_strategy_split and core_asset not in (None, "None") and core_allocation_pct > 0)
    if enable_core_strategy_split and core_allocation_pct > 0 and not core_enabled:
        raise ValueError("Choose a core asset when Core + Strategy Split is enabled.")
    names = [primary]
    if secondary and secondary != "None" and secondary not in names:
        names.append(secondary)
    if benchmark_asset and benchmark_asset != "None" and benchmark_asset not in names:
        names.append(benchmark_asset)

    # Preserve the original signal universe. The external filter can be any
    # loaded dataset and must never become a tradable candidate itself.
    signal_names = list(names)
    filter_asset = None if market_filter_asset in (None, "None") else market_filter_asset
    if filter_asset and filter_asset not in names:
        names.append(filter_asset)
    if core_enabled and core_asset not in names:
        names.append(core_asset)

    prices = aligned_prices(names, base_currency)
    start = pd.Timestamp(start_date).normalize() if start_date is not None else None
    end = pd.Timestamp(end_date).normalize() if end_date is not None else None
    if start is not None and end is not None and start > end:
        raise ValueError("Start date must be on or before end date.")
    # Retain earlier observations for momentum and moving-average calculations.
    history = prices.loc[:end] if end is not None else prices
    prices = history.loc[start:] if start is not None else history
    if len(prices) < 2:
        raise ValueError("Choose a date range containing at least two available trading days.")

    signals = generate_signals(
        prices=history[signal_names],
        strategy=strategy,
        primary=primary,
        secondary=None if secondary=="None" else secondary,
        frequency=frequency,
        lookback_months=lookback_months,
        sma_days=sma_days,
        envelope_pct=envelope_pct,
    )
    signals = apply_market_filter(
        signals=signals,
        prices=history,
        filter_asset=filter_asset,
        rule=market_filter_rule,
        sma_days=sma_days,
        evaluation_frequency=market_filter_evaluation_frequency,
    )

    # Carry the latest known target into the first available test close.
    # Use fresh capital and a new cost basis, not the pre-start portfolio value.
    # Keeping the original signal date lets the execution mapper distinguish
    # a prior signal from a new signal generated on the first trading day.
    prior_signals = signals.loc[signals.index < prices.index.min()]
    period_signals = signals.loc[prices.index.min():prices.index.max()]
    signals = pd.concat([prior_signals.tail(1), period_signals])

    strategy_initial_value = initial_value * (1 - core_allocation_pct / 100) if core_enabled else initial_value
    strategy_equity, strategy_trades, strategy_holdings = backtest(
        prices=prices,
        signals=signals,
        initial_value=strategy_initial_value,
        tax_rate=tax_rate,
        tx_cost=tx_cost,
        cash_rate=cash_rate,
        execution_mode=execution_mode,
    )
    core_equity = None
    core_trades = pd.DataFrame()
    core_holdings = None
    if core_enabled:
        core_equity, core_trades, core_holdings = backtest_buy_and_hold_core(
            prices=prices,
            asset=core_asset,
            initial_value=initial_value * core_allocation_pct / 100,
            tax_rate=tax_rate,
            tx_cost=tx_cost,
            cash_rate=cash_rate,
        )
        equity = strategy_equity.add(core_equity, fill_value=0.0)
        strategy_labeled_trades = strategy_trades.assign(Portion="Strategy")
        core_labeled_trades = core_trades.assign(Portion="Core")
        trades = pd.concat([strategy_labeled_trades, core_labeled_trades], ignore_index=True)
        if not trades.empty:
            trades = trades.sort_values("date", kind="stable").reset_index(drop=True)
        holdings = pd.Series(
            [f"Core: {core_holdings.loc[dt]} | Strategy: {strategy_holdings.loc[dt]}" for dt in equity.index],
            index=equity.index,
            dtype="object",
        )
    else:
        equity, trades, holdings = strategy_equity, strategy_trades, strategy_holdings
    filter_diagnostics = None
    if filter_asset:
        filter_diagnostics = build_market_filter_diagnostics(
            prices=history,
            filter_asset=filter_asset,
            rule=market_filter_rule,
            sma_days=sma_days,
            evaluation_frequency=market_filter_evaluation_frequency,
            holdings=strategy_holdings,
            trades=strategy_trades,
        )
        filter_diagnostics = filter_diagnostics.loc[
            filter_diagnostics["Date"].isin(prices.index)
        ].reset_index(drop=True)

    metrics = calculate_metrics(
        equity, trades, holdings, initial_value=initial_value
    )

    bench = None
    bench_metrics = None
    if benchmark_asset and benchmark_asset != "None":
        bench = benchmark_buyhold(prices.loc[equity.index.min():equity.index.max()], benchmark_asset, initial_value)
        # Align to equity
        bench = bench.reindex(equity.index).ffill().dropna()
        bench_metrics = calculate_metrics(
            bench,
            pd.DataFrame(),
            pd.Series("BENCH", index=bench.index),
            initial_value=initial_value,
        )
        # A buy-and-hold benchmark makes one opening purchase and no sale.
        bench_metrics["Trades"] = 1

    settings = {
        "Strategy": strategy,
        "Primary asset": primary,
        "Secondary asset": secondary if secondary not in (None, "None") else None,
        "Benchmark asset": benchmark_asset if benchmark_asset not in (None, "None") else None,
        "Base currency": base_currency,
        "Signal frequency": frequency,
        "ROC months": lookback_months,
        "SMA days": sma_days,
        "Envelope %": envelope_pct * 100,
        "Tax %": tax_rate * 100,
        "Transaction cost %": tx_cost * 100,
        "Annual cash return %": cash_rate * 100,
        "Execution": execution_mode,
        "Initial capital": initial_value,
        "Core + Strategy Split": core_enabled,
        "Core asset": core_asset if core_enabled else None,
        "Core allocation %": core_allocation_pct if core_enabled else None,
        "Strategy allocation %": 100 - core_allocation_pct if core_enabled else None,
        "Market filter asset": filter_asset,
        "Market filter rule": market_filter_rule if filter_asset else None,
        "Market filter evaluation frequency": (
            market_filter_evaluation_frequency if filter_asset else None
        ),
    }

    return {
        "prices": prices,
        "signals": signals,
        "equity": equity,
        "trades": trades,
        "holdings": holdings,
        "metrics": metrics,
        "benchmark": bench,
        "benchmark_metrics": bench_metrics,
        "external_flows": pd.Series(0.0, index=equity.index, dtype=float),
        "initial_value": initial_value,
        "total_contributions": 0.0,
        "core_enabled": core_enabled,
        "core_asset": core_asset if core_enabled else None,
        "core_allocation_pct": core_allocation_pct if core_enabled else 0.0,
        "core_final_value": float(core_equity.iloc[-1]) if core_enabled else None,
        "strategy_final_value": float(strategy_equity.iloc[-1]) if core_enabled else None,
        "settings": settings,
        "cash_rate": cash_rate,
        "market_filter_asset": filter_asset,
        "market_filter_rule": market_filter_rule if filter_asset else None,
        "market_filter_sma_days": sma_days if filter_asset else None,
        "market_filter_evaluation_frequency": (
            market_filter_evaluation_frequency if filter_asset else None
        ),
        "filter_diagnostics": filter_diagnostics,
    }

def run_agitq_test(
    variant,
    signal_asset,
    traded_asset,
    defensive_asset="SGOV",
    parking_asset="SPYM",
    short_sma_days=None,
    long_sma_days=None,
    hybrid_entry_sma_days=None,
    hybrid_exit_sma_days=None,
    envelope_pct=0.05,
    envelope_enabled=True,
    confirmation_days=0,
    partial_profit_taking=False,
    contribution_amount=0.0,
    contribution_frequency="Monthly",
    base_currency="USD",
    tax_rate=0.25,
    tx_cost=0.001,
    cash_rate=0.0,
    benchmark_asset=None,
    initial_value=100000,
    start_date=None,
    end_date=None,
):
    if partial_profit_taking:
        raise ValueError(
            "AGITQ partial profit-taking is unavailable because this backtester does not "
            "yet track tax lots and milestone bases precisely. Leave it OFF."
        )
    if defensive_asset in (None, "None", ""):
        defensive_asset = "CASH"
    selected_positions = [traded_asset, parking_asset]
    if defensive_asset != "CASH":
        selected_positions.append(defensive_asset)
    if len(set(selected_positions)) != len(selected_positions):
        raise ValueError("Traded, defensive, and parking assets must be different.")

    names = []
    for name in [signal_asset, traded_asset, defensive_asset, parking_asset, benchmark_asset]:
        if name and name not in ("None", "CASH") and name not in names:
            names.append(name)
    if not names:
        raise ValueError("Load the AGITQ assets before running the strategy.")

    all_prices = aligned_prices(names, base_currency)
    start = pd.Timestamp(start_date).normalize() if start_date is not None else None
    end = pd.Timestamp(end_date).normalize() if end_date is not None else None
    if start is not None and end is not None and start > end:
        raise ValueError("Start date must be on or before end date.")
    history = all_prices.loc[:end] if end is not None else all_prices
    prices = history.loc[start:] if start is not None else history
    if len(prices) < 2:
        raise ValueError("Choose a date range containing at least two available trading days.")

    agitq_state = generate_agitq_state(
        prices=history,
        variant=variant,
        signal_asset=signal_asset,
        traded_asset=traded_asset,
        defensive_asset=defensive_asset,
        parking_asset=parking_asset,
        short_sma_days=short_sma_days,
        long_sma_days=long_sma_days,
        hybrid_entry_sma_days=hybrid_entry_sma_days,
        hybrid_exit_sma_days=hybrid_exit_sma_days,
        envelope_pct=envelope_pct,
        envelope_enabled=envelope_enabled,
        confirmation_days=confirmation_days,
    )
    equity, trades, holdings, events, total_contributions, external_flows = backtest_agitq(
        prices=prices,
        state_history=agitq_state,
        traded_asset=traded_asset,
        defensive_asset=defensive_asset,
        parking_asset=parking_asset,
        initial_value=initial_value,
        tax_rate=tax_rate,
        tx_cost=tx_cost,
        cash_rate=cash_rate,
        confirmation_days=confirmation_days,
        contribution_amount=contribution_amount,
        contribution_frequency=contribution_frequency,
    )
    metrics = calculate_metrics(
        equity,
        trades,
        holdings,
        external_flows=external_flows,
        initial_value=initial_value,
    )

    bench = None
    bench_metrics = None
    if benchmark_asset and benchmark_asset != "None":
        bench = benchmark_buyhold(
            prices.loc[equity.index.min():equity.index.max()],
            benchmark_asset,
            initial_value,
            external_flows=external_flows,
        )
        bench = bench.reindex(equity.index).ffill().dropna()
        bench_flows = external_flows.reindex(bench.index, fill_value=0.0)
        bench_metrics = calculate_metrics(
            bench,
            pd.DataFrame(),
            pd.Series("BENCH", index=bench.index),
            external_flows=bench_flows,
            initial_value=initial_value,
        )
        # Opening purchase plus each contribution that buys more benchmark shares.
        bench_metrics["Trades"] = int(initial_value > 0) + int((bench_flows > 0).sum())

    settings = {
        "Strategy": AGITQ_STRATEGY_NAME,
        "AGITQ variant": variant,
        "Signal asset": signal_asset,
        "Traded asset": traded_asset,
        "Defensive asset": defensive_asset,
        "S&P parking asset": parking_asset,
        "Short SMA": short_sma_days,
        "Long/reference SMA": long_sma_days,
        "Hybrid entry SMA": hybrid_entry_sma_days,
        "Hybrid exit SMA": hybrid_exit_sma_days,
        "Envelope %": envelope_pct * 100,
        "Overheat envelope": "ON" if envelope_enabled else "OFF",
        "Signal confirmation days": confirmation_days,
        "Partial profit-taking": "OFF — unavailable" if not partial_profit_taking else "ON",
        "Contribution amount": contribution_amount,
        "Contribution frequency": contribution_frequency,
        "Initial capital": initial_value,
        "Tax %": tax_rate * 100,
        "Transaction cost %": tx_cost * 100,
        "Annual cash return %": cash_rate * 100,
        "Execution": "Next available close",
        "Base currency": base_currency,
        "Benchmark asset": benchmark_asset if benchmark_asset not in (None, "None") else None,
    }
    audit_state = agitq_state.loc[prices.index.min():prices.index.max()].copy()
    audit_state["actual_holding"] = holdings.reindex(audit_state.index).ffill()
    return {
        "prices": prices,
        "signals": agitq_state["confirmed_signal"],
        "equity": equity,
        "trades": trades,
        "holdings": holdings,
        "metrics": metrics,
        "benchmark": bench,
        "benchmark_metrics": bench_metrics,
        "cash_rate": cash_rate,
        "agitq_settings": settings,
        "agitq_state": audit_state,
        "agitq_events": events,
        "total_contributions": total_contributions,
        "external_flows": external_flows,
        "initial_value": initial_value,
    }

# ---------- 16. DISPLAY ----------
def pct(x):
    return "—" if pd.isna(x) else f"{x*100:.1f}%"

def money(x):
    return f"{x:,.0f}"

def show_result(result, benchmark_label=None):
    m = result["metrics"]
    has_contributions = m.get("Periodic Contributions", 0) > 0

    def formatted_metrics(metrics):
        return [
            pct(metrics["CAGR"]),
            pct(metrics["Money-Weighted Return"]),
            pct(metrics["Max Drawdown"]),
            pct(metrics["Volatility"]),
            "—" if pd.isna(metrics["Sharpe"]) else f"{metrics['Sharpe']:.2f}",
            money(metrics["Initial Capital"]),
            money(metrics["Periodic Contributions"]),
            money(metrics["Total Invested"]),
            money(metrics["Final Value"]),
            money(metrics["Net Profit"]),
            str(metrics["Trades"]),
            pct(metrics["Worst Year"]),
            f"{metrics['Start'].date()} → {metrics['End'].date()}",
        ]

    table = {
        "Metric": [
            "Annual return (TWR)" if has_contributions else "CAGR",
            "Money-weighted return (XIRR)",
            "Max drawdown", "Volatility", "Sharpe",
            "Initial capital", "Periodic contributions", "Total invested",
            "Final value", "Net profit", "Trades", "Worst year", "Period",
        ],
        "Strategy": formatted_metrics(m),
    }
    if result["benchmark_metrics"] is not None:
        table[benchmark_label or "Benchmark"] = formatted_metrics(result["benchmark_metrics"])
    display(pd.DataFrame(table))

    if result.get("agitq_settings"):
        print("\nAGITQ settings:")
        display(pd.DataFrame(
            list(result["agitq_settings"].items()), columns=["Setting", "Value"]
        ))
        if result.get("total_contributions", 0) > 0:
            print(
                f"Total periodic contributions: {money(result['total_contributions'])}. "
                "TWR statistics remove deposits from investment performance; XIRR reflects their actual timing."
            )
    elif result.get("market_filter_asset"):
        print(
            f"\nMarket filter: {result['market_filter_asset']} — "
            f"{result['market_filter_rule']} ({result['market_filter_sma_days']}-day SMA). "
            f"Filter evaluation frequency: {result['market_filter_evaluation_frequency']}. "
            "When the rule is false or the SMA is unavailable, the strategy holds CASH."
        )
        diagnostics = result.get("filter_diagnostics")
        if diagnostics is not None:
            print("\nFilter diagnostic (daily raw condition versus persisted state used by the strategy):")
            display(diagnostics)
    else:
        print("\nMarket filter: None")

    print(
        f"\nAnnual cash return while the strategy holds CASH: "
        f"{result.get('cash_rate', 0.0) * 100:.2f}% "
        "(included in the equity curve and strategy metrics)."
    )

    print("\nApproximate time in each holding:")
    alloc = m["Allocation"].rename("Share").to_frame()
    alloc["Share"] = alloc["Share"].map(lambda x: f"{x*100:.1f}%")
    display(alloc)

    plt.figure(figsize=(11,5))
    normalized = m["PerformanceIndex"] * 100
    plt.plot(normalized.index, normalized.values, label="Strategy")
    if result["benchmark"] is not None:
        b = result["benchmark_metrics"]["PerformanceIndex"] * 100
        plt.plot(b.index, b.values, label=benchmark_label or "Benchmark")
    plt.title("Time-weighted growth of 100" if has_contributions else "Growth of 100")
    plt.ylabel("Value")
    plt.legend()
    plt.grid(alpha=0.2)
    plt.show()

    plt.figure(figsize=(11,4))
    dd = m["DrawdownSeries"] * 100
    plt.plot(dd.index, dd.values)
    plt.title("Drawdown")
    plt.ylabel("%")
    plt.grid(alpha=0.2)
    plt.show()

    if not result["trades"].empty:
        print("\nRecent trades:")
        display(result["trades"].tail(15))

    if result.get("agitq_settings"):
        if not result["agitq_events"].empty:
            print("\nRecent AGITQ events:")
            display(result["agitq_events"].tail(20))
        print("\nRecent daily AGITQ state (full history is included in the export):")
        display(result["agitq_state"].tail(20))

        export_btn = widgets.Button(description="Export AGITQ audit ZIP", button_style="info")
        export_out = widgets.Output()

        def _export_clicked(_):
            with export_out:
                clear_output()
                try:
                    export_agitq_result(result)
                except Exception as exc:
                    print("ERROR:", exc)

        export_btn.on_click(_export_clicked)
        display(export_btn, export_out)

    print(
        "\nApproximation note: results depend on the selected price data. "
        "Dividends are only included if the chosen price series includes them. "
        "Tax is simplified to 25% of realized positive gains on each sale."
    )

def export_agitq_result(result):
    if not result.get("agitq_settings"):
        raise ValueError("The current result is not an AGITQ backtest.")
    export_dir = tempfile.mkdtemp(prefix="agitq_export_")
    settings_path = f"{export_dir}/agitq_settings.csv"
    pd.DataFrame(
        list(result["agitq_settings"].items()), columns=["Setting", "Value"]
    ).to_csv(settings_path, index=False)
    result["agitq_state"].to_csv(f"{export_dir}/agitq_daily_state.csv", index_label="date")
    result["trades"].to_csv(f"{export_dir}/agitq_trades.csv", index=False)
    result["agitq_events"].to_csv(f"{export_dir}/agitq_events.csv", index=False)
    result["equity"].rename("equity").to_csv(f"{export_dir}/agitq_equity.csv", index_label="date")
    pd.DataFrame({
        "equity": result["equity"],
        "external_flow": result["external_flows"].reindex(result["equity"].index, fill_value=0.0),
        "cash_flow_adjusted_return": result["metrics"]["ReturnSeries"],
        "time_weighted_index": result["metrics"]["PerformanceIndex"],
    }).to_csv(f"{export_dir}/agitq_performance.csv", index_label="date")
    metrics_export = {
        key: value for key, value in result["metrics"].items()
        if key not in (
            "Allocation", "DrawdownSeries", "ReturnSeries", "PerformanceIndex"
        )
    }
    pd.DataFrame(list(metrics_export.items()), columns=["Metric", "Value"]).to_csv(
        f"{export_dir}/agitq_metrics.csv", index=False
    )
    zip_path = f"{export_dir}/AGITQ_backtest_audit.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for filename in [
            "agitq_settings.csv", "agitq_daily_state.csv", "agitq_trades.csv",
            "agitq_events.csv", "agitq_equity.csv", "agitq_performance.csv",
            "agitq_metrics.csv",
        ]:
            archive.write(f"{export_dir}/{filename}", arcname=filename)
    if IN_COLAB:
        files.download(zip_path)
    else:
        display(FileLink(zip_path))
    return zip_path

# ---------- 17. RESULT HISTORY ----------
RESULT_COLUMNS = {
    "Saved": "created_at",
    "Run type": "run_type",
    "Strategy": "strategy",
    "Primary asset": "primary_asset",
    "Secondary asset": "secondary_asset",
    "Benchmark asset": "benchmark_asset",
    "CAGR / annual return": "cagr",
    "Max drawdown": "max_drawdown",
    "Volatility": "volatility",
    "Sharpe": "sharpe",
    "Final value": "final_value",
    "Trades": "trades",
    "Worst year": "worst_year",
    "Period start": "period_start",
    "Period end": "period_end",
}

def _json_default(value):
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()
    if isinstance(value, np.generic):
        return value.item()
    return str(value)

def _results_connection():
    connection = sqlite3.connect(RESULTS_DB_PATH)
    connection.execute("""
        CREATE TABLE IF NOT EXISTS backtest_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            run_type TEXT NOT NULL,
            strategy TEXT NOT NULL,
            primary_asset TEXT,
            secondary_asset TEXT,
            benchmark_asset TEXT,
            assets_json TEXT NOT NULL,
            settings_json TEXT NOT NULL,
            period_start TEXT NOT NULL,
            period_end TEXT NOT NULL,
            cagr REAL,
            max_drawdown REAL,
            volatility REAL,
            sharpe REAL,
            final_value REAL,
            trades INTEGER,
            worst_year REAL
        )
    """)
    connection.execute("CREATE INDEX IF NOT EXISTS idx_results_strategy ON backtest_results(strategy)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_results_period ON backtest_results(period_start, period_end)")
    connection.commit()
    return connection

def _result_assets_and_settings(result):
    settings = dict(result.get("settings") or result.get("agitq_settings") or {})
    if result.get("agitq_settings"):
        assets = {
            "signal": settings.get("Signal asset"),
            "traded": settings.get("Traded asset"),
            "defensive": settings.get("Defensive asset"),
            "parking": settings.get("S&P parking asset"),
            "benchmark": settings.get("Benchmark asset"),
        }
        primary = settings.get("Traded asset")
        secondary = settings.get("Signal asset")
    else:
        assets = {
            "primary": settings.get("Primary asset"),
            "secondary": settings.get("Secondary asset"),
            "benchmark": settings.get("Benchmark asset"),
            "market_filter": settings.get("Market filter asset"),
        }
        primary = settings.get("Primary asset")
        secondary = settings.get("Secondary asset")
    assets = {key: value for key, value in assets.items() if value not in (None, "", "None")}
    return assets, settings, primary, secondary, settings.get("Benchmark asset")

def save_backtest_result(result, run_type="Individual", batch_context=None):
    """Persist a completed test and its full settings without changing its results."""
    metrics = result["metrics"]
    assets, settings, primary, secondary, benchmark = _result_assets_and_settings(result)
    if batch_context:
        settings["Batch context"] = batch_context
    created_at = datetime.now().astimezone().isoformat(timespec="seconds")
    row = (
        created_at,
        run_type,
        settings.get("Strategy", "Unknown"),
        primary,
        secondary,
        benchmark,
        json.dumps(assets, default=_json_default, sort_keys=True),
        json.dumps(settings, default=_json_default, sort_keys=True),
        str(metrics["Start"].date()),
        str(metrics["End"].date()),
        float(metrics["CAGR"]) if pd.notna(metrics["CAGR"]) else None,
        float(metrics["Max Drawdown"]) if pd.notna(metrics["Max Drawdown"]) else None,
        float(metrics["Volatility"]) if pd.notna(metrics["Volatility"]) else None,
        float(metrics["Sharpe"]) if pd.notna(metrics["Sharpe"]) else None,
        float(metrics["Final Value"]),
        int(metrics["Trades"]),
        float(metrics["Worst Year"]) if pd.notna(metrics["Worst Year"]) else None,
    )
    with _results_connection() as connection:
        cursor = connection.execute("""
            INSERT INTO backtest_results (
                created_at, run_type, strategy, primary_asset, secondary_asset,
                benchmark_asset, assets_json, settings_json, period_start, period_end,
                cagr, max_drawdown, volatility, sharpe, final_value, trades, worst_year
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, row)
        return cursor.lastrowid

def _parse_numeric_filter(expression, ratio=False):
    expression = (expression or "").strip()
    if not expression:
        return None
    matched = re.fullmatch(r"(<=|>=|=|<|>)\s*(-?(?:\d+(?:\.\d*)?|\.\d+))\s*(%)?", expression)
    if not matched:
        raise ValueError(f"Invalid numeric filter: {expression!r}. Use, for example, > 10% or >= 0.6.")
    operator, raw_value, percent = matched.groups()
    value = float(raw_value)
    if ratio and (percent or abs(value) > 1):
        value /= 100
    return operator, value

def load_saved_backtests(
    strategy=None,
    asset_text=None,
    parameter_text=None,
    start_on_or_after=None,
    end_on_or_before=None,
    numeric_filters=None,
    sort_by="created_at",
    ascending=False,
):
    if sort_by not in RESULT_COLUMNS.values():
        raise ValueError("Unknown history sort column.")
    clauses, values = [], []
    if strategy and strategy != "All strategies":
        clauses.append("strategy = ?")
        values.append(strategy)
    if asset_text and asset_text.strip():
        clauses.append("assets_json LIKE ?")
        values.append(f"%{asset_text.strip()}%")
    if parameter_text and parameter_text.strip():
        clauses.append("settings_json LIKE ?")
        values.append(f"%{parameter_text.strip()}%")
    if start_on_or_after:
        clauses.append("period_start >= ?")
        values.append(str(pd.Timestamp(start_on_or_after).date()))
    if end_on_or_before:
        clauses.append("period_end <= ?")
        values.append(str(pd.Timestamp(end_on_or_before).date()))
    for column, expression in (numeric_filters or {}).items():
        if column not in {"cagr", "max_drawdown", "volatility", "sharpe", "trades", "final_value", "worst_year"}:
            raise ValueError(f"Unknown numeric history filter: {column}")
        parsed = _parse_numeric_filter(expression, ratio=column in {"cagr", "max_drawdown", "volatility", "worst_year"})
        if parsed:
            operator, value = parsed
            clauses.append(f"{column} {operator} ?")
            values.append(value)
    where = " WHERE " + " AND ".join(clauses) if clauses else ""
    order = "ASC" if ascending else "DESC"
    query = "SELECT * FROM backtest_results" + where + f" ORDER BY {sort_by} {order}, id {order}"
    with _results_connection() as connection:
        return pd.read_sql_query(query, connection, params=values)

def delete_saved_backtests(result_ids):
    ids = [int(value) for value in result_ids]
    if not ids:
        return 0
    placeholders = ", ".join("?" for _ in ids)
    with _results_connection() as connection:
        cursor = connection.execute(f"DELETE FROM backtest_results WHERE id IN ({placeholders})", ids)
        return cursor.rowcount

def saved_result_strategies():
    with _results_connection() as connection:
        rows = connection.execute(
            "SELECT DISTINCT strategy FROM backtest_results ORDER BY strategy"
        ).fetchall()
    return [row[0] for row in rows]

# ---------- 18. BATCH TEST ----------
STANDARD_BATCH_STRATEGIES = [
    "Buy & Hold",
    "Absolute Momentum",
    "Dual Momentum",
    "Relative Momentum",
    "SMA Trend",
    "SMA Hysteresis Envelope",
]

def parse_batch_values(value_text, converter, label):
    """Parse comma-separated existing setting values for the Batch Test controls."""
    values = [part.strip() for part in (value_text or "").split(",") if part.strip()]
    if not values:
        raise ValueError(f"Enter at least one {label} value.")
    try:
        return [converter(value) for value in values]
    except ValueError as exc:
        raise ValueError(f"Invalid {label} value. Use comma-separated numbers.") from exc

def batch_combination_count(**selections):
    """Return the exact Cartesian-product size before executing a batch."""
    count = 1
    for values in selections.values():
        count *= len(values)
    return count

def run_batch_test(
    strategies,
    primary_assets,
    secondary_assets,
    base_currencies,
    frequencies,
    lookbacks,
    sma_days_values,
    envelope_pcts,
    tax_rates,
    transaction_costs,
    cash_rates,
    initial_values,
    execution_modes,
    benchmark_assets,
    filter_assets,
    filter_frequencies,
    start_date=None,
    end_date=None,
):
    """Run selected standard-strategy settings through the normal ``run_test`` engine."""
    selections = {
        "strategies": strategies, "primary_assets": primary_assets,
        "secondary_assets": secondary_assets, "base_currencies": base_currencies,
        "frequencies": frequencies, "lookbacks": lookbacks,
        "sma_days_values": sma_days_values, "envelope_pcts": envelope_pcts,
        "tax_rates": tax_rates, "transaction_costs": transaction_costs,
        "cash_rates": cash_rates, "initial_values": initial_values,
        "execution_modes": execution_modes, "benchmark_assets": benchmark_assets,
        "filter_assets": filter_assets, "filter_frequencies": filter_frequencies,
    }
    empty = [name for name, values in selections.items() if not values]
    if empty:
        raise ValueError(f"Batch Test requires a selection for: {', '.join(empty)}")
    unsupported = set(strategies) - set(STANDARD_BATCH_STRATEGIES)
    if unsupported:
        raise ValueError("Batch Test supports standard strategies only; run AGITQ individually.")

    rows = []
    for values in product(*selections.values()):
        (
            strategy, primary, secondary, base_currency, frequency, lookback,
            sma_days, envelope_pct, tax_rate, tx_cost, cash_rate, initial_value,
            execution_mode, benchmark, filter_asset, filter_frequency,
        ) = values
        parameters = {
            "Signal frequency": frequency,
            "ROC months": lookback,
            "SMA days": sma_days,
            "Envelope %": envelope_pct * 100,
            "Tax %": tax_rate * 100,
            "Transaction cost %": tx_cost * 100,
            "Cash return %": cash_rate * 100,
            "Initial capital": initial_value,
            "Execution": execution_mode,
            "Base": base_currency,
            "Filter": filter_asset,
            "Filter evaluation": filter_frequency if filter_asset != "None" else None,
        }
        assets = ", ".join(
            value for value in [primary, None if secondary == "None" else secondary,
                                None if benchmark == "None" else f"Benchmark: {benchmark}",
                                None if filter_asset == "None" else f"Filter: {filter_asset}"]
            if value
        )
        row = {
            "Strategy": strategy,
            "Assets": assets,
            "Parameters": "; ".join(f"{key}={value}" for key, value in parameters.items() if value is not None),
        }
        try:
            result = run_test(
                strategy=strategy, primary=primary, secondary=secondary,
                base_currency=base_currency, frequency=frequency,
                lookback_months=lookback, sma_days=sma_days, envelope_pct=envelope_pct,
                tax_rate=tax_rate, tx_cost=tx_cost, cash_rate=cash_rate,
                execution_mode=execution_mode, benchmark_asset=benchmark,
                initial_value=initial_value, start_date=start_date, end_date=end_date,
                market_filter_asset=filter_asset,
                market_filter_evaluation_frequency=filter_frequency,
            )
            metrics = result["metrics"]
            row.update({
                "Test period": f"{metrics['Start'].date()} → {metrics['End'].date()}",
                "CAGR": metrics["CAGR"],
                "Max drawdown": metrics["Max Drawdown"],
                "Volatility": metrics["Volatility"],
                "Sharpe": metrics["Sharpe"],
                "Final value": metrics["Final Value"],
                "Trades": metrics["Trades"],
                "Worst year": metrics["Worst Year"],
                "Error": "",
            })
        except Exception as exc:
            row.update({
                "Test period": "", "CAGR": np.nan, "Max drawdown": np.nan,
                "Volatility": np.nan, "Sharpe": np.nan, "Final value": np.nan,
                "Trades": np.nan, "Worst year": np.nan, "Error": str(exc),
            })
        rows.append(row)
    return pd.DataFrame(rows)

def filter_and_sort_batch_results(frame, sort_by="CAGR", ascending=False, metric=None, expression=None, text=""):
    """Filter Batch Test results by any displayed metric and sort any displayed column."""
    output = frame.copy()
    if text and text.strip():
        output = output[output.astype(str).apply(
            lambda row: row.str.contains(text.strip(), case=False, na=False).any(), axis=1
        )]
    if metric and expression and expression.strip():
        parsed = _parse_numeric_filter(
            expression,
            ratio=metric in {"CAGR", "Max drawdown", "Volatility", "Worst year"},
        )
        if parsed:
            operator, value = parsed
            output = output.loc[getattr(output[metric], {
                ">": "gt", ">=": "ge", "<": "lt", "<=": "le", "=": "eq",
            }[operator])(value)]
    return output.sort_values(sort_by, ascending=ascending, na_position="last", kind="stable")

def batch_result_highlights(frame):
    """Return the best valid result for each requested Batch Test highlight."""
    valid = frame[frame["Error"].eq("")].copy()
    if valid.empty:
        return pd.DataFrame(columns=["Highlight", "Strategy", "Assets", "Parameters", "Value"])
    choices = [
        ("Highest CAGR", "CAGR", "max"),
        ("Lowest max drawdown", "Max drawdown", "min"),
        ("Highest Sharpe", "Sharpe", "max"),
    ]
    rows = []
    for label, metric, direction in choices:
        series = valid[metric].dropna()
        if series.empty:
            continue
        selected = valid.loc[series.idxmax() if direction == "max" else series.idxmin()]
        rows.append({
            "Highlight": label, "Strategy": selected["Strategy"], "Assets": selected["Assets"],
            "Parameters": selected["Parameters"], "Value": selected[metric],
        })
    return pd.DataFrame(rows)

def batch_momentum_test(
    primary,
    secondary_options,
    lookbacks=(10,11,12,13),
    strategy="Dual Momentum",
    base_currency="USD",
    frequency="Monthly",
    tax_rate=0.25,
    tx_cost=0.001,
    cash_rate=0.0,
    execution_mode="Next available close",
    sma_days=200,
    market_filter_asset=None,
    market_filter_rule="Price > SMA",
    market_filter_evaluation_frequency="Daily",
):
    rows = []
    for secondary in secondary_options:
        for lb in lookbacks:
            try:
                r = run_test(
                    strategy=strategy,
                    primary=primary,
                    secondary=secondary,
                    base_currency=base_currency,
                    frequency=frequency,
                    lookback_months=lb,
                    tax_rate=tax_rate,
                    tx_cost=tx_cost,
                    cash_rate=cash_rate,
                    execution_mode=execution_mode,
                    benchmark_asset=primary,
                    sma_days=sma_days,
                    market_filter_asset=market_filter_asset,
                    market_filter_rule=market_filter_rule,
                    market_filter_evaluation_frequency=market_filter_evaluation_frequency,
                )
                save_backtest_result(
                    r,
                    run_type="Batch",
                    batch_context={
                        "Secondary asset": secondary,
                        "ROC months": lb,
                    },
                )
                m = r["metrics"]
                rows.append({
                    "Primary": primary,
                    "Secondary": secondary,
                    "Lookback": lb,
                    "CAGR": m["CAGR"],
                    "MaxDD": m["Max Drawdown"],
                    "Sharpe": m["Sharpe"],
                    "Trades": m["Trades"],
                    "FinalValue": m["Final Value"],
                })
            except Exception as e:
                rows.append({
                    "Primary": primary,
                    "Secondary": secondary,
                    "Lookback": lb,
                    "Error": str(e)
                })

    out = pd.DataFrame(rows)
    if "CAGR" in out:
        out = out.sort_values("CAGR", ascending=False, na_position="last")
    return out

# ---------- 18. SIMPLE MAIN UI ----------
strategy_dd = widgets.Dropdown(
    options=[
        "Buy & Hold",
        "Absolute Momentum",
        "Dual Momentum",
        "Relative Momentum",
        "SMA Trend",
        "SMA Hysteresis Envelope",
        HAA_SIMPLE_STRATEGY_NAME,
        HAA_BALANCED_STRATEGY_NAME,
        AGITQ_STRATEGY_NAME,
    ],
    description="Strategy:"
)

primary_dd = widgets.Dropdown(options=[""], description="Primary:")
secondary_dd = widgets.Dropdown(options=["None"], description="Second:")
benchmark_dd = widgets.Dropdown(options=["None"], description="Benchmark:")
market_filter_asset_dd = widgets.Dropdown(options=["None"], value="None", description="Market filter:")
market_filter_rule_dd = widgets.Dropdown(
    options=["Price > SMA"], value="Price > SMA", description="Filter rule:"
)
market_filter_frequency_dd = widgets.Dropdown(
    options=["Daily", "Weekly", "Monthly"],
    value="Daily",
    description="Filter evaluation frequency:",
    style={"description_width": "initial"},
    layout=widgets.Layout(width="330px"),
)

base_dd = widgets.Dropdown(options=["USD","ILS"], value="USD", description="Base:")
freq_dd = widgets.Dropdown(options=["Daily","Weekly","Monthly"], value="Monthly", description="Signal:")
execution_dd = widgets.Dropdown(
    options=["Next available close","Same close"],
    value="Next available close",
    description="Execution:"
)

lookback_slider = widgets.IntSlider(value=12, min=1, max=24, step=1, description="ROC months:")
sma_slider = widgets.IntSlider(value=200, min=20, max=300, step=1, description="SMA days:")
envelope_slider = widgets.FloatSlider(value=5.0, min=0.0, max=15.0, step=0.5, description="Envelope %:")

agitq_variant_dd = widgets.Dropdown(
    options=list(AGITQ_PRESETS.keys()), value="Original 200", description="Strategy variant:"
)
agitq_signal_asset_dd = widgets.Dropdown(options=["TQQQ"], value="TQQQ", description="Signal asset:")
agitq_traded_asset_dd = widgets.Dropdown(options=["TQQQ"], value="TQQQ", description="Traded asset:")
agitq_defensive_asset_dd = widgets.Dropdown(options=["SGOV", "CASH"], value="SGOV", description="Defensive:")
agitq_parking_asset_dd = widgets.Dropdown(options=["SPYM"], value="SPYM", description="S&P parking:")
agitq_short_sma_box = widgets.IntText(value=3, description="Short SMA:")
agitq_long_sma_box = widgets.IntText(value=200, description="Long SMA:")
agitq_entry_sma_box = widgets.IntText(value=185, description="Entry SMA:")
agitq_exit_sma_box = widgets.IntText(value=161, description="Exit SMA:")
agitq_envelope_box = widgets.FloatText(value=5.0, description="Envelope %:")
agitq_envelope_enabled_cb = widgets.Checkbox(value=True, description="Overheat envelope ON")
agitq_confirmation_dd = widgets.Dropdown(options=[0, 2, 3], value=0, description="Confirmation days:")
agitq_profit_taking_cb = widgets.Checkbox(
    value=False, description="Partial profit-taking (unavailable)", disabled=True
)
agitq_contribution_box = widgets.FloatText(value=0.0, description="Contribution:")
agitq_contribution_frequency_dd = widgets.Dropdown(
    options=["Monthly", "Weekly", "Daily"], value="Monthly", description="Contribution frequency:"
)

agitq_short_row = widgets.HBox([agitq_short_sma_box, agitq_long_sma_box])
agitq_hybrid_row = widgets.HBox([agitq_entry_sma_box, agitq_exit_sma_box])
agitq_controls_box = widgets.VBox([
    widgets.HTML("<h4>AGITQ / TQQQ Playbook settings</h4>"),
    agitq_variant_dd,
    widgets.HBox([agitq_signal_asset_dd, agitq_traded_asset_dd]),
    widgets.HBox([agitq_defensive_asset_dd, agitq_parking_asset_dd]),
    agitq_short_row,
    agitq_hybrid_row,
    widgets.HBox([agitq_envelope_box, agitq_envelope_enabled_cb, agitq_confirmation_dd]),
    widgets.HBox([agitq_contribution_box, agitq_contribution_frequency_dd]),
    agitq_profit_taking_cb,
    widgets.HTML(
        "AGITQ uses daily adjusted closes and executes confirmed signals at the next available close. "
        "Partial profit-taking remains disabled because exact tax-lot milestone accounting is not available."
    ),
])
agitq_controls_box.layout.display = "none"

market_filter_controls_box = widgets.VBox([
    widgets.HBox([market_filter_asset_dd, market_filter_rule_dd, market_filter_frequency_dd]),
    widgets.HTML(
        "The market filter uses the SMA days setting below and forces CASH when the selected asset is not above its SMA. "
        "Filter evaluation frequency controls when its state can update; it does not change the strategy signal frequency."
    ),
])

tax_box = widgets.FloatText(value=25.0, description="Tax %:")
fee_box = widgets.FloatText(value=0.1, description="Trade cost %:")
cash_box = widgets.FloatText(value=0.0, description="Cash return %:")
initial_box = widgets.FloatText(value=100000.0, description="Start value:")

start_date_picker = widgets.DatePicker(description="Start date:")
end_date_picker = widgets.DatePicker(description="End date:")

run_btn = widgets.Button(description="Run backtest", button_style="success")
result_output = widgets.Output()

# Batch Test uses the same standard-strategy engine as an individual run. Values
# entered in the text controls are comma-separated, for example: ``6, 12, 18``.
batch_strategies_sm = widgets.SelectMultiple(
    options=STANDARD_BATCH_STRATEGIES, value=("Dual Momentum",), description="Strategies:",
    layout=widgets.Layout(width="330px", height="130px"),
)
batch_primary_sm = widgets.SelectMultiple(options=[], description="Primary assets:", layout=widgets.Layout(width="260px", height="100px"))
batch_secondary_sm = widgets.SelectMultiple(options=["None"], value=("None",), description="Second assets:", layout=widgets.Layout(width="260px", height="100px"))
batch_benchmark_sm = widgets.SelectMultiple(options=["None"], value=("None",), description="Benchmarks:", layout=widgets.Layout(width="260px", height="100px"))
batch_filter_asset_sm = widgets.SelectMultiple(options=["None"], value=("None",), description="Filter assets:", layout=widgets.Layout(width="260px", height="100px"))
batch_base_sm = widgets.SelectMultiple(options=["USD", "ILS"], value=("USD",), description="Base:")
batch_frequency_sm = widgets.SelectMultiple(options=["Daily", "Weekly", "Monthly"], value=("Monthly",), description="Signal frequency:")
batch_execution_sm = widgets.SelectMultiple(options=["Next available close", "Same close"], value=("Next available close",), description="Execution:")
batch_filter_frequency_sm = widgets.SelectMultiple(options=["Daily", "Weekly", "Monthly"], value=("Daily",), description="Filter frequency:")
batch_lookbacks_box = widgets.Text(value="12", description="ROC months:", placeholder="6, 12, 18")
batch_sma_days_box = widgets.Text(value="200", description="SMA days:", placeholder="100, 200")
batch_envelopes_box = widgets.Text(value="5", description="Envelope %:", placeholder="0, 5")
batch_tax_box = widgets.Text(value="25", description="Tax %:", placeholder="0, 25")
batch_fee_box = widgets.Text(value="0.1", description="Trade cost %:", placeholder="0, 0.1")
batch_cash_box = widgets.Text(value="0", description="Cash return %:", placeholder="0, 4")
batch_initial_box = widgets.Text(value="100000", description="Start value:", placeholder="50000, 100000")
batch_start_picker = widgets.DatePicker(description="Start date:")
batch_end_picker = widgets.DatePicker(description="End date:")
batch_count_html = widgets.HTML()
batch_run_btn = widgets.Button(description="Run Batch Test", button_style="success")
batch_sort_dd = widgets.Dropdown(
    options=["Strategy", "Assets", "Parameters", "Test period", "CAGR", "Max drawdown", "Volatility", "Sharpe", "Final value", "Trades", "Worst year", "Error"],
    value="CAGR", description="Sort by:",
)
batch_ascending_cb = widgets.Checkbox(value=False, description="Ascending")
batch_metric_filter_dd = widgets.Dropdown(
    options=["CAGR", "Max drawdown", "Volatility", "Sharpe", "Final value", "Trades", "Worst year"],
    value="CAGR", description="Metric filter:",
)
batch_metric_filter_box = widgets.Text(description="Condition:", placeholder="> 10%")
batch_text_filter_box = widgets.Text(description="Text filter:", placeholder="Strategy, asset, or parameter")
batch_apply_filters_btn = widgets.Button(description="Apply sort / filters", button_style="info")
batch_output = widgets.Output()
batch_results = pd.DataFrame()

history_strategy_dd = widgets.Dropdown(options=["All strategies"], description="Strategy:")
history_asset_box = widgets.Text(description="Asset contains:", placeholder="QQQ")
history_parameter_box = widgets.Text(description="Setting contains:", placeholder="SMA days")
history_start_picker = widgets.DatePicker(description="Start on/after:")
history_end_picker = widgets.DatePicker(description="End on/before:")
history_cagr_box = widgets.Text(description="CAGR:", placeholder="> 10%")
history_drawdown_box = widgets.Text(description="Max DD:", placeholder="> -30%")
history_sharpe_box = widgets.Text(description="Sharpe:", placeholder="> 0.6")
history_volatility_box = widgets.Text(description="Volatility:", placeholder="< 20%")
history_trades_box = widgets.Text(description="Trades:", placeholder="< 20")
history_final_value_box = widgets.Text(description="Final value:", placeholder="> 100000")
history_sort_dd = widgets.Dropdown(
    options=[(label, column) for label, column in RESULT_COLUMNS.items()],
    value="created_at",
    description="Sort by:",
)
history_ascending_cb = widgets.Checkbox(value=False, description="Ascending")
history_refresh_btn = widgets.Button(description="Apply filters / refresh", button_style="info")
history_clear_btn = widgets.Button(description="Clear all filters")
history_export_filtered_btn = widgets.Button(description="Export filtered CSV")
history_export_all_btn = widgets.Button(description="Export all CSV")
history_delete_btn = widgets.Button(description="Delete selected", button_style="danger")
history_selection = widgets.SelectMultiple(
    options=[], description="Select results:", layout=widgets.Layout(width="100%", height="130px")
)
history_status = widgets.HTML()
history_output = widgets.Output()

def _history_numeric_filters():
    return {
        "cagr": history_cagr_box.value,
        "max_drawdown": history_drawdown_box.value,
        "sharpe": history_sharpe_box.value,
        "volatility": history_volatility_box.value,
        "trades": history_trades_box.value,
        "final_value": history_final_value_box.value,
    }

def _history_query():
    return load_saved_backtests(
        strategy=history_strategy_dd.value,
        asset_text=history_asset_box.value,
        parameter_text=history_parameter_box.value,
        start_on_or_after=history_start_picker.value,
        end_on_or_before=history_end_picker.value,
        numeric_filters=_history_numeric_filters(),
        sort_by=history_sort_dd.value,
        ascending=history_ascending_cb.value,
    )

def refresh_results_history_controls():
    current = history_strategy_dd.value
    options = ["All strategies"] + saved_result_strategies()
    history_strategy_dd.options = options
    history_strategy_dd.value = current if current in options else "All strategies"

def _display_history_dataframe(frame):
    columns = [
        "id", "created_at", "run_type", "strategy", "primary_asset", "secondary_asset",
        "benchmark_asset", "period_start", "period_end", "cagr", "max_drawdown",
        "volatility", "sharpe", "final_value", "trades", "worst_year",
    ]
    display_frame = frame.reindex(columns=columns).copy()
    for column in ["cagr", "max_drawdown", "volatility", "worst_year"]:
        display_frame[column] = display_frame[column].map(
            lambda value: "—" if pd.isna(value) else f"{value * 100:.2f}%"
        )
    display_frame["sharpe"] = display_frame["sharpe"].map(
        lambda value: "—" if pd.isna(value) else f"{value:.2f}"
    )
    display_frame["final_value"] = display_frame["final_value"].map(
        lambda value: "—" if pd.isna(value) else f"{value:,.0f}"
    )
    display_frame.columns = [
        "ID", "Saved", "Run type", "Strategy", "Primary", "Secondary", "Benchmark",
        "Start", "End", "CAGR", "Max DD", "Volatility", "Sharpe", "Final value",
        "Trades", "Worst year",
    ]
    return display_frame

def refresh_results_history(_=None):
    try:
        refresh_results_history_controls()
        frame = _history_query()
        options = []
        for row in frame.itertuples(index=False):
            cagr_text = "—" if pd.isna(row.cagr) else f"{row.cagr * 100:.2f}%"
            label = (
                f"#{row.id} | {row.strategy} | {row.primary_asset or '—'} | "
                f"{row.period_start} → {row.period_end} | CAGR {cagr_text}"
            )
            options.append((label, row.id))
        history_selection.options = options
        history_selection.value = ()
        with history_output:
            clear_output()
            if frame.empty:
                print("No saved backtests match the current filters.")
            else:
                display(_display_history_dataframe(frame))
        history_status.value = f"<b>{len(frame)}</b> saved backtest(s) shown."
        return frame
    except Exception as exc:
        history_status.value = f"<span style='color:#b00020'>History error: {exc}</span>"
        return pd.DataFrame()

def _clear_history_filters(_=None):
    history_strategy_dd.value = "All strategies"
    history_asset_box.value = ""
    history_parameter_box.value = ""
    history_start_picker.value = None
    history_end_picker.value = None
    for widget in [
        history_cagr_box, history_drawdown_box, history_sharpe_box,
        history_volatility_box, history_trades_box, history_final_value_box,
    ]:
        widget.value = ""
    history_sort_dd.value = "created_at"
    history_ascending_cb.value = False
    refresh_results_history()

def _export_history(frame):
    if frame.empty:
        raise ValueError("There are no saved results to export.")
    export_dir = tempfile.mkdtemp(prefix="backtest_history_")
    path = os.path.join(export_dir, "backtest_results_history.csv")
    frame.to_csv(path, index=False)
    if IN_COLAB:
        files.download(path)
    else:
        display(FileLink(path))
    return path

def _export_filtered_history(_=None):
    with history_output:
        try:
            _export_history(_history_query())
        except Exception as exc:
            print("ERROR:", exc)

def _export_all_history(_=None):
    with history_output:
        try:
            _export_history(load_saved_backtests())
        except Exception as exc:
            print("ERROR:", exc)

def _delete_selected_history(_=None):
    selected = list(history_selection.value)
    if not selected:
        history_status.value = "Select one or more saved results to delete."
        return
    deleted = delete_saved_backtests(selected)
    history_status.value = f"Deleted {deleted} saved backtest(s)."
    refresh_results_history()

history_refresh_btn.on_click(refresh_results_history)
history_clear_btn.on_click(_clear_history_filters)
history_export_filtered_btn.on_click(_export_filtered_history)
history_export_all_btn.on_click(_export_all_history)
history_delete_btn.on_click(_delete_selected_history)

history_box = widgets.VBox([
    widgets.HTML(
        "<h3>Results History</h3><p>Use multiple filters together. Percentage fields accept "
        "expressions such as <code>&gt; 10%</code> or <code>&gt; -30%</code>.</p>"
    ),
    widgets.HBox([history_strategy_dd, history_asset_box, history_parameter_box]),
    widgets.HBox([history_start_picker, history_end_picker, history_sort_dd, history_ascending_cb]),
    widgets.HBox([history_cagr_box, history_drawdown_box, history_sharpe_box]),
    widgets.HBox([history_volatility_box, history_trades_box, history_final_value_box]),
    widgets.HBox([
        history_refresh_btn, history_clear_btn, history_export_filtered_btn,
        history_export_all_btn, history_delete_btn,
    ]),
    history_status,
    history_selection,
    history_output,
])

online_ticker_box = widgets.Text(value="", description="Ticker:", placeholder="QQQ")
online_currency_dd = widgets.Dropdown(options=["USD", "ILS"], value="USD", description="Currency:")
online_start_picker = widgets.DatePicker(description="From:")
online_end_picker = widgets.DatePicker(description="To:")
online_download_btn = widgets.Button(description="Download online data", button_style="info")
online_output = widgets.Output()

def _download_online_clicked(_):
    with online_output:
        clear_output()
        try:
            name, series = download_yahoo_asset(
                ticker=online_ticker_box.value,
                start_date=online_start_picker.value,
                end_date=online_end_picker.value,
                currency=online_currency_dd.value,
            )
            refresh_strategy_dropdowns()
            refresh_fx_dropdown()
            print(
                f"Loaded {name}: {len(series):,} adjusted daily prices, "
                f"{series.index.min().date()} → {series.index.max().date()}, "
                f"currency {online_currency_dd.value}."
            )
            print("Yahoo Adjusted Close includes split and dividend adjustments when Yahoo supplies them.")
        except Exception as e:
            print("ERROR:", e)

online_download_btn.on_click(_download_online_clicked)

def _agitq_asset_options():
    defaults = ["QQQ", "TQQQ", "QLD", "SGOV", "SPYM"]
    return list(dict.fromkeys(defaults + list(ASSETS.keys())))

def refresh_agitq_asset_dropdowns():
    options = _agitq_asset_options()
    for dropdown, allow_cash in [
        (agitq_signal_asset_dd, False),
        (agitq_traded_asset_dd, False),
        (agitq_defensive_asset_dd, True),
        (agitq_parking_asset_dd, False),
    ]:
        selected = dropdown.value
        dropdown.options = (["CASH"] if allow_cash else []) + options
        if selected in dropdown.options:
            dropdown.value = selected

def _update_agitq_variant_visibility():
    hybrid = agitq_variant_dd.value == "Hybrid 3/185/161"
    original = agitq_variant_dd.value == "Original 200"
    agitq_short_sma_box.layout.display = "none" if original else ""
    agitq_long_sma_box.layout.display = "none" if hybrid else ""
    agitq_hybrid_row.layout.display = "" if hybrid else "none"

def _apply_agitq_preset(_=None):
    preset = AGITQ_PRESETS[agitq_variant_dd.value]
    refresh_agitq_asset_dropdowns()
    for dropdown, value in [
        (agitq_signal_asset_dd, preset["signal_asset"]),
        (agitq_traded_asset_dd, preset["traded_asset"]),
        (agitq_defensive_asset_dd, "SGOV"),
        (agitq_parking_asset_dd, "SPYM"),
    ]:
        if value in dropdown.options:
            dropdown.value = value
    if preset["short_sma"] is not None:
        agitq_short_sma_box.value = preset["short_sma"]
    if preset["long_sma"] is not None:
        agitq_long_sma_box.value = preset["long_sma"]
    if preset["entry_sma"] is not None:
        agitq_entry_sma_box.value = preset["entry_sma"]
    if preset["exit_sma"] is not None:
        agitq_exit_sma_box.value = preset["exit_sma"]
    agitq_envelope_box.value = 5.0
    agitq_envelope_enabled_cb.value = True
    agitq_confirmation_dd.value = 0
    agitq_profit_taking_cb.value = False
    _update_agitq_variant_visibility()

def _strategy_changed(change=None):
    is_agitq = strategy_dd.value == AGITQ_STRATEGY_NAME
    is_haa = strategy_dd.value in (HAA_SIMPLE_STRATEGY_NAME, HAA_BALANCED_STRATEGY_NAME)
    agitq_controls_box.layout.display = "" if is_agitq else "none"
    market_filter_controls_box.layout.display = "none" if is_agitq else ""
    for widget in [
        primary_dd, secondary_dd, freq_dd, execution_dd,
        lookback_slider, sma_slider, envelope_slider,
    ]:
        widget.layout.display = "none" if (is_agitq or (is_haa and widget is not execution_dd)) else ""
    if is_haa:
        market_filter_controls_box.layout.display = "none"

agitq_variant_dd.observe(_apply_agitq_preset, names="value")
strategy_dd.observe(_strategy_changed, names="value")
_apply_agitq_preset()
_strategy_changed()

def refresh_strategy_dropdowns():
    selected_filter = market_filter_asset_dd.value
    opts = list(ASSETS.keys())
    primary_dd.options = opts if opts else [""]
    secondary_dd.options = ["None"] + opts
    benchmark_dd.options = ["None"] + opts
    market_filter_asset_dd.options = ["None"] + opts
    market_filter_asset_dd.value = selected_filter if selected_filter in opts else "None"
    for widget, options in [
        (batch_primary_sm, opts),
        (batch_secondary_sm, ["None"] + opts),
        (batch_benchmark_sm, ["None"] + opts),
        (batch_filter_asset_sm, ["None"] + opts),
    ]:
        selected = tuple(value for value in widget.value if value in options)
        widget.options = options
        widget.value = selected or (("None",) if "None" in options else tuple())
    if opts:
        primary_dd.value = opts[0]
    refresh_agitq_asset_dropdowns()

def _batch_selections_from_controls():
    return {
        "strategies": list(batch_strategies_sm.value),
        "primary_assets": list(batch_primary_sm.value),
        "secondary_assets": list(batch_secondary_sm.value),
        "base_currencies": list(batch_base_sm.value),
        "frequencies": list(batch_frequency_sm.value),
        "lookbacks": parse_batch_values(batch_lookbacks_box.value, int, "ROC months"),
        "sma_days_values": parse_batch_values(batch_sma_days_box.value, int, "SMA days"),
        "envelope_pcts": [value / 100 for value in parse_batch_values(batch_envelopes_box.value, float, "envelope percentage")],
        "tax_rates": [value / 100 for value in parse_batch_values(batch_tax_box.value, float, "tax percentage")],
        "transaction_costs": [value / 100 for value in parse_batch_values(batch_fee_box.value, float, "transaction-cost percentage")],
        "cash_rates": [value / 100 for value in parse_batch_values(batch_cash_box.value, float, "cash-return percentage")],
        "initial_values": parse_batch_values(batch_initial_box.value, float, "starting value"),
        "execution_modes": list(batch_execution_sm.value),
        "benchmark_assets": list(batch_benchmark_sm.value),
        "filter_assets": list(batch_filter_asset_sm.value),
        "filter_frequencies": list(batch_filter_frequency_sm.value),
    }

def _refresh_batch_count(_=None):
    try:
        selections = _batch_selections_from_controls()
        count = batch_combination_count(**selections)
        batch_count_html.value = f"<b>{count:,}</b> combination(s) will be tested."
    except Exception as exc:
        batch_count_html.value = f"<span style='color:#b00020'>Fix Batch Test values: {exc}</span>"

def _display_batch_results():
    if batch_results.empty:
        print("No Batch Test results yet.")
        return
    filtered = filter_and_sort_batch_results(
        batch_results,
        sort_by=batch_sort_dd.value,
        ascending=batch_ascending_cb.value,
        metric=batch_metric_filter_dd.value,
        expression=batch_metric_filter_box.value,
        text=batch_text_filter_box.value,
    )
    print(f"Showing {len(filtered):,} of {len(batch_results):,} completed combination(s).")
    highlights = batch_result_highlights(filtered)
    if not highlights.empty:
        print("\nHighlights")
        display(highlights)
    display_frame = filtered.copy()
    for column in ["CAGR", "Max drawdown", "Volatility", "Worst year"]:
        display_frame[column] = display_frame[column].map(lambda value: "—" if pd.isna(value) else f"{value * 100:.2f}%")
    display_frame["Sharpe"] = display_frame["Sharpe"].map(lambda value: "—" if pd.isna(value) else f"{value:.2f}")
    display_frame["Final value"] = display_frame["Final value"].map(lambda value: "—" if pd.isna(value) else money(value))
    display(display_frame)

def _run_batch_clicked(_=None):
    global batch_results
    with batch_output:
        clear_output()
        try:
            selections = _batch_selections_from_controls()
            count = batch_combination_count(**selections)
            if count == 0:
                raise ValueError("Select at least one value for every Batch Test control.")
            print(f"Running {count:,} combination(s) with the standard individual-backtest engine...")
            batch_results = run_batch_test(
                **selections,
                start_date=batch_start_picker.value,
                end_date=batch_end_picker.value,
            )
            _display_batch_results()
        except Exception as exc:
            print("ERROR:", exc)

def _apply_batch_filters(_=None):
    with batch_output:
        clear_output()
        try:
            _display_batch_results()
        except Exception as exc:
            print("ERROR:", exc)

for control in [
    batch_strategies_sm, batch_primary_sm, batch_secondary_sm, batch_benchmark_sm,
    batch_filter_asset_sm, batch_base_sm, batch_frequency_sm, batch_execution_sm,
    batch_filter_frequency_sm, batch_lookbacks_box, batch_sma_days_box,
    batch_envelopes_box, batch_tax_box, batch_fee_box, batch_cash_box, batch_initial_box,
]:
    control.observe(_refresh_batch_count, names="value")
batch_run_btn.on_click(_run_batch_clicked)
batch_apply_filters_btn.on_click(_apply_batch_filters)
_refresh_batch_count()

def _run_clicked(_):
    with result_output:
        clear_output()
        try:
            if strategy_dd.value in (HAA_SIMPLE_STRATEGY_NAME, HAA_BALANCED_STRATEGY_NAME):
                result = run_haa_test(
                    strategy=strategy_dd.value, base_currency=base_dd.value,
                    initial_value=initial_box.value, tax_rate=tax_box.value / 100, tx_cost=fee_box.value / 100,
                    execution_mode=execution_dd.value, start_date=start_date_picker.value,
                    end_date=end_date_picker.value,
                )
                show_result(result)
                print("\nCanonical HAA monthly decision audit:")
                display(result["haa_audit"])
                return
            if strategy_dd.value == AGITQ_STRATEGY_NAME:
                variant = agitq_variant_dd.value
                hybrid = variant == "Hybrid 3/185/161"
                original = variant == "Original 200"
                result = run_agitq_test(
                    variant=variant,
                    signal_asset=agitq_signal_asset_dd.value,
                    traded_asset=agitq_traded_asset_dd.value,
                    defensive_asset=agitq_defensive_asset_dd.value,
                    parking_asset=agitq_parking_asset_dd.value,
                    short_sma_days=None if original else agitq_short_sma_box.value,
                    long_sma_days=None if hybrid else agitq_long_sma_box.value,
                    hybrid_entry_sma_days=agitq_entry_sma_box.value if hybrid else None,
                    hybrid_exit_sma_days=agitq_exit_sma_box.value if hybrid else None,
                    envelope_pct=agitq_envelope_box.value / 100,
                    envelope_enabled=agitq_envelope_enabled_cb.value,
                    confirmation_days=agitq_confirmation_dd.value,
                    partial_profit_taking=agitq_profit_taking_cb.value,
                    contribution_amount=agitq_contribution_box.value,
                    contribution_frequency=agitq_contribution_frequency_dd.value,
                    base_currency=base_dd.value,
                    tax_rate=tax_box.value / 100,
                    tx_cost=fee_box.value / 100,
                    cash_rate=cash_box.value / 100,
                    benchmark_asset=benchmark_dd.value,
                    initial_value=initial_box.value,
                    start_date=start_date_picker.value,
                    end_date=end_date_picker.value,
                )
                saved_id = save_backtest_result(result)
                refresh_results_history()
                print(f"Saved backtest #{saved_id} to Results History.")
                show_result(result, benchmark_dd.value)
                return

            if not primary_dd.value:
                print("Upload and configure at least one dataset first.")
                return

            result = run_test(
                strategy=strategy_dd.value,
                primary=primary_dd.value,
                secondary=secondary_dd.value,
                base_currency=base_dd.value,
                frequency=freq_dd.value,
                lookback_months=lookback_slider.value,
                sma_days=sma_slider.value,
                envelope_pct=envelope_slider.value / 100,
                tax_rate=tax_box.value / 100,
                tx_cost=fee_box.value / 100,
                cash_rate=cash_box.value / 100,
                execution_mode=execution_dd.value,
                benchmark_asset=benchmark_dd.value,
                initial_value=initial_box.value,
                start_date=start_date_picker.value,
                end_date=end_date_picker.value,
                market_filter_asset=market_filter_asset_dd.value,
                market_filter_rule=market_filter_rule_dd.value,
                market_filter_evaluation_frequency=market_filter_frequency_dd.value,
            )
            saved_id = save_backtest_result(result)
            refresh_results_history()
            print(f"Saved backtest #{saved_id} to Results History.")
            show_result(result, benchmark_dd.value)
        except Exception as e:
            print("ERROR:", e)

run_btn.on_click(_run_clicked)

# ---------- 19. FX UI ----------
fx_asset_dd = widgets.Dropdown(options=["None"], description="FX dataset:")
fx_orientation_dd = widgets.Dropdown(
    options=["ILS per USD", "USD per ILS"],
    value="ILS per USD",
    description="FX format:"
)
fx_btn = widgets.Button(description="Set FX series")
fx_out = widgets.Output()

def refresh_fx_dropdown():
    selected = fx_asset_dd.value
    options = ["None"] + list(ASSETS.keys())
    fx_asset_dd.options = options
    fx_asset_dd.value = selected if selected in options else "None"

def _set_fx(_):
    with fx_out:
        clear_output()
        if fx_asset_dd.value == "None":
            print("Choose the USD/ILS dataset.")
            return
        try:
            set_fx_from_asset(fx_asset_dd.value, fx_orientation_dd.value)
        except Exception as e:
            print("ERROR:", e)

fx_btn.on_click(_set_fx)

# ---------- 20. APP ----------
def launch_backtester():
    refresh_strategy_dropdowns()
    refresh_fx_dropdown()
    refresh_results_history_controls()

    upload_btn = widgets.Button(description="Upload CSV / Excel", button_style="info", layout=widgets.Layout(width="190px"))
    upload_btn.on_click(lambda _: upload_csvs())

    backtest_box = widgets.VBox([
        widgets.HTML("<h2>Simple ETF Backtester</h2>"),
        widgets.HTML(
            "<b>Purpose:</b> approximate, consistent comparison of ETF strategies. "
            "Not a brokerage-grade tax simulator."
        ),
        widgets.HTML("<h3>Download ETF or stock prices online</h3>"),
        widgets.HBox([online_ticker_box, online_currency_dd, online_download_btn]),
        widgets.HBox([online_start_picker, online_end_picker]),
        widgets.HTML(
            "Enter a Yahoo Finance ticker. Leave the dates blank for all available history. "
            "Downloaded prices use Adjusted Close when available."
        ),
        online_output,
        widgets.HTML("<h3>Or upload CSV / Excel files</h3>"),
        upload_btn,
        import_output,
        widgets.HTML("<h3>Optional FX conversion</h3>"),
        widgets.HBox([fx_asset_dd, fx_orientation_dd, fx_btn]),
        fx_out,
        widgets.HTML("<h3>Backtest</h3>"),
        widgets.HBox([strategy_dd, primary_dd, secondary_dd]),
        widgets.HBox([benchmark_dd, base_dd, freq_dd, execution_dd]),
        market_filter_controls_box,
        agitq_controls_box,
        widgets.HBox([start_date_picker, end_date_picker]),
        widgets.HTML("Leave dates blank for the full available history. Earlier data is used for indicators; "
                     "the latest earlier signal sets the position at the first available close. "
                     "Without an earlier signal, the strategy stays in cash until one can execute."),
        lookback_slider,
        sma_slider,
        envelope_slider,
        widgets.HBox([tax_box, fee_box, cash_box, initial_box]),
        widgets.HTML("Cash return is an annual rate applied while the strategy holds CASH; it compounds between trading dates."),
        run_btn,
        result_output
    ])
    batch_box = widgets.VBox([
        widgets.HTML("<h2>Batch Test</h2>"),
        widgets.HTML(
            "Choose multiple values to run their complete Cartesian product with the same engine used by an individual backtest. "
            "AGITQ remains an individual-only test because it has a separate allocation engine."
        ),
        widgets.HBox([batch_strategies_sm, batch_primary_sm, batch_secondary_sm]),
        widgets.HBox([batch_benchmark_sm, batch_filter_asset_sm]),
        widgets.HBox([batch_base_sm, batch_frequency_sm, batch_execution_sm, batch_filter_frequency_sm]),
        widgets.HBox([batch_lookbacks_box, batch_sma_days_box, batch_envelopes_box]),
        widgets.HBox([batch_tax_box, batch_fee_box, batch_cash_box, batch_initial_box]),
        widgets.HBox([batch_start_picker, batch_end_picker]),
        batch_count_html,
        batch_run_btn,
        widgets.HTML("<h3>Results</h3><p>Sort by any displayed column. Filter a metric with expressions such as <code>&gt; 10%</code> or <code>&lt; 20</code>; text filtering searches all columns.</p>"),
        widgets.HBox([batch_sort_dd, batch_ascending_cb, batch_metric_filter_dd, batch_metric_filter_box]),
        widgets.HBox([batch_text_filter_box, batch_apply_filters_btn]),
        batch_output,
    ])
    app_tabs = widgets.Tab(children=[backtest_box, batch_box, history_box])
    app_tabs.set_title(0, "Backtest")
    app_tabs.set_title(1, "Batch Test")
    app_tabs.set_title(2, "Results History")
    app_tabs.selected_index = 0
    display(app_tabs)
    refresh_results_history()

launch_backtester()
