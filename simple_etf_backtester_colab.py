
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
# - CAGR, max drawdown, volatility, Sharpe, trades, allocation
#
# IMPORTANT:
# This is a ballpark backtester. It prioritizes consistency and transparency.
# Dividends are included only when the chosen price column already includes them
# (for example, Adjusted Close).
# ============================================================

# ---------- 1. INSTALL / IMPORT ----------
import sys, subprocess, pkgutil, io, math, warnings, json, zipfile, tempfile
from dataclasses import dataclass
from datetime import date, datetime
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
    )

# ---------- 13. METRICS ----------
def calculate_metrics(equity: pd.Series, trades: pd.DataFrame, holdings: pd.Series):
    if len(equity) < 2:
        raise ValueError("Not enough equity data.")

    years = (equity.index[-1] - equity.index[0]).days / 365.25
    cagr = (equity.iloc[-1] / equity.iloc[0]) ** (1 / years) - 1 if years > 0 else np.nan

    daily_ret = equity.pct_change().dropna()
    vol = daily_ret.std() * np.sqrt(252)
    sharpe = (daily_ret.mean() / daily_ret.std()) * np.sqrt(252) if daily_ret.std() > 0 else np.nan

    peak = equity.cummax()
    dd = equity / peak - 1
    max_dd = dd.min()

    annual = equity.resample("YE").last().pct_change().dropna()
    worst_year = annual.min() if not annual.empty else np.nan

    # Every executed BUY or SELL is one trade.
    trade_count = int(len(trades)) if not trades.empty else 0

    alloc = holdings.value_counts(normalize=True).sort_values(ascending=False)

    return {
        "CAGR": cagr,
        "Max Drawdown": max_dd,
        "Volatility": vol,
        "Sharpe": sharpe,
        "Final Value": equity.iloc[-1],
        "Trades": trade_count,
        "Worst Year": worst_year,
        "Start": equity.index[0],
        "End": equity.index[-1],
        "Allocation": alloc,
        "DrawdownSeries": dd,
    }

# ---------- 14. BENCHMARK ----------
def benchmark_buyhold(prices: pd.DataFrame, asset: str, initial_value=100000):
    s = prices[asset].dropna()
    return initial_value * (s / s.iloc[0])

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
):
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

    equity, trades, holdings = backtest(
        prices=prices,
        signals=signals,
        initial_value=initial_value,
        tax_rate=tax_rate,
        tx_cost=tx_cost,
        cash_rate=cash_rate,
        execution_mode=execution_mode,
    )

    metrics = calculate_metrics(equity, trades, holdings)

    bench = None
    bench_metrics = None
    if benchmark_asset and benchmark_asset != "None":
        bench = benchmark_buyhold(prices.loc[equity.index.min():equity.index.max()], benchmark_asset, initial_value)
        # Align to equity
        bench = bench.reindex(equity.index).ffill().dropna()
        bench_metrics = calculate_metrics(bench, pd.DataFrame(), pd.Series("BENCH", index=bench.index))
        # A buy-and-hold benchmark makes one opening purchase and no sale.
        bench_metrics["Trades"] = 1

    return {
        "prices": prices,
        "signals": signals,
        "equity": equity,
        "trades": trades,
        "holdings": holdings,
        "metrics": metrics,
        "benchmark": bench,
        "benchmark_metrics": bench_metrics,
        "cash_rate": cash_rate,
        "market_filter_asset": filter_asset,
        "market_filter_rule": market_filter_rule if filter_asset else None,
        "market_filter_sma_days": sma_days if filter_asset else None,
        "market_filter_evaluation_frequency": (
            market_filter_evaluation_frequency if filter_asset else None
        ),
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
    equity, trades, holdings, events, total_contributions = backtest_agitq(
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
    metrics = calculate_metrics(equity, trades, holdings)

    bench = None
    bench_metrics = None
    if benchmark_asset and benchmark_asset != "None":
        bench = benchmark_buyhold(
            prices.loc[equity.index.min():equity.index.max()], benchmark_asset, initial_value
        )
        bench = bench.reindex(equity.index).ffill().dropna()
        bench_metrics = calculate_metrics(bench, pd.DataFrame(), pd.Series("BENCH", index=bench.index))
        bench_metrics["Trades"] = 1

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
        "Execution": "Next available close",
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
    }

# ---------- 16. DISPLAY ----------
def pct(x):
    return "—" if pd.isna(x) else f"{x*100:.1f}%"

def money(x):
    return f"{x:,.0f}"

def show_result(result, benchmark_label=None):
    m = result["metrics"]

    def formatted_metrics(metrics):
        return [
            pct(metrics["CAGR"]),
            pct(metrics["Max Drawdown"]),
            pct(metrics["Volatility"]),
            "—" if pd.isna(metrics["Sharpe"]) else f"{metrics['Sharpe']:.2f}",
            money(metrics["Final Value"]),
            str(metrics["Trades"]),
            pct(metrics["Worst Year"]),
            f"{metrics['Start'].date()} → {metrics['End'].date()}",
        ]

    table = {
        "Metric": [
            "CAGR", "Max drawdown", "Volatility", "Sharpe",
            "Final value", "Trades", "Worst year", "Period",
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
                "Existing performance calculations are retained and therefore include deposited capital."
            )
    elif result.get("market_filter_asset"):
        print(
            f"\nMarket filter: {result['market_filter_asset']} — "
            f"{result['market_filter_rule']} ({result['market_filter_sma_days']}-day SMA). "
            f"Filter evaluation frequency: {result['market_filter_evaluation_frequency']}. "
            "When the rule is false or the SMA is unavailable, the strategy holds CASH."
        )
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
    normalized = result["equity"] / result["equity"].iloc[0] * 100
    plt.plot(normalized.index, normalized.values, label="Strategy")
    if result["benchmark"] is not None:
        b = result["benchmark"]
        b = b / b.iloc[0] * 100
        plt.plot(b.index, b.values, label=benchmark_label or "Benchmark")
    plt.title("Growth of 100")
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
    metrics_export = {
        key: value for key, value in result["metrics"].items()
        if key not in ("Allocation", "DrawdownSeries")
    }
    pd.DataFrame(list(metrics_export.items()), columns=["Metric", "Value"]).to_csv(
        f"{export_dir}/agitq_metrics.csv", index=False
    )
    zip_path = f"{export_dir}/AGITQ_backtest_audit.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for filename in [
            "agitq_settings.csv", "agitq_daily_state.csv", "agitq_trades.csv",
            "agitq_events.csv", "agitq_equity.csv", "agitq_metrics.csv",
        ]:
            archive.write(f"{export_dir}/{filename}", arcname=filename)
    if IN_COLAB:
        files.download(zip_path)
    else:
        display(FileLink(zip_path))
    return zip_path

# ---------- 17. BATCH TEST ----------
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
    agitq_controls_box.layout.display = "" if is_agitq else "none"
    market_filter_controls_box.layout.display = "none" if is_agitq else ""
    for widget in [
        primary_dd, secondary_dd, freq_dd, execution_dd,
        lookback_slider, sma_slider, envelope_slider,
    ]:
        widget.layout.display = "none" if is_agitq else ""

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
    if opts:
        primary_dd.value = opts[0]
    refresh_agitq_asset_dropdowns()

def _run_clicked(_):
    with result_output:
        clear_output()
        try:
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

    upload_btn = widgets.Button(description="Upload CSV / Excel", button_style="info", layout=widgets.Layout(width="190px"))
    upload_btn.on_click(lambda _: upload_csvs())

    display(widgets.HTML("<h2>Simple ETF Backtester</h2>"))
    display(widgets.HTML(
        "<b>Purpose:</b> approximate, consistent comparison of ETF strategies. "
        "Not a brokerage-grade tax simulator."
    ))

    display(widgets.HTML("<h3>Download ETF or stock prices online</h3>"))
    display(widgets.HBox([online_ticker_box, online_currency_dd, online_download_btn]))
    display(widgets.HBox([online_start_picker, online_end_picker]))
    display(widgets.HTML(
        "Enter a Yahoo Finance ticker. Leave the dates blank for all available history. "
        "Downloaded prices use Adjusted Close when available."
    ))
    display(online_output)

    display(widgets.HTML("<h3>Or upload CSV / Excel files</h3>"))
    display(upload_btn)
    display(import_output)

    display(widgets.HTML("<h3>Optional FX conversion</h3>"))
    display(widgets.HBox([fx_asset_dd, fx_orientation_dd, fx_btn]))
    display(fx_out)

    display(widgets.HTML("<h3>Backtest</h3>"))
    display(widgets.VBox([
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
    ]))

launch_backtester()
