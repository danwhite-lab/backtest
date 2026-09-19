# Simple ETF Backtester

An interactive Google Colab backtester for ETF and stock strategies.

[Open the backtester in Google Colab](https://colab.research.google.com/github/danwhite-lab/backtest/blob/main/Simple_ETF_Backtester.ipynb)

## What it supports

- Download adjusted historical ETF or stock prices by Yahoo Finance ticker.
- Upload CSV, XLSX, or XLS price files as an alternative.
- Automatically detect common date formats and price columns.
- Compare buy-and-hold, absolute momentum, dual momentum, relative momentum, SMA trend, and SMA Hysteresis Envelope strategies.
- Apply an optional external `Price > SMA` market filter using any loaded dataset, including an asset outside the tradable universe.
- Convert between USD and ILS using an uploaded FX series.
- Set start and end dates, signal frequency, execution timing, tax, transaction cost, and annual cash return.
- Compare complete strategy and benchmark metrics in one table.
- Count every executed buy and sell as a trade.

## Use in Colab

1. Open the notebook using the link above.
2. Run its code cell.
3. Enter a ticker such as `QQQ` or `SPY`, select its currency, and click **Download online data**.
4. Choose the strategy, primary asset, optional second asset, and benchmark.
5. Optionally choose a **Market filter** asset. Its price must be above the selected **SMA days** value or the strategy holds cash.
6. Enter the remaining backtest assumptions and click **Run backtest**.

Leave the online date fields blank to request all available history. Online downloads use Yahoo Finance `Adjusted Close` when it is available, so splits and distributions are reflected in the price series supplied by Yahoo.

## Files

- `Simple_ETF_Backtester.ipynb` — interactive Colab notebook.
- `simple_etf_backtester_colab.py` — standalone Python version.

Results are approximate and depend on the selected data and assumptions. The capital-gains tax calculation is deliberately simplified and is not a brokerage or tax-accounting simulator.
