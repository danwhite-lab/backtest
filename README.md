# Simple ETF Backtester

An interactive Google Colab backtester for ETF and stock strategies.

[Open the backtester in Google Colab](https://colab.research.google.com/github/danwhite-lab/backtest/blob/main/Simple_ETF_Backtester.ipynb)

## What it supports

- Download adjusted historical ETF or stock prices by Yahoo Finance ticker.
- Upload CSV, XLSX, or XLS price files as an alternative.
- Automatically detect common date formats and price columns.
- Compare buy-and-hold, absolute momentum, dual momentum, relative momentum, SMA trend, and SMA Hysteresis Envelope strategies.
- Run Original 200, QQQ 3/161, QLD ISA, TQQQ 5/218, and Hybrid 3/185/161 AGITQ playbook variants.
- Configure AGITQ signal/traded/defensive/parking assets, confirmation days, SMA periods, and the optional overheat envelope.
- Route AGITQ periodic contributions to the leveraged or S&P parking asset according to the daily temperature state.
- Remove external contributions from time-weighted CAGR, volatility, Sharpe, drawdown, and worst-year calculations.
- Report contribution totals, net profit, and money-weighted return (XIRR) separately.
- Apply the same contribution schedule to a selected buy-and-hold benchmark.
- Inspect and export AGITQ settings, trades, events, performance, and full daily signal state.
- Apply an optional external `Price > SMA` market filter using any loaded dataset, including an asset outside the tradable universe.
- Evaluate the external market filter daily, weekly, or monthly without changing the strategy's own signal frequency.
- Automatically save individual and batch backtests to a local SQLite history database, including full settings and performance metrics.
- Filter, sort, delete, compare, and export saved results from the Results History panel.
- Convert between USD and ILS using an uploaded FX series.
- Set start and end dates, signal frequency, execution timing, tax, transaction cost, and annual cash return.
- Compare complete strategy and benchmark metrics in one table.
- Count every executed buy and sell as a trade.

## Use in Colab

1. Open the notebook using the link above.
2. Run its code cell.
3. Enter a ticker such as `QQQ` or `SPY`, select its currency, and click **Download online data**.
4. Choose the strategy, primary asset, optional second asset, and benchmark.
5. Optionally choose a **Market filter** asset and **Filter evaluation frequency**. Its price must be above the selected **SMA days** value or the strategy holds cash.
6. Enter the remaining backtest assumptions and click **Run backtest**.

For AGITQ, first load each required adjusted-price series, normally `QQQ`, `TQQQ` or `QLD`, `SGOV`, and `SPYM`. Select **AGITQ / TQQQ Playbook**, choose a preset, and adjust the relevant controls if needed. AGITQ signals use daily observations and execute at the next available close.

AGITQ partial profit-taking is shown as unavailable and remains disabled. The current accounting engine does not track the tax lots and milestone bases needed to implement that feature exactly.

When periodic contributions are enabled, the results table labels the strategy's annual return as **Annual return (TWR)**. This measures the strategy independently of deposits. **Money-weighted return (XIRR)** measures the investor's return using the actual date of each deposit. Drawdown, volatility, Sharpe, and worst year also use the cash-flow-adjusted return series. A deposit into cash is recorded as a contribution but is not counted as a trade; a contribution that executes an asset purchase is counted as a buy.

## Results history

Every successful interactive test is saved automatically, and each completed combination in `batch_momentum_test` is saved separately. The database is `simple_etf_backtester_results.sqlite` in the current working directory. The Results History panel supports combined strategy, asset, setting, period, and numeric filters such as `> 10%`, `> -30%`, and `> 0.6`, plus sorting, deletion of selected rows, and CSV export. To store the database elsewhere, set the `ETF_BACKTESTER_RESULTS_DB` environment variable before running the notebook. In Colab, point it at an already-mounted Google Drive location if you need results to survive a runtime reset.

Leave the online date fields blank to request all available history. Online downloads use Yahoo Finance `Adjusted Close` when it is available, so splits and distributions are reflected in the price series supplied by Yahoo.

## Files

- `Simple_ETF_Backtester.ipynb` — interactive Colab notebook.
- `simple_etf_backtester_colab.py` — standalone Python version.

Results are approximate and depend on the selected data and assumptions. The capital-gains tax calculation is deliberately simplified and is not a brokerage or tax-accounting simulator.
