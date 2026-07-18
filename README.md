# Cash-and-Carry Arbitrage in Soybean Futures (CBOT)

Python implementation used in the research paper *"Testing Cash-and-Carry Arbitrage Soybean Futures"*.

This repository contains the full data pipeline used to test the classical cost-of-carry model against real market data from Interactive Brokers (IBKR), estimate implied net carry via a calendar-spread method, and evaluate cash-and-carry arbitrage signals on CBOT soybean futures.

## Repository contents

| File | Purpose |
|---|---|
| `cash_carry_backtester.py` | Historical batch processor — downloads 6 months of daily BID_ASK data from Interactive Brokers and exports a consolidated two-sheet Excel file (raw observations + summary statistics). This is the script used to generate all results reported in the paper (Section 5). |
| `generate_figures.py` | Reads the Excel output of the backtester and produces the three figures used in the paper (price comparison, implied carry over time, signal distribution). Does not connect to IBKR — figures are always reproducible from the same input file. |

## Requirements

- **Python 3.13** (Python 3.14 is not supported — `ib_insync`'s dependency `eventkit` relies on an `asyncio` function removed in 3.14)
- Interactive Brokers **Trader Workstation (TWS)** or **IB Gateway**, running and logged in
- API access enabled in TWS: *Configuration → API → Settings → "Enable ActiveX and Socket Clients"*

### Python dependencies

```bash
pip install ib_insync nest_asyncio numpy pandas requests yfinance openpyxl matplotlib
```

## IBKR connection ports

| Account type | TWS port | IB Gateway port |
|---|---|---|
| Paper trading | 7497 | 4002 |
| Live account | 7496 | 4001 |

The port is set in the `port` parameter of `BacktestConfig`.

## Usage

### 1. Historical backtest (used to generate the paper's dataset)

```bash
python cash_carry_backtester.py
```

Requires TWS/IB Gateway running and connected. Produces an Excel file (`cash_carry_backtest_<timestamp>.xlsx`) in the same folder as the script, with two sheets:
- **Raw Data** — one row per trading day, with bid/ask prices, implied carry, and arbitrage signal.
- **Summary Stats** — descriptive statistics used in the paper's Results section.

### 2. Generating the figures

```bash
python generate_figures.py path/to/cash_carry_backtest_<timestamp>.xlsx
```

This script only reads the Excel file passed as an argument — it never connects to IBKR, and will always reproduce the same figures from the same input file, regardless of current market conditions.

## Data notes

- Market data is retrieved via `whatToShow="BID_ASK"` on IBKR's `reqHistoricalData`, preserving real bid/ask microstructure. If unavailable for a given contract, the backtester falls back to `TRADES` bars with a synthetic half-spread (documented per-row in the `spread_method_*` columns of the output).
- CBOT grain futures (ZS, ZC, ZW) are quoted in **cents per bushel**; the backtester converts to USD/bushel via the `price_scale` parameter (default 100).
- The risk-free rate is obtained from the U.S. Treasury's daily 6-month yield curve, matched to each observation date with no look-ahead bias.

## Market hours

CBOT soybean futures are closed on weekends. Running `cash_carry_backtester.py` outside market hours does not raise an error for historical data (past sessions are still returned), but running it on a contract-day with no trading activity may return incomplete or missing bars.

## Citation

If referencing this code, please cite the accompanying paper. See the paper's References section for the theoretical sources (Hull; Kaldor, 1939; Working, 1949; Brennan, 1958) underlying the cost-of-carry and convenience yield framework implemented here.
