# Cash-and-Carry Arbitrage in Soybean Futures (CBOT)

Python implementation used in the research paper *"Testing Cash-and-Carry Arbitrage in Soybeans Futures"*

This repository contains the full data pipeline used to test the classical cost-of-carry model against real market data from Interactive Brokers (IBKR), estimate implied net carry via a calendar-spread method, and evaluate cash-and-carry arbitrage signals on CBOT soybean futures.

## Repository contents

| File | Purpose |
|---|---|
| `cash_carry_research.py` | Live monitoring daemon — connects to IBKR and logs real-time observations continuously (CSV output). |
| `cash_carry_backtester.py` | Historical batch processor — downloads 6 months of daily BID_ASK data and exports a consolidated two-sheet Excel file (raw observations + summary statistics). Used to generate the results reported in the paper. |
| `generate_figures.py` | Reads the Excel output of the backtester and produces the three figures used in the paper (price comparison, implied carry over time, signal distribution). Does not connect to IBKR. |

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

The port is set in the `port` parameter of `BacktestConfig` / `ArbitrageResearchOrchestrator`, depending on the script.

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

### 3. Live monitoring (optional, not used for the paper's dataset)

```bash
python cash_carry_research.py
```

Runs indefinitely, logging one observation per polling interval to a CSV file. Stop with `Ctrl+C`.

## Data notes

- Market data is retrieved via `whatToShow="BID_ASK"` on IBKR's `reqHistoricalData`, preserving real bid/ask microstructure. If unavailable for a given contract, the backtester falls back to `TRADES` bars with a synthetic half-spread (documented per-row in the `spread_method_*` columns of the output).
- CBOT grain futures (ZS, ZC, ZW) are quoted in **cents per bushel**; the backtester converts to USD/bushel via the `price_scale` parameter (default 100).
- The risk-free rate is obtained from the U.S. Treasury's daily 6-month yield curve, matched to each observation date with no look-ahead bias.

## Market hours

CBOT soybean futures are closed on weekends. Running the scripts outside market hours will not raise an error, but observations will be logged as `INCOMPLETE_DATA` (live daemon) due to an empty order book.

## Citation

If referencing this code, please cite the accompanying paper. See the paper's References section for the theoretical sources (Hull; Kaldor, 1939; Working, 1949; Brennan, 1958) underlying the cost-of-carry and convenience yield framework implemented here.
