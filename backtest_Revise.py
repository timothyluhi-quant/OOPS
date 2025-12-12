import pandas as pd 
import numpy as np
import yfinance as yf
import matplotlib.pyplot as plt

# ---------------------------------------------------------
# Load the HQM Strategy produced by your original code
# ---------------------------------------------------------
file_path = "momentum_strategy.xlsx"   # output from your strategy code
strategy = pd.read_excel(file_path)

tickers = strategy["Ticker"].tolist()
shares = strategy["Number of Shares to Buy"].tolist()

print(f"Loaded {len(tickers)} tickers from momentum_strategy.xlsx")

# ---------------------------------------------------------
# Download future price data for backtest
# ---------------------------------------------------------
start_date = pd.Timestamp.today().normalize() - pd.DateOffset(days=1)
start_date = start_date.strftime("%Y-%m-%d")

print("Backtest start date:", start_date)

prices = yf.download(tickers, start=start_date, period="1y", progress=False)

# Flatten MultiIndex if exists
if isinstance(prices.columns, pd.MultiIndex):
    prices.columns = prices.columns.get_level_values(0)

# Keep only close prices
prices = prices["Close"].ffill()

# ---------------------------------------------------------
# Compute Portfolio Value Over Time
# ---------------------------------------------------------
portfolio_value = pd.DataFrame(index=prices.index)
portfolio_value["Portfolio"] = 0

for i, t in enumerate(tickers):
    try:
        portfolio_value["Portfolio"] += prices[t] * shares[i]
    except KeyError:
        print(f"Warning: missing data for {t}")

# Normalize for comparison
portfolio_norm = portfolio_value["Portfolio"] / portfolio_value["Portfolio"].iloc[0]

# ---------------------------------------------------------
# Drawdown calculation & Stop-Loss (20%)
# ---------------------------------------------------------
portfolio_value["Peak"] = portfolio_value["Portfolio"].cummax()
portfolio_value["Drawdown"] = (portfolio_value["Portfolio"] - portfolio_value["Peak"]) / portfolio_value["Peak"]

# Apply Stop-Loss: freeze portfolio once drawdown exceeds -20%
stop_loss = -0.20
portfolio_with_stop = portfolio_value["Portfolio"].copy()
stop_triggered = False

for i in range(len(portfolio_with_stop)):
    if stop_triggered:
        portfolio_with_stop.iloc[i] = portfolio_with_stop.iloc[i-1]  # freeze equity after stop-loss
    elif portfolio_value["Drawdown"].iloc[i] <= stop_loss:
        stop_triggered = True
        portfolio_with_stop.iloc[i] = portfolio_with_stop.iloc[i]  # keep value at stop day

portfolio_value["Portfolio_StopLoss"] = portfolio_with_stop
portfolio_stop_norm = portfolio_value["Portfolio_StopLoss"] / portfolio_value["Portfolio_StopLoss"].iloc[0]

# ---------------------------------------------------------
# Benchmark: SPY
# ---------------------------------------------------------
spy = yf.download("SPY", start=portfolio_value.index[0], end=portfolio_value.index[-1], progress=False)["Close"]
spy_norm = spy / spy.iloc[0]

# ---------------------------------------------------------
# Calculate Performance Statistics
# ---------------------------------------------------------
def calculate_stats(portfolio_series):
    returns = portfolio_series.pct_change().dropna()
    days = (portfolio_series.index[-1] - portfolio_series.index[0]).days
    years = days / 365
    cagr = (portfolio_series.iloc[-1] / portfolio_series.iloc[0]) ** (1 / years) - 1
    vol = returns.std() * np.sqrt(252)
    sharpe = cagr / vol
    return cagr, vol, sharpe

cagr_orig, vol_orig, sharpe_orig = calculate_stats(portfolio_value["Portfolio"])
cagr_stop, vol_stop, sharpe_stop = calculate_stats(portfolio_value["Portfolio_StopLoss"])

print("\n---------------- Performance ----------------")
print("Original Strategy:")
print(f"CAGR: {cagr_orig:.2%}, Annual Volatility: {vol_orig:.2%}, Sharpe Ratio: {sharpe_orig:.2f}")
print("\nWith 20% Stop-Loss:")
print(f"CAGR: {cagr_stop:.2%}, Annual Volatility: {vol_stop:.2%}, Sharpe Ratio: {sharpe_stop:.2f}")
print("------------------------------------------------\n")

# ---------------------------------------------------------
# Plot: Portfolio vs SPY
# ---------------------------------------------------------
plt.figure(figsize=(12,6))
plt.plot(portfolio_norm, label="HQM Strategy")
plt.plot(portfolio_stop_norm, label="HQM Strategy + 20% Stop-Loss", linestyle="--")
plt.plot(spy_norm, label="SPY Benchmark", alpha=0.8)
plt.title("HQM Strategy vs SPY Benchmark")
plt.ylabel("Normalized Return")
plt.legend()
plt.grid(True)
plt.show()

# ---------------------------------------------------------
# Plot: Equity Curve
# ---------------------------------------------------------
plt.figure(figsize=(12,6))
plt.plot(portfolio_value["Portfolio"], label="Portfolio Value")
plt.plot(portfolio_value["Portfolio_StopLoss"], label="Portfolio Stop-Loss", linestyle="--")
plt.title("HQM Strategy Equity Curve")
plt.ylabel("Equity ($)")
plt.legend()
plt.grid(True)
plt.show()

# ---------------------------------------------------------
# Plot: Drawdown
# ---------------------------------------------------------
plt.figure(figsize=(12,4))
plt.plot(portfolio_value["Drawdown"], color='red', label="Drawdown")
plt.title("Drawdown (%)")
plt.grid(True)
plt.legend()
plt.show()
