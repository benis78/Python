import yfinance as yf

# Hent historiske data for Apple (AAPL)
data = yf.download('AAPL', start='2020-01-01', end='2023-01-01')
print(data.head())