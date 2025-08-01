import yfinance as yf

# List of scrips (stock symbols) to fetch data for
scrips = ['AAPL', 'GOOGL', 'MSFT', 'TSLA', 'AMZN']

# Fetch and display data for each scrip
for scrip in scrips:
    print(f"Fetching data for {scrip}...")
    ticker = yf.Ticker(scrip)
    data = ticker.history(period="5d")  # Fetch last 5 days of data
    print(data)
    print("-" * 40) 