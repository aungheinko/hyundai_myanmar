import requests
import pandas as pd

# Function to get the order book for a given symbol
def get_order_book(symbol):
    url = f"https://api.binance.com/api/v3/depth"
    params = {"symbol": symbol, "limit": 10}
    response = requests.get(url, params=params)
    
    if response.status_code == 200:
        order_book = response.json()
        return order_book['bids'], order_book['asks']
    else:
        print("Error fetching order book data.")
        return None, None

# Function to display order book data in a human-readable format
def display_order_book(bids, asks, total_buy_volume, total_sell_volume):
    print("\nTop 5 Buy Orders (Bids):")
    print(f"{'Price':<15}{'Quantity':<15}")
    for bid in bids:
        price, quantity = bid
        print(f"{price:<15}{quantity:<15}")
    print(f"Total Buy Volume: {total_buy_volume}")
    
    print("\nTop 5 Sell Orders (Asks):")
    print(f"{'Price':<15}{'Quantity':<15}")
    for ask in asks:
        price, quantity = ask
        print(f"{price:<15}{quantity:<15}")
    print(f"Total Sell Volume: {total_sell_volume}")

# Function to fetch kline (candlestick) data for a given symbol and interval
def get_klines(symbol, interval):
    url = "https://api.binance.com/api/v3/klines"
    params = {"symbol": symbol, "interval": interval, "limit": 1000}  # Fetch 1000 bars
    response = requests.get(url, params=params)
    
    if response.status_code == 200:
        data = response.json()
        df = pd.DataFrame(data, columns=[
            "Open time", "Open", "High", "Low", "Close", 
            "Volume", "Close time", "Quote asset volume", 
            "Number of trades", "Taker buy base volume", 
            "Taker buy quote volume", "Ignore"
        ])
        
        df["Open time"] = pd.to_datetime(df["Open time"], unit='ms')
        df["Close time"] = pd.to_datetime(df["Close time"], unit='ms')
        df["Buy volume"] = df["Taker buy base volume"].astype(float)  # Buy orders
        df["Sell volume"] = (df["Volume"].astype(float) - df["Buy volume"])  # Sell orders
        
        return df
    else:
        print("Error fetching kline data.")
        return None

# Main function to ask user for input and display market insights
def main():
    symbol = input("Enter the trading pair (e.g., BTCUSDT, TONUSDT, LINKUSDT): ").upper()
    interval = input("Enter the timeframe (e.g., 1m, 5m, 15m, 1h): ").lower()
    
    # Fetch order book
    bids, asks = get_order_book(symbol)
    
    # Fetch kline data for analysis
    df = get_klines(symbol, interval)
    
    if df is not None and bids and asks:
        total_buy_volume = df["Buy volume"].sum()
        total_sell_volume = df["Sell volume"].sum()

        display_order_book(bids, asks, total_buy_volume, total_sell_volume)
        
        if total_buy_volume > total_sell_volume:
            print("\nBuyers are dominant in the market today.")
        else:
            print("\nSellers are dominant in the market today.")
    else:
        print("No market data available to analyze.")

# Run the main function
if __name__ == "__main__":
    main()
