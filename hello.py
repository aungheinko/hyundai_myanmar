import requests
import pandas as pd

# Function to get the order book for the given symbol (Top 5 Bids and Asks)
def get_order_book(symbol):
    url = f"https://api.binance.com/api/v3/depth"
    params = {"symbol": symbol, "limit": 5}
    
    # Make the request and handle any errors
    response = requests.get(url, params=params)
    if response.status_code == 200:
        order_book = response.json()
        return order_book['bids'], order_book['asks']
    else:
        print("Error fetching order book data.")
        return None, None

# Function to display the order book in a simple human-readable format
def display_order_book(bids, asks, total_buy_volume, total_sell_volume):
    print("\nTop 5 Buy Orders (Bids):")
    print(f"{'Price':<15}{'Quantity':<15}")
    for bid in bids:
        price, quantity = bid
        print(f"{float(price):<15.2f}{float(quantity):<15,.0f}")  # Format price to 2 decimals and quantity as whole number
    print(f"Total Buy Volume: {total_buy_volume:,.0f}")  # Format volume as whole number with commas

    print("\nTop 5 Sell Orders (Asks):")
    print(f"{'Price':<15}{'Quantity':<15}")
    for ask in asks:
        price, quantity = ask
        print(f"{float(price):<15.2f}{float(quantity):<15,.0f}")  # Format price to 2 decimals and quantity as whole number
    print(f"Total Sell Volume: {total_sell_volume:,.0f}")  # Format volume as whole number with commas

# Function to get candlestick (kline) data and calculate buy/sell volumes
def get_klines(symbol, interval):
    url = "https://api.binance.com/api/v3/klines"
    params = {"symbol": symbol, "interval": interval, "limit": 1000}  # Fetch last 1000 data points
    
    # Request kline data
    response = requests.get(url, params=params)
    if response.status_code == 200:
        # Parse the response and create a DataFrame
        data = response.json()
        df = pd.DataFrame(data, columns=[
            "Open time", "Open", "High", "Low", "Close", 
            "Volume", "Close time", "Quote asset volume", 
            "Number of trades", "Taker buy base volume", 
            "Taker buy quote volume", "Ignore"
        ])
        
        # Convert timestamps and calculate buy/sell volumes
        df["Open time"] = pd.to_datetime(df["Open time"], unit='ms')
        df["Close time"] = pd.to_datetime(df["Close time"], unit='ms')
        df["Buy volume"] = df["Taker buy base volume"].astype(float)  # Buy orders
        df["Sell volume"] = df["Volume"].astype(float) - df["Buy volume"]  # Sell orders
        
        return df
    else:
        print("Error fetching kline data.")
        return None

# Main function to interact with the user and analyze market data
def main():
    # Ask the user for the trading pair and time interval
    symbol = input("Enter the trading pair (e.g., BTCUSDT, TONUSDT): ").upper()
    interval = input("Enter the timeframe (e.g., 1m, 5m, 15m): ").lower()
    
    # Fetch order book and kline data
    bids, asks = get_order_book(symbol)
    df = get_klines(symbol, interval)
    
    # If data was fetched successfully, display the results
    if df is not None and bids and asks:
        total_buy_volume = df["Buy volume"].sum()
        total_sell_volume = df["Sell volume"].sum()

        # Show the order book and volumes
        display_order_book(bids, asks, total_buy_volume, total_sell_volume)
        
        # Determine whether buyers or sellers are stronger today
        if total_buy_volume > total_sell_volume:
            print("\nBuyers are dominant today.")
        else:
            print("\nSellers are dominant today.")
    else:
        print("No market data available to analyze.")

# Run the main function
if __name__ == "__main__":
    main()
