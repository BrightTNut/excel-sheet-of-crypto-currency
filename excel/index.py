import requests
import pandas as pd
import time
import os

# Step 1: Fetch Live Data from CoinGecko API
def fetch_crypto_data():
    url = 'https://api.coingecko.com/api/v3/coins/markets'
    params = {
        'vs_currency': 'usd',
        'order': 'market_cap_desc',
        'per_page': 50,
        'page': 1
    }
    response = requests.get(url, params=params)
    data = response.json()

    # Structuring the data into a DataFrame
    crypto_list = []
    for coin in data:
        crypto_list.append({
            'Name': coin['name'],
            'Symbol': coin['symbol'],
            'Current Price (USD)': coin['current_price'],
            'Market Capitalization (USD)': coin['market_cap'],
            '24h Trading Volume (USD)': coin['total_volume'],
            'Price Change (24h %)': coin['price_change_percentage_24h']
        })

    return pd.DataFrame(crypto_list)

# Step 2: Write Data to Excel
def write_to_excel(df, filename='crypto_data.xlsx'):
    # Use a unique temporary filename to avoid conflicts
    temp_filename = "temp_crypto_data_" + str(int(time.time())) + ".xlsx"
    max_retries = 5
    retries = 0

    while retries < max_retries:
        try:
            with pd.ExcelWriter(temp_filename, engine='openpyxl', mode='w') as writer:
                df.to_excel(writer, index=False, sheet_name='Crypto Data')
            # Replace the original file with the updated file
            os.replace(temp_filename, filename)
            print(f"Data written to {filename}")
            break
        except PermissionError as e:
            print(f"Error writing to Excel: {e}")
            print(f"Retrying in 10 seconds... (Attempt {retries + 1}/{max_retries})")
            time.sleep(10)
            retries += 1
        except Exception as e:
            print(f"Unexpected error: {e}")
            break

# Step 3: Set up continuous update every 5 minutes
def update_excel_every_5_minutes():
    while True:
        # Fetch the latest data
        df = fetch_crypto_data()
        # Write the data to Excel
        write_to_excel(df)
        print("Updated Excel with live data")
        # Sleep for 5 minutes (300 seconds)
        time.sleep(300)

# Run the continuous update
if __name__ == "__main__":
    update_excel_every_5_minutes()
