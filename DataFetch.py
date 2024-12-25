import requests
import pandas as pd


API_URL = "https://api.coingecko.com/api/v3/coins/markets"


params = {
    "vs_currency": "usd",       
    "order": "market_cap_desc",
    "per_page": 50,             
    "page": 1,                  
    "sparkline": False          
}

def fetch_crypto_data():
    response = requests.get(API_URL, params=params)
    if response.status_code == 200:
        data = response.json()
    
        df = pd.DataFrame(data, columns=[
            "name", 
            "symbol", 
            "current_price", 
            "market_cap", 
            "total_volume", 
            "price_change_percentage_24h"
        ])
        return df
    else:
        print(f"Failed to fetch data. Status code: {response.status_code}")
        return None

crypto_data = fetch_crypto_data()
if crypto_data is not None:
    print(crypto_data.head())  