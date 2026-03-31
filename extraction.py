import requests
import json
import pandas as pd
import os
from dotenv import load_dotenv
from excel_formatter import csv_to_excel

load_dotenv()

def get_json():
    response = requests.get("https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=100&page=1&sparkline=false")
    data = response.json()
    with open ("crypto_data.json","w") as f:
        json.dump(data,f,indent=4)
    return(data)

data = get_json()

def get_info(data):
    master_data = []
    for values in data:
        entry = {}
        entry = {
            "id" : values.get("id","Not Found"),
            "symbol" : values.get("symbol","Not Found"),
            "current price" : values.get("current_price","Not Found"),
            "market_cap_change_percentage_24h" : values.get("market_cap_change_percentage_24h","Not Found")
        }
        master_data.append(entry)
    with open ("master_data.json","w") as f:
        json.dump(master_data,f,indent=4)
    return(master_data)

master_data =  get_info(data)

def make_csv(master_data):
    Crypto_Data = pd.DataFrame(master_data)
    Crypto_Data.to_csv("Crypto_Data.csv",index=False)
    Crypto_Data = Crypto_Data.drop(index=[0,1])
    Crypto_Data = Crypto_Data[Crypto_Data["market_cap_change_percentage_24h"] < -5.0]
    Crypto_Data.to_csv("Crypto_Data.csv",index=False)
    excel_file_path = csv_to_excel(
        input_path = "Crypto_Data.csv",
        output_path= "Crypto_Data.xlsx",
        report_title="Crypto Daily Update",
        sheet_name="Details",
        include_summary=True
    )
    print("File successfully created")
    return(excel_file_path)

excel_file_path = make_csv(master_data)

def web_hook(excel_file_path):
    with open(excel_file_path,"rb") as f:
        payload_file = {
            "files" : (str(excel_file_path),f,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        }
        respond = requests.post(os.environ.get("N8N_link"),files=payload_file)
    
    print("Transfer Successful!")
    return()

web_hook(excel_file_path)

