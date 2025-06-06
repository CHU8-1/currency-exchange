import requests
import pandas as pd
import plotly.express as px
from datetime import datetime
from os.path import exists

# 設定 API 參數
BASE_URL = "https://api.frankfurter.app"
BASE_CURRENCY = "USD"
TARGET_CURRENCIES = ["TWD", "EUR", "JPY", "GBP", "CNY", "KRW"]

# 匯率查詢函式
def fetch_exchange_rates():
    symbols = ",".join(TARGET_CURRENCIES)
    url = f"{BASE_URL}/latest?from={BASE_CURRENCY}&to={symbols}"
    response = requests.get(url)
    data = response.json()

    rates = data['rates']
    df = pd.DataFrame(list(rates.items()), columns=['貨幣', '匯率'])
    df['時間'] = data['date'] + "（以 USD 為基準）"
    return df

# 匯率圖表
def plot_exchange(df):
    fig = px.bar(df, x='貨幣', y='匯率',
                 title='即時匯率（以 USD 為基準）',
                 color='貨幣',
                 text='匯率')
    fig.update_traces(texttemplate='%{text:.2f}', textposition='outside')
    fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
    fig.show()

# 匯出 Excel
def save_to_excel(df):
    filename = '匯率報表.xlsx'
    file_exists = exists(filename)
    mode = 'a' if file_exists else 'w'
    sheet_name = datetime.now().strftime('%Y-%m-%d')

    with pd.ExcelWriter(filename, engine='openpyxl', mode=mode, if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# 主程式
if __name__ == '__main__':
    df = fetch_exchange_rates()
    print(df)
    plot_exchange(df)
    save_to_excel(df)
    print("✅ 匯率資料已儲存為 Excel。")
