import requests
from datetime import datetime, timedelta
import pandas as pd

# Replace with your NewsAPI API key
api_key = '6111c55b95354ef9bb7d4d161356fb6a'

# Calculate dates for 'yesterday'
yesterday = datetime.now() - timedelta(1)
yesterday_date = yesterday.strftime('%Y-%m-%d')

# Use yesterday's date in 'YYYYMMDD' format for the filename
filename_date = yesterday.strftime('%Y%m%d')

# NewsAPI endpoint for top headlines
url = 'https://newsapi.org/v2/everything'

# Parameters for the API call
params = {
    'q': 'artificial intelligence OR digital Human',
    'from': yesterday_date,
    'to': yesterday_date,
    'sortBy': 'publishedAt',
    'language': 'en',
    'apiKey': api_key,
}

response = requests.get(url, params=params)

# Prepare the filename using yesterday's date
filename = f"AI_News_{filename_date}.xlsx"

# Check if the request was successful
if response.status_code == 200:
    articles = response.json().get('articles', [])
    # Creating a list of dictionaries, each representing an article
    data = [{
        'Title': article['title'],
        'Description': article['description'],
        'URL': article['url']
    } for article in articles]

    # Create a Pandas Excel writer using openpyxl as the engine
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Create a DataFrame for the introductory text and write it to the first sheet
        intro_df = pd.DataFrame({
            'Introduction': [
                "威来可期 (ai-expected.com)是一家专注于AI数字人应用的公司，欢迎各界有识之士交流和加入，联系方式 zhangliwei@ai-expected.com, 手机 15601908986\n\n新闻信息，详见下一页"]
        })
        intro_df.to_excel(writer, index=False, sheet_name='Introduction')

        # Convert the list of dictionaries into a pandas DataFrame for articles and write it to the second sheet
        df = pd.DataFrame(data)
        df.to_excel(writer, index=False, sheet_name='News')

    print(f"News articles successfully saved to {filename}")
else:
    print(f"Failed to fetch news articles. Status code: {response.status_code}")
