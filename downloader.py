import uuid
from bs4 import BeautifulSoup
import requests
import pandas as pd
import os
from urllib.parse import urlparse, parse_qs
import shutil
import tempfile
from datetime import datetime

def download_files():
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel, */*',
        'Accept-Language': 'en-US,en;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
        'Referer': 'https://std.eng.cu.edu.eg/',
        'Origin': 'https://std.eng.cu.edu.eg',
    }
    
    URL = 'https://std.eng.cu.edu.eg/ClassList.aspx?s=1'
    page = requests.get(URL, headers=headers)
    soup = BeautifulSoup(page.text, 'html.parser')
    
    table = soup.find('table')
    columns = table.find_all('th')
    column_names = [col.text for col in columns]
    
    df = pd.DataFrame(columns=column_names)
    column_data = table.find_all('tr')[3:]
    column_data.pop()
    
    for row in column_data:
        row_data = row.find_all('td')
        individual_row_data = [data.text.strip() for data in row_data]
        if not individual_row_data:
            continue
        row_link = row.find('a')
        individual_row_data.pop()
        individual_row_data.append("https://std.eng.cu.edu.eg/" + row_link.get("href"))
        individual_row_data[2] = individual_row_data[2].split(']')[0]
        length = len(df)
        df.loc[length] = individual_row_data
    
    # Create a temporary directory for downloading
    temp_dir = tempfile.mkdtemp()
    print(f"Downloading files to temporary directory: {temp_dir}")
    
    # Download each file to the temporary directory
    for index, row in df.iterrows():
        try:
            url = row['Download']
            filename = f"{row['Code']}_{row['Session Type']}-{str(uuid.uuid4())[:8]}.xlsx"
            file_path = os.path.join(temp_dir, filename)
                
            print(f"Downloading {filename}")
                
            response = requests.get(url, headers=headers, timeout=30, stream=True)
            response.raise_for_status()
                
            with open(file_path, 'wb') as f:
                f.write(response.content)
                
            print(f"Successfully saved: {filename}")
                
        except Exception as e:
            print(f"Error downloading {url}: {e}")
    
    return temp_dir

def main():
    print(f"Starting download at {datetime.now().isoformat()}")
    
    # Create a permanent directory for storing files between runs
    permanent_dir = "downloaded_files"
    os.makedirs(permanent_dir, exist_ok=True)
    
    # Download files to a temporary directory first
    temp_dir = download_files()
    
    # Replace the old files with the new ones
    if os.path.exists(permanent_dir):
        shutil.rmtree(permanent_dir)
    shutil.move(temp_dir, permanent_dir)
    
    print(f"All files have been updated successfully at {datetime.now().isoformat()}!")
    print(f"Total files downloaded: {len(os.listdir(permanent_dir))}")

if __name__ == "__main__":
    main()