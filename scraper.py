 
from bs4 import BeautifulSoup
import requests
import pandas as pd
import os
from collections import defaultdict
import glob
from urllib.parse import urlparse, parse_qs
 
headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel, */*',
        'Accept-Language': 'en-US,en;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
        'Referer': 'https://std.eng.cu.edu.eg/',
        'Origin': 'https://std.eng.cu.edu.eg',
    }
    
URL='https://std.eng.cu.edu.eg/ClassList.aspx?s=1'
page =requests.get(URL,headers=headers)
soup= BeautifulSoup(page.text,'html')

 
table=soup.find('table')

 

columns=table.find_all('th')
column_names=[col.text for col in columns]

 
df=pd.DataFrame(columns=column_names)


 
column_data=table.find_all('tr')[3:]
column_data.pop()

 
for row in column_data:
    row_data=row.find_all('td')
    indvidual_row_data=[data.text.strip() for data in row_data]
    # print(indvidual_row_data)
    if not indvidual_row_data:
        continue
    # print(indvidual_row_data)
    row_link=row.find('a')
    indvidual_row_data.pop()
    indvidual_row_data.append("https://std.eng.cu.edu.eg/"+row_link.get("href"))
    indvidual_row_data[2]=indvidual_row_data[2].split(']')[0]
    length=len(df)
    df.loc[length]=indvidual_row_data



df.to_csv('output.csv')

 

download_dir = "downloaded_files"
os.makedirs(download_dir, exist_ok=True)

# df=pd.read_csv('output.csv')
#  Download each file
for index, row in df.iterrows():
    try:
        url = row['Download']
        
        # Parse the URL and extract query parameters
        parsed_url = urlparse(url)
        query_params = parse_qs(parsed_url.query)
        schid = query_params.get('schid', [''])[0]
        filename = f"{row['Code']}_{row['Session Type']}-{schid}.xlsx"
        file_path = os.path.join(download_dir, filename)
            
        print(f'Downloading {f"{row['Code']}_{row['Session Type']}-{schid}.xlsx"}...')
            
        response = requests.get(url,headers=headers,timeout=30, stream=True)
        response.raise_for_status()
            
        with open(file_path, 'wb') as f:
            f.write(response.content)
            
        print(f"Successfully saved: {f"{row['Name']}_{row['Session Type']}-{schid}.xlsx"}")
            
    except Exception as e:
        print(f"Error downloading {url}: {e}")



def process_excel_files(name):
    temp = os.path.dirname(os.path.abspath(__file__))
    current_dir = os.path.join(temp, 'downloaded_files')

    
    excel_files = glob.glob(os.path.join(current_dir, "*.xlsx"))
    
    excel_files = [f for f in excel_files if not f.endswith('results.xlsx')]
    
    if not excel_files:
        print("No Excel files (.xlsx) found in the current directory.")
        return
    
    print(f"Found {len(excel_files)} Excel file(s): {[os.path.basename(f) for f in excel_files]}")
    
    student_data = defaultdict(lambda: {'count': 0, 'files': set()})
    
    count=False
    for excel_file in excel_files:
        filename = os.path.basename(excel_file)
        subj=filename.split('-')[0]
        try:
            df = pd.read_excel(excel_file)
            
            if 'Student Name' not in df.columns:
                print(f"Warning: 'Student Name' column not found in {filename}")
                continue
            
            for student_name in df['Student Name'].dropna().astype(str).str.strip(): 
                    if name in student_name :
                        count=True
                        break
                
            if count is True:
               for student_name in df['Student Name'].dropna().astype(str).str.strip():
                if student_name:  # Skip empty names
                    student_data[student_name]['count'] += 1
                    student_data[student_name]['files'].add(subj)
            count=False
            
        except Exception as e:
            print(f"Error processing {filename}: {e}")
    
    if not student_data:
        print("No student data found in any Excel files.")
        return
    
    output_data = []
    for student_name, info in student_data.items():
        output_data.append({
            'Student Name': student_name,
            'Frequency': info['count'],
            'File(s) of Appearance': ', '.join(sorted(info['files']))
        })
    
    results_df = pd.DataFrame(output_data)
    results_df = results_df.sort_values('Frequency', ascending=False)
    
    output_file = os.path.join(temp, "results.xlsx")
    try:
        results_df.to_excel(output_file, index=False)
        print(f"Results successfully written to {output_file}")
        print(f"Total unique students found: {len(results_df)}")
        
    except Exception as e:
        print(f"Error writing to Excel file: {e}")
process_excel_files('زياد عبدالرؤف حسن ابوالحديد')
