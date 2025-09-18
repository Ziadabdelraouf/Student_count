import pandas as pd
import os
from collections import defaultdict
import glob
import sys
import json
from datetime import datetime

def process_excel_files(name):
    current_dir = os.path.dirname(os.path.abspath(__file__))
    download_dir = os.path.join(current_dir, "downloaded_files")
    
    excel_files = glob.glob(os.path.join(download_dir, "*.xlsx"))
    
    if not excel_files:
        print("No Excel files (.xlsx) found in the downloaded_files directory.")
        return {"error": "No Excel files found"}
    
    print(f"Found {len(excel_files)} Excel file(s): {[os.path.basename(f) for f in excel_files]}")
    
    student_data = defaultdict(lambda: {'count': 0, 'files': set()})
    
    count = False
    for excel_file in excel_files:
        filename = os.path.basename(excel_file)
        subj = filename.split('-')[0]
        try:
            df = pd.read_excel(excel_file)
            
            if 'Student Name' not in df.columns:
                print(f"Warning: 'Student Name' column not found in {filename}")
                continue
            
            for student_name in df['Student Name'].dropna().astype(str).str.strip(): 
                if name in student_name:
                    count = True
                    break
                
            if count is True:
                for student_name in df['Student Name'].dropna().astype(str).str.strip():
                    if student_name:  # Skip empty names
                        student_data[student_name]['count'] += 1
                        student_data[student_name]['files'].add(subj)
            count = False
            
        except Exception as e:
            print(f"Error processing {filename}: {e}")
    
    if not student_data:
        print("No student data found in any Excel files.")
        return {"error": "No matching students found"}
    
    output_data = []
    for student_name, info in student_data.items():
        output_data.append({
            'Student Name': student_name,
            'Frequency': info['count'],
            'File(s) of Appearance': ', '.join(sorted(info['files']))
        })
    
    results_df = pd.DataFrame(output_data)
    results_df = results_df.sort_values('Frequency', ascending=False)
    
    output_file = "results.xlsx"
    try:
        results_df.to_excel(output_file, index=False)
        print(f"Results successfully written to {output_file}")
        print(f"Total unique students found: {len(results_df)}")
        
        # Also create a JSON summary for easy webhook responses
        summary = {
            "search_parameter": name,
            "total_matches": len(results_df),
            "top_matches": results_df.head(5).to_dict('records'),
            "last_updated": datetime.now().isoformat()
        }
        
        with open("search_summary.json", "w") as f:
            json.dump(summary, f)
            
        return summary
        
    except Exception as e:
        print(f"Error writing to Excel file: {e}")
        return {"error": str(e)}

def main():
    if len(sys.argv) > 1:
        namee = sys.argv[1]
    else:
        namee = "زياد عبدالرؤف حسن ابوالحديد"
    
    result = process_excel_files(namee)
    print(f"Search result: {result}")

if __name__ == "__main__":
    main()