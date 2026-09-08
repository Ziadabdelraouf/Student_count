import os
import time
import shutil
import uuid
import math
import re
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- CONFIGURATION ---
# Use 4 workers in CI/GitHub Actions to prevent out-of-memory errors; 
# local runs can override this via an env var or use 4 by default.
NUM_BROWSERS = 4 if os.getenv("CI") else int(os.getenv("NUM_BROWSERS", "6"))
URL = "https://std.eng.cu.edu.eg/ClassList.aspx?s=1"
DOWNLOAD_DIR = "downloaded_files"
# ---------------------

def setup_driver(unique_temp_dir):
    """Configures a headless Chrome/Brave instance compatible with Linux and Windows."""
    chrome_options = Options()

    # Only point to Brave if running locally on Windows and the path exists
    local_brave_path = r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"
    if os.name == 'nt' and os.path.exists(local_brave_path) and not os.getenv("CI"):
        chrome_options.binary_location = local_brave_path

    # Headless mode for server/CI environments
    chrome_options.add_argument("--headless=new") 
    
    # Stability and speed flags for Linux / container runners
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--blink-settings=imagesEnabled=false") 
    
    prefs = {
        "download.default_directory": os.path.abspath(unique_temp_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    service = Service() 
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def wait_for_downloads(folder_path, timeout=60):
    """Waits until .crdownload files are gone."""
    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            files = os.listdir(folder_path)
        except Exception:
            return False
        
        if not files or any(f.endswith('.crdownload') or f.endswith('.tmp') for f in files):
            time.sleep(0.5)
        else:
            return True
    return False

def rename_latest_file(folder_path, code, name, session_type, day, start_time, end_time):
    """Renames the file in the specific worker's folder with sanitization."""
    try:
        files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if not f.endswith('.crdownload')]
        if not files:
            return None
            
        latest_file = max(files, key=os.path.getctime)
        _, ext = os.path.splitext(latest_file)
        if not ext:
            ext = ".xlsx"

        # Sanitize filenames to remove illegal filesystem characters
        safe_code = re.sub(r'[<>:"/\\|?*]', '-', code).strip()
        safe_name = re.sub(r'[<>:"/\\|?*]', '-', name).strip()
        safe_session = re.sub(r'[<>:"/\\|?*]', '-', session_type).strip()
        safe_day = re.sub(r'[<>:"/\\|?*]', '-', day).strip()
        safe_start_time = re.sub(r'[<>:"/\\|?*]', '-', start_time).strip()
        safe_end_time = re.sub(r'[<>:"/\\|?*]', '-', end_time).strip()

        new_name = f"{safe_code}_{safe_name}_{safe_session}_{safe_day}_{safe_start_time}-{safe_end_time}_{str(uuid.uuid4())[:8]}{ext}"
        new_path = os.path.join(folder_path, new_name)
        
        if os.path.exists(new_path):
            os.remove(new_path)
        os.rename(latest_file, new_path)
        return new_path
    except Exception as e:
        print(f"⚠️ Rename Error: {e}")
        return None

def process_chunk(chunk_indices, worker_id, permanent_dir):
    """Controls ONE browser handling a specific list of files."""
    worker_temp_dir = os.path.abspath(f"temp_worker_{worker_id}")
    if os.path.exists(worker_temp_dir):
        shutil.rmtree(worker_temp_dir)
    os.makedirs(worker_temp_dir, exist_ok=True)
    
    driver = None
    files_downloaded = 0
    
    try:
        print(f"🔹 [Worker {worker_id}] Launching Browser...")
        driver = setup_driver(worker_temp_dir)
        
        driver.get(URL)
        wait = WebDriverWait(driver, 30)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))

        for table_index in chunk_indices:
            try:
                # 1. Clean temp dir
                for f in os.listdir(worker_temp_dir):
                    try:
                        os.remove(os.path.join(worker_temp_dir, f))
                    except Exception:
                        pass
                
                # 2. Re-find row
                current_rows = driver.find_elements(By.CSS_SELECTOR, "table tr")
                target_idx = 3 + table_index 
                
                if target_idx >= len(current_rows):
                    continue

                row = current_rows[target_idx]
                cols = row.find_elements(By.TAG_NAME, "td")
                if not cols:
                    continue

                code = cols[0].text.strip()
                name = cols[1].text.strip()
                session_type = cols[3].text.strip()
                day = cols[4].text.strip()
                start_time = cols[5].text
                end_time = cols[6].text

                link = row.find_element(By.TAG_NAME, "a")
                print(f"⬇️ [Worker {worker_id}] Clicking: {code}")
                driver.execute_script("arguments[0].click();", link)
                
                # 3. Wait & Rename
                if wait_for_downloads(worker_temp_dir, timeout=45):
                    final_path = rename_latest_file(worker_temp_dir, code, name, session_type, day, start_time, end_time)
                    if final_path:
                        shutil.move(final_path, os.path.join(permanent_dir, os.path.basename(final_path)))
                        files_downloaded += 1
                        print(f"✅ [Worker {worker_id}] Success: {code}")
                else:
                    print(f"❌ [Worker {worker_id}] Timeout: {code}")

                time.sleep(1)

            except Exception as e:
                print(f"⚠️ [Worker {worker_id}] Error: {e}")
                try:
                    driver.get(URL)
                    time.sleep(2)
                except Exception:
                    pass

    except Exception as e:
        print(f"🔥 [Worker {worker_id}] Crash: {e}")
    finally:
        if driver:
            driver.quit()
        if os.path.exists(worker_temp_dir):
            shutil.rmtree(worker_temp_dir)
    
    return files_downloaded

def main():
    start_time = datetime.now()
    print(f"🚀 Starting Optimized Parallel Downloader at {start_time.isoformat()}")
    
    permanent_dir = os.path.abspath(DOWNLOAD_DIR)
    os.makedirs(permanent_dir, exist_ok=True)

    # --- Phase 1: Scan --- 
    print("🔍 Scanning page to split work...")
    temp_driver = setup_driver("temp_scan")
    try:
        temp_driver.get(URL)
        WebDriverWait(temp_driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
        rows = temp_driver.find_elements(By.TAG_NAME, "tr")
        data_rows = rows[3:] 
        if data_rows:
            data_rows.pop() 
        
        total_rows = len(data_rows)
        all_indices = list(range(total_rows))
        print(f"📋 Found {total_rows} files. Launching {NUM_BROWSERS} browsers.")
    except Exception as e:
        print(f"❌ Scan failed: {e}")
        return
    finally:
        temp_driver.quit()
        if os.path.exists("temp_scan"):
            shutil.rmtree("temp_scan")

    # --- Phase 2: Split ---
    chunk_size = math.ceil(len(all_indices) / NUM_BROWSERS)
    chunks = [all_indices[i:i + chunk_size] for i in range(0, len(all_indices), chunk_size)]

    # --- Phase 3: Execute ---
    total_success = 0
    with ThreadPoolExecutor(max_workers=NUM_BROWSERS) as executor:
        futures = []
        for i, chunk in enumerate(chunks):
            futures.append(executor.submit(process_chunk, chunk, i + 1, permanent_dir))
            
        for future in as_completed(futures):
            total_success += future.result()

    end_time = datetime.now()
    duration = end_time - start_time
    print(f"🎉 DONE! Total files: {total_success}")
    print(f"⏱️ Time taken: {duration}")

if __name__ == "__main__":
    main()
