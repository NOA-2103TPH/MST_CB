import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException

try:
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    print("Please install webdriver_manager: pip install webdriver-manager")
    exit(1)

import sys
import io

# Force UTF-8 encoding for stdout/stderr to handle Vietnamese characters
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Logging check
log_file = open("lookup.log", "w", encoding="utf-8")
def log(msg):
    print(msg)
    log_file.write(str(msg) + "\n")
    log_file.flush()

# Configuration (allow override via env vars)
INPUT_FILE = os.getenv("INPUT_FILE", "data/data.xlsx")
OUTPUT_FILE = os.getenv("OUTPUT_FILE", "result/result5.xlsx")

# ---------- Small Selenium helpers to reduce stale-element errors ----------

def safe_click(driver, locator, retries=3, timeout=10):
    """Click an element; if it goes stale, re-locate and retry."""
    for _ in range(retries):
        try:
            el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(locator))
            el.click()
            return
        except StaleElementReferenceException:
            continue
    raise RuntimeError("element stayed stale after retries")

def wait_for_presence(driver, locator, timeout=10):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located(locator))

def dismiss_popups(driver):
    """Best-effort close common ad/consent overlays so search box becomes clickable."""
    selectors = [
        "button[aria-label='Close']",
        "button.close",
        "div.modal.show button.close",
        "div[role='dialog'] button[aria-label='Close']",
        "#qc-cmp2-ui button[mode='primary']",
        ".qc-cmp2-summary-buttons button",
    ]
    for css in selectors:
        try:
            btns = driver.find_elements(By.CSS_SELECTOR, css)
            if btns:
                btns[0].click()
                return True
        except Exception:
            continue
    return False

def init_driver(headless=False):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument('--headless=new')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')

    # Support Streamlit Cloud/containers where Chrome is at a custom path
    chrome_bin = os.getenv("CHROME_BIN")
    if chrome_bin and os.path.exists(chrome_bin):
        options.binary_location = chrome_bin

    driver_path = os.getenv("CHROMEDRIVER_PATH")
    if driver_path and os.path.exists(driver_path):
        service = Service(driver_path)
    else:
        service = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=service, options=options)
    return driver

def lookup_mst(driver, cccd, log_fn=log):
    log_fn(f"Looking up MST for CCCD: {cccd}")
    
    try:
        driver.get("https://masothue.com")
        
        # Wait loop for search box to be ready (attempt to auto-close ads/consent)
        search_box = None
        max_retries = 30 # Wait up to 60 seconds (30 * 2s)
        
        for i in range(max_retries):
            try:
                dismiss_popups(driver)
                # Try to find and click the search box
                wait = WebDriverWait(driver, 2)
                search_box = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='q']")))
                search_box.click() # Test interaction
                log_fn("Search box is ready.")
                break
            except Exception:
                log_fn(f"Waiting for ads to be closed/search box to be ready... ({i+1}/{max_retries})")
                if (i + 1) % 10 == 0:
                    driver.refresh()
                time.sleep(2)
                
        if not search_box:
            return "lỗi: quá thời gian chờ tắt quảng cáo", "", ""

        # Interaction with search box (robust against staleness)
        search_box_locator = (By.CSS_SELECTOR, "input[name='q']")
        search_btn_locator = (By.CSS_SELECTOR, "button.search-btn, button[type='submit']")

        try:
            box = wait_for_presence(driver, search_box_locator)
            box.clear()
            box.send_keys(cccd)
            time.sleep(0.5)

            try:
                safe_click(driver, search_btn_locator)
                log_fn("Clicked search button.")
            except Exception:
                log_fn("Search button not found, using Enter key.")
                box = wait_for_presence(driver, search_box_locator)
                box.send_keys(Keys.RETURN)

            log_fn(f"Current URL after submit: {driver.current_url}")

        except Exception as e:
            log_fn(f"Standard interaction failed: {e}. Trying JS fallback.")
            try:
                box = wait_for_presence(driver, search_box_locator)
                driver.execute_script("arguments[0].value = '';", box)
                box.send_keys(cccd)
                safe_click(driver, search_btn_locator)
            except Exception as js_e:
                log_fn(f"JS fallback also failed: {js_e}")
                return "lỗi tương tác search", "", ""
        
        # Wait for results
        try:
            wait = WebDriverWait(driver, 10)
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "h1")))
            
            mst = ""
            name = ""
            
            # 1. Try to get Name from H1
            try:
                h1_text = driver.find_element(By.TAG_NAME, "h1").text.strip()
                if " - " in h1_text:
                    parts = h1_text.split(" - ", 1)
                    if not name:
                         name = parts[1].strip()
            except:
                pass

            # 2. Extract MST and Name from table
            try:
                mst_element = driver.find_element(By.XPATH, "//td[contains(., 'Mã số thuế')]/following-sibling::td")
                mst = mst_element.text.strip()
            except:
                try:
                    mst = driver.find_element(By.CSS_SELECTOR, "[itemprop='taxID']").text.strip()
                except: pass

            if not name:
                try:
                    name_element = driver.find_element(By.XPATH, "//td[contains(., 'Người đại diện')]/following-sibling::td")
                    name = name_element.text.strip()
                except:
                    try:
                        name = driver.find_element(By.CSS_SELECTOR, "[itemprop='name']").text.strip()
                    except: pass
            
            # Final check
            if mst:
                return "thành công", mst, name
            else:
                if "Search" in driver.current_url:
                     return "không tìm thấy (dạng danh sách)", "", ""
                
                with open("page_dump.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                log_fn("Dumped page source to page_dump.html due to failure.")
                return "không tìm thấy thông tin chi tiết", "", ""

        except Exception as e:
            log_fn(f"Timeout or parsed error: {e}")
            with open("page_dump_error.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            return "kết nối thất bại/timeout", "", ""

    except Exception as e:
        log_fn(f"Error looking up {cccd}: {e}")
        return "lỗi hệ thống", "", ""

def run_lookup(
    input_path,
    output_path,
    log_fn=log,
    batch_size=120,
    rest_seconds=120,
    headless=False,
    progress_fn=None,
):
    """
    Chạy tra cứu toàn bộ file input và lưu ra output. Có thể truyền hàm log_fn để đẩy log lên UI.
    """
    start_time = time.time()
    log_fn("Starting lookup process.")

    df = None
    if os.path.exists(output_path):
        log_fn(f"Found existing result file: {output_path}. Loading to resume...")
        df = pd.read_excel(output_path, dtype=str)
    elif os.path.exists(input_path):
        log_fn(f"Loading input data from {input_path}...")
        df = pd.read_excel(input_path, dtype=str)
    else:
        log_fn(f"Input file {input_path} not found!")
        return

    # Normalize and harmonize column names (handle case/accents)
    col_map = {}
    for col in df.columns:
        key = col.strip().lower()
        if key == 'cccd':
            col_map[col] = 'CCCD'
        elif key == 'mst':
            col_map[col] = 'MST'
        elif key in ('tên', 'ten'):
            col_map[col] = 'Tên'
        elif key in ('trạng thái', 'trang thai', 'trang thái'):
            col_map[col] = 'Trạng thái'
    if col_map:
        df.rename(columns=col_map, inplace=True)

    # Ensure required columns exist
    for col in ['CCCD', 'MST', 'Tên', 'Trạng thái']:
        if col not in df.columns:
            df[col] = ''
    
    # Fill NaN with empty string to avoid float issues if any slipped through
    df.fillna('', inplace=True)

    total_rows = len(df)
    log_fn(f"Loaded data with {total_rows} rows.")
    log_fn(f"Columns: {df.columns.tolist()}")

    driver = init_driver(headless=headless)
    processed_since_break = 0

    try:
        for index, row in df.iterrows():
            status = str(row.get('Trạng thái', '')).strip().lower()
            cccd = str(row.get('CCCD', '')).strip()
            log_fn(f"Row {index}: CCCD='{cccd}', Status='{status}'")

            if not cccd:
                log_fn("CCCD trống, bỏ qua hàng này.")
                continue

            pending_statuses = {'', 'chưa xử lý', 'chua xu ly'}
            if status in pending_statuses or 'lỗi' in status or 'không tìm thấy' in status:
                attempts = 0
                while True:
                    status_new, mst, name = lookup_mst(driver, cccd, log_fn=log_fn)

                    # Retry once with a fresh driver if we hit site/interaction errors
                    if (status_new.startswith('lỗi') or 'kết nối' in status_new) and attempts == 0:
                        attempts += 1
                        log_fn("Encountered lỗi, restarting trình duyệt và thử lại...")
                        try:
                            driver.quit()
                        except Exception:
                            pass
                        driver = init_driver(headless=headless)
                        continue
                    break
                
                df.at[index, 'MST'] = mst
                df.at[index, 'Tên'] = name
                df.at[index, 'Trạng thái'] = status_new
                
                log_fn(f"Processed {cccd}: {status_new}, MST: {mst}, Name: {name}")
                processed_since_break += 1
                
                # Save incrementally to result file
                df.to_excel(output_path, index=False)
                time.sleep(2) # Polite delay

                # Optional rest to avoid throttling
                if batch_size and processed_since_break % batch_size == 0:
                    log_fn(f"Đã xử lý {processed_since_break} bản ghi, nghỉ {rest_seconds} giây...")
                    time.sleep(rest_seconds)

            # Cập nhật tiến trình
            if progress_fn:
                progress_fn(index + 1, total_rows)
                
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        elapsed = time.time() - start_time
        log_fn(f"Done. Thời gian xử lý: {elapsed:.2f} giây.")

def main():
    run_lookup(INPUT_FILE, OUTPUT_FILE)

if __name__ == "__main__":
    main()
