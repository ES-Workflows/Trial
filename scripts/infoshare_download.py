import os, time, glob, pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# -----------------------------
# Configuration
# -----------------------------
DOWNLOAD_DIR = os.environ.get("DOWNLOAD_DIR", os.path.join(os.getcwd(), "downloads"))
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

chrome_options = Options()
chrome_options.add_argument("--headless=new")     # run headless on CI
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)

# ---- IMPORTANT: allow downloads in headless mode via Chrome DevTools ----
try:
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": DOWNLOAD_DIR}
    )
except Exception:
    # Not fatal on newer Chromes, but keep going
    pass

wait = WebDriverWait(driver, 30)

def wait_for_download(timeout=120):
    """Poll the download dir until a file appears."""
    start = time.time()
    while time.time() - start < timeout:
        files = glob.glob(os.path.join(DOWNLOAD_DIR, "*"))
        # ignore temporary files .crdownload
        files = [f for f in files if not f.endswith(".crdownload")]
        if files:
            return max(files, key=os.path.getmtime)
        time.sleep(1)
    return None

try:
    print("Opening Infoshare…")
    driver.get("https://infoshare.stats.govt.nz/")

    # Navigate
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Browse"))).click()
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Imports and exports"))).click()
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Imports - Summary Data - IMP"))).click()
    # Use the QUARTERLY dataset (adjust if you really want Annual)
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Imports - confidential - values and quantities (Qrtly-Mar/Jun/Sep/Dec)"))).click()

    # Select all variables (fast demo; later you can be precise)
    print("Selecting all variables…")
    select_all_btns = driver.find_elements(
        By.XPATH,
        "//span[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'select all')]"
    )
    for btn in select_all_btns:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        time.sleep(0.5)
        try:
            btn.click()
        except Exception:
            pass

    # Choose CSV file type (fallback: any option containing 'csv')
    print("Choosing CSV format…")
    dropdown = wait.until(EC.presence_of_element_located((By.NAME, "ctl00$MainContent$dlOutputOptions")))
    sel = Select(dropdown)
    chosen = False
    for opt in sel.options:
        if "csv" in opt.text.lower():
            sel.select_by_visible_text(opt.text)
            print(f"Selected: {opt.text}")
            chosen = True
            break
    if not chosen:
        print("CSV option not found; using default export type.")

    # Click Go
    print("Clicking Go…")
    try:
        go_btn = driver.find_element(By.NAME, "ctl00$MainContent$btnGo")
    except Exception:
        go_btn = driver.find_element(By.XPATH, "//input[@type='submit' and contains(@value,'Go')]")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", go_btn)
    go_btn.click()

    print(f"Waiting for download in {DOWNLOAD_DIR} …")
    downloaded = wait_for_download(timeout=120)
    if not downloaded:
        raise RuntimeError("File not downloaded in time")

    print(f"✅ Downloaded: {downloaded}")

    # Optional: Save a normalized CSV for Power BI
    out_csv = os.path.join(DOWNLOAD_DIR, "infoshare_master.csv")
    if downloaded.lower().endswith(".csv"):
        df = pd.read_csv(downloaded)
    else:
        df = pd.read_excel(downloaded)
    df.dropna(how="all", inplace=True)
    df.to_csv(out_csv, index=False)
    print(f"Saved cleaned file: {out_csv}")

except Exception as e:
    import traceback
    print("❌ Error:", e)
    traceback.print_exc()
finally:
    driver.quit()
    print("Browser closed.")
    # Also drop a small marker so you know the step ran even if download failed
    with open(os.path.join(DOWNLOAD_DIR, "_run_marker.txt"), "w") as f:
        f.write("selenium step completed\n")
