import os, time, glob, pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# -----------------------------
# Configuration
# -----------------------------
DOWNLOAD_DIR = os.environ.get("DOWNLOAD_DIR", os.path.join(os.getcwd(), "downloads"))
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

chrome_options = Options()
chrome_options.add_argument("--headless=new")           # headless = run invisible
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
wait = WebDriverWait(driver, 20)

try:
    print("Opening Infoshare...")
    driver.get("https://infoshare.stats.govt.nz/")
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Browse"))).click()
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Imports and exports"))).click()
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Imports - Summary Data - IMP"))).click()
    wait.until(EC.element_to_be_clickable(
        (By.LINK_TEXT, "Imports - confidential - values and quantities (Qrtly-Mar/Jun/Sep/Dec)"))
    ).click()

    # Select all variables (simplified — adjust for what you need)
    print("Selecting all variables...")
    select_all_btns = driver.find_elements(By.XPATH,
        "//span[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'select all')]")
    for btn in select_all_btns:
        driver.execute_script("arguments[0].scrollIntoView(true);", btn)
        time.sleep(1)
        try:
            btn.click()
        except Exception:
            pass

    # Choose CSV file type
    from selenium.webdriver.support.ui import Select
    print("Choosing CSV format...")
    dropdown = driver.find_element(By.NAME, "ctl00$MainContent$dlOutputOptions")
    select = Select(dropdown)
    found = False
    for opt in select.options:
        if "csv" in opt.text.lower():
            select.select_by_visible_text(opt.text)
            print(f"Selected {opt.text}")
            found = True
            break
    if not found:
        print("CSV option not found, using default.")

    # Click Go button
    print("Clicking Go...")
    go_btn = driver.find_element(By.NAME, "ctl00$MainContent$btnGo")
    driver.execute_script("arguments[0].scrollIntoView(true);", go_btn)
    go_btn.click()

    # Wait for file download
    print("Waiting for download...")
    for i in range(90):
        files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")) + \
                glob.glob(os.path.join(DOWNLOAD_DIR, "*.xls")) + \
                glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx"))
        if files:
            latest = max(files, key=os.path.getmtime)
            print(f"✅ Downloaded: {latest}")
            break
        time.sleep(1)
    else:
        raise TimeoutError("File not downloaded in 90 s")

    # Optional: clean / convert to master CSV
    df = pd.read_excel(latest) if latest.endswith(("xls", "xlsx")) else pd.read_csv(latest)
    df.dropna(how="all", inplace=True)
    out_path = os.path.join(DOWNLOAD_DIR, "infoshare_master.csv")
    df.to_csv(out_path, index=False)
    print(f"Saved cleaned file: {out_path}")

except Exception as e:
    import traceback
    print("❌ Error:", e)
    traceback.print_exc()
finally:
    driver.quit()
    print("Browser closed.")
