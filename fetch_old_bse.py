import os
import re
import pandas as pd
import requests
import urllib.parse
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# CONFIG
YEARS        = [2023, 2024]
EXCEL_PATH   = r"companies.xlsx"
DOWNLOAD_DIR = r"XBRL Files"
BASE_URL     = "https://www.bseindia.com/stock-share-price/{name}/{nse}/{bse}/flag/7/corporate-governance"
BSE_ROOT     = "https://www.bseindia.com"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# SELENIUM SETUP
chrome_opts = Options()
# chrome_opts.add_argument("--headless")
chrome_opts.add_argument("--disable-gpu")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                          options=chrome_opts)

df = pd.read_excel(EXCEL_PATH, usecols=["Name", "BSE", "NSE"])
for idx, row in df.iterrows():
    if pd.isna(row["BSE"]) or pd.isna(row["NSE"]):
        print(f"[{idx+1}/{len(df)}] Skipping — missing BSE/NSE for {row['Name']}")
        continue
    name = str(row["Name"]).strip().lower().replace("&", "").replace(" ", "-")
    bse  = str(int(float(str(row["BSE"]).strip())))[:6]
    nse  = str(row["NSE"]).strip().lower().replace("&", "")
    url  = BASE_URL.format(name=name, nse=nse, bse=bse)
    print(f"[{idx+1}/{len(df)}] → {url}")
    driver.get(url)

    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//a[normalize-space(text())='XBRL']"))
        )
    except:
        print(f"   • No XBRL table found for {nse}/{bse}, skipping.")
        continue

    xbrl_links = driver.find_elements(By.XPATH, "//a[normalize-space(text())='XBRL']")

    # prepare requests session with Selenium cookies & headers (shared across years)
    session = requests.Session()
    for c in driver.get_cookies():
        session.cookies.set(c["name"], c["value"])
    session.headers.update({
        "User-Agent": driver.execute_script("return navigator.userAgent;"),
        "Referer": url
    })

    # make NSE symbol safe as a filename (remove weird chars)
    nse_symbol = re.sub(r'[^A-Za-z0-9._-]+', '_', str(row["NSE"]).strip())

    for yr in YEARS:
        # find first XBRL link whose immediate parent <tr> contains the year label
        yr_label = f"{yr} - {yr + 1}"
        target_el = None
        for a in xbrl_links:
            try:
                parent_tr = a.find_element(By.XPATH, "./ancestor::tr[1]")
                if yr_label in parent_tr.text:
                    target_el = a
                    break
            except Exception:
                continue

        if target_el is None:
            print(f"   • No XBRL for {yr_label} — {nse}/{bse}, skipping.")
            continue

        rel_href = target_el.get_attribute("href")
        xml_url  = rel_href if rel_href.startswith("http") else BSE_ROOT + rel_href

        out_path = os.path.join(f"XBRL_{yr}", f"{nse_symbol}.xml")
        try:
            resp = session.get(xml_url, timeout=15, stream=True)
            resp.raise_for_status()
            with open(out_path, "wb") as f:
                for chunk in resp.iter_content(8192):
                    f.write(chunk)
            print(f"   ✔ [{yr}] Saved: {out_path}")
        except Exception as e:
            print(f"   ! [{yr}] Download failed for {nse}/{bse}: {e}")

driver.quit()
