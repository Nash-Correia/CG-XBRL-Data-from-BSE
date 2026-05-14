import os
import re
import subprocess
import pandas as pd
import requests
import urllib.parse
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# CONFIG
EXCEL_PATH   = r"companies.xlsx"
DOWNLOAD_DIR = r"XBRL_2025"
BASE_URL     = "https://www.bseindia.com/stock-share-price/{name}/{nse}/{bse}/flag/7/corporate-governance"
BSE_ROOT     = "https://www.bseindia.com"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# SELENIUM SETUP
edge_opts = Options()
# edge_opts.add_argument("--headless")
edge_opts.add_argument("--disable-gpu")
edge_opts.add_argument("--log-level=3")
edge_opts.add_experimental_option("excludeSwitches", ["enable-logging"])
driver = webdriver.Edge(
    service=Service(log_output=subprocess.DEVNULL),
    options=edge_opts
)

df = pd.read_excel(EXCEL_PATH, sheet_name="Active", usecols=["Name", "BSE", "NSE"])
for idx, row in df.iterrows():
  try:
    name = str(row["Name"]).strip().lower().replace("&", "").replace(" ", "-")
    name = urllib.parse.quote(name, safe="-")
    bse  = str(int(float(str(row["BSE"]).strip())))[:6]
    nse  = str(row["NSE"]).strip().lower().replace("&", "")
    url  = BASE_URL.format(name=name, nse=urllib.parse.quote(nse, safe="-"), bse=bse)
    print(f"[{idx+1}/{len(df)}] → {url}")
    driver.get(url)

    try:
        # Wait for at least one .xml link to appear in the table (works for both old
        # and new BSE page layouts — the #deribody selector no longer exists).
        el = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.XPATH, "(//a[contains(@href,'.xml') or contains(@href,'.XML')])[1]")
            )
        )
    except:
        print(f"   • No XBRL link for {nse}/{bse}, skipping.")
        continue

    href = el.get_attribute("href") or ""
    xml_url = href if href.startswith("http") else BSE_ROOT + href

    # prepare requests session with Selenium cookies & headers
    session = requests.Session()
    for c in driver.get_cookies():
        session.cookies.set(c["name"], c["value"])
    session.headers.update({
        "User-Agent": driver.execute_script("return navigator.userAgent;"),
        "Referer": url
    })

    # make NSE symbol safe as a filename (remove weird chars)
    nse_symbol = re.sub(r'[^A-Za-z0-9._-]+', '_', str(row["NSE"]).strip())
    out_path = os.path.join(DOWNLOAD_DIR, f"{nse_symbol}.xml")

    try:
        resp = session.get(xml_url, timeout=15, stream=True)
        resp.raise_for_status()
        with open(out_path, "wb") as f:
            for chunk in resp.iter_content(8192):
                f.write(chunk)
        print(f"   ✔ Saved: {out_path}")
    except Exception as e:
        print(f"   ! Download failed for {nse}/{bse}: {e}")

  except Exception as e:
    print(f"   ! Unexpected error for row {idx+1}: {e}, skipping.")
    continue

driver.quit()
