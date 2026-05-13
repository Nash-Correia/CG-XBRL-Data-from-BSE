"""
Diagnostic script: opens one BSE corporate-governance page,
dumps the inner-HTML of candidate table containers and prints
every <tr> text it finds so we can identify the right CSS selector.
"""
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

URL = "https://www.bseindia.com/stock-share-price/abb-india-ltd/abb/500002/flag/7/corporate-governance"

chrome_opts = Options()
chrome_opts.add_argument("--disable-gpu")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_opts)
driver.get(URL)

# Wait up to 20 s for ANY table row to appear anywhere on the page
try:
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "table tr"))
    )
except:
    print("No table found within 20 s")

# Extra settle time for AJAX
time.sleep(3)

# ── 1. Print all table rows and their text (trimmed) ──────────────────────────
print("\n=== ALL <tr> TEXT ON PAGE ===")
rows = driver.find_elements(By.CSS_SELECTOR, "tr")
print(f"Total <tr> elements: {len(rows)}")
for i, tr in enumerate(rows):
    txt = tr.text.strip()
    if txt:
        print(f"  [{i}] {txt[:120]}")

# ── 2. Dump outer HTML of #deribody if it exists ──────────────────────────────
print("\n=== #deribody children (div count) ===")
try:
    deribody = driver.find_element(By.ID, "deribody")
    divs = deribody.find_elements(By.XPATH, "./div")
    print(f"#deribody has {len(divs)} direct <div> children")
    for i, d in enumerate(divs):
        print(f"  div[{i}] classes={d.get_attribute('class')!r}  text[:80]={d.text[:80]!r}")
except Exception as e:
    print(f"  #deribody not found: {e}")

# ── 3. Look for any element whose text contains "XBRL" ───────────────────────
print("\n=== Elements containing 'XBRL' ===")
xbrl_els = driver.find_elements(By.XPATH, "//*[contains(text(),'XBRL') or contains(text(),'xbrl')]")
for el in xbrl_els:
    print(f"  tag={el.tag_name}  text={el.text[:120]!r}")

# ── 4. Look for links that end in .xml ───────────────────────────────────────
print("\n=== Links ending in .xml ===")
xml_links = driver.find_elements(By.XPATH, "//a[contains(@href,'.xml') or contains(@href,'.XML')]")
for a in xml_links:
    print(f"  href={a.get_attribute('href')}  text={a.text[:60]!r}")

driver.quit()
