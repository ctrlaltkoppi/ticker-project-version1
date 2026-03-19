import pandas as pd
import re
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

EXCEL_FILE = "/home/njjlim/python_projects/NEOS Test 3_03202026_IN.xlsx"
OUTPUT_FILE = "/home/njjlim/python_projects/NEOS Test 3_03202026_OUT.xlsx"


def clean_percent_to_float(text):
    if text is None:
        return None
    match = re.search(r"(\d+(?:\.\d+)?)\s*%", str(text))
    if match:
        return float(match.group(1))
    return None


def get_neos_distribution_rate(driver, url):
    try:

        driver.get(url)

        WebDriverWait(driver, 10).until(
             EC.presence_of_element_located((By.CLASS_NAME, "stat"))
        )

        # Find the exact small tag for Distribution Rate, then its next sibling div.stat
        smalls = driver.find_elements(By.TAG_NAME, "small")

        for s in smalls:
            title = s.get_attribute("data-original-title")
            if title and title.strip() == "Distribution Rate":
                stat = s.find_element(By.XPATH, "following-sibling::div[contains(@class,'stat')]")
                return stat.text.strip()

        return None

    except Exception as e:
        print(f"Error scraping {url}: {e}")
        return None


print("Opening Excel file...")
df = pd.read_excel(EXCEL_FILE)
df["Yield"] = pd.to_numeric(df["Yield"], errors="coerce")

chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=chrome_options)

for index, row in df.iterrows():
    ticker = row["Ticker"]
    website = row["Website Source"]

    print(f"Checking: {ticker}")

    rate_text = get_neos_distribution_rate(driver, website)
    print(f"Raw extracted text: {rate_text}")

    if rate_text:
        numeric_rate = float(rate_text.replace("%","")) / 100
        df.at[index, "Yield"] = numeric_rate
        print(f"Stored: {numeric_rate}")
    else:
        df.at[index, "Yield"] = None
        print("No rate found")

driver.quit()

df.to_excel(OUTPUT_FILE, index=False)

print("\nFinished.")
print(f"Updated file saved to: {OUTPUT_FILE}")
