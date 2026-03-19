import pandas as pd
import requests
from bs4 import BeautifulSoup
import re

# ----------------------------------
# LOCATION OF YOUR EXCEL FILE
# ----------------------------------

# ------------------------------
# Excel file path (Windows)
# ------------------------------
EXCEL_FILE = "/home/njjlim/Downloads/NEOS Test 1.xlsx"

# Output file
OUTPUT_FILE = "/home/njjlim/Downloads/NEOS_Test_1_UPDATED.xlsx"

# ----------------------------------
# FUNCTION TO GET NEOS DISTRIBUTION RATE
# ----------------------------------

def get_neos_distribution_rate(url):
    try:
        headers = {
            "User-Agent": "Mozilla/5.0"
        }

        response = requests.get(url, headers=headers, timeout=20)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")

        # Find the exact tooltip tag for Distribution Rate
        label = soup.find(
            "small",
            attrs={"data-original-title": re.compile(r"^\s*Distribution Rate\s*$", re.I)}
        )

        if label:
            # Look specifically for the next sibling div with class="stat"
            stat = label.find_next_sibling("div", class_="stat")
            if stat:
                return stat.get_text(strip=True)

        return None

    except Exception as e:
        print(f"Error scraping {url}: {e}")
        return None

# ----------------------------------
# MAIN PROGRAM
# ----------------------------------

print("Opening Excel file...")

df = pd.read_excel(EXCEL_FILE)

for index, row in df.iterrows():

    ticker = row["Ticker"]
    website = row["Website Source"]

    print("Checking:", ticker)

    rate = get_neos_distribution_rate(website)

    if rate:
        numeric_rate = float(rate.replace("%", "").strip())
        df.at[index, "Yield"] = numeric_rate
        print("Stored:", numeric_rate)
    else:
        print("No rate found")

# ----------------------------------
# SAVE UPDATED FILE
# ----------------------------------

OUTPUT_FILE = "/home/njjlim/Downloads/NEOS_Test_1_UPDATED.xlsx"

df.to_excel(OUTPUT_FILE, index=False)

print("\nFinished.")
print("Updated file saved to:")
print(OUTPUT_FILE)
