import os
import re
import time
import math
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# =========================
# Config
# =========================
EDGE_DRIVER_PATH = "msedgedriver.exe"  # change if needed
SITE_URL = "https://findmypollsite.vote.nyc/"
OUTPUT_COLS = ["AD", "ED", "Cong D", "SD", "Council D", "JD"]

# Accept both short headers and bilingual headers
EXACT_NAMES = {
    "house": [
        "HOUSE #", "HOUSE#", "HOUSE NUMBER", "HOUSE NO", "HOUSE NO.",
        "House Number (ex: 3514) / 门牌号码", "House Number / 门牌号码", "House Number / 门牌号"
    ],
    "street": [
        "STREET NAME", "STREET",
        "Street Name (ex: 19Th Ave) / 街道名称"
    ],
    "zip": [
        "ZIP CODE",
        "Zip Code (Ex: 11204) / 邮政编码"
    ],
}

# Regex fallbacks (flexible)
REQUIRED_PATTERNS = {
    "house": [
        r"\bhouse\s*number\b",
        r"\bhouse\s*#(?:\s|$)",   # handles literal "HOUSE #"
        r"\bhouse\s*no\.?\b",
        r"门牌(号|号码)"
    ],
    "street": [
        r"\bstreet\s*name\b", r"\bstreet\b", r"\bst\.?\b",
        r"街道名称", r"街道"
    ],
    "zip": [
        r"\bzip\s*code\b", r"\bzip\b", r"\bpostal\s*code\b",
        r"邮政编码"
    ],
}

# =========================
# Helpers
# =========================
def norm_keep_symbols(s: str) -> str:
    s = re.sub(r"\s+", " ", str(s)).strip().lower()
    s = re.sub(r"\(.*?\)", "", s)  # drop "(ex: 3514)"
    s = s.replace("/", " ")
    return s

def find_exact(df, names):
    target = {n.strip().lower() for n in names}
    for col in df.columns:
        if col.strip().lower() in target:
            return col
    return None

def find_regex(df, patterns):
    best = None
    for col in df.columns:
        c = norm_keep_symbols(col)
        for rank, pat in enumerate(patterns):
            if re.search(pat, c):
                if best is None or (rank, -len(c)) < best[0]:
                    best = ((rank, -len(c)), col)
                break
    return None if best is None else best[1]

def safe_text(x) -> str:
    """Convert any cell to a clean string; handle NaN/float/int safely."""
    if x is None:
        return ""
    # pandas NaN
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    # numeric types
    if isinstance(x, (int,)):
        return str(x)
    if isinstance(x, float):
        if math.isnan(x):
            return ""
        # Avoid trailing .0 for integers
        return str(int(x)) if x.is_integer() else str(x)
    # everything else
    return str(x)

def clean_house(x) -> str:
    s = safe_text(x).strip()
    s = re.sub(r"\.0$", "", s)          # "1876.0" -> "1876"
    s = re.sub(r"\s+", " ", s)
    return s

def clean_street(x) -> str:
    s = safe_text(x).strip().upper()
    s = re.sub(r"\s+", " ", s)          # collapse multiple spaces
    return s

def clean_zip(x) -> str:
    s = safe_text(x).strip()
    s = re.sub(r"[^\d]", "", s)         # keep digits only
    if len(s) == 4:                      # e.g., "1120" -> "01120"
        s = s.zfill(5)
    if len(s) >= 5:
        s = s[:5]                        # use 5-digit ZIP
    return s

def safe_get_text(driver, by, value, fallback=""):
    try:
        elem = driver.find_element(by, value)
        txt = elem.text.strip() if elem and elem.text else ""
        return txt or fallback
    except Exception:
        return fallback

# =========================
# Pick file
# =========================
Tk().withdraw()
file_path = askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
if not file_path:
    print("No file selected. Exiting.")
    raise SystemExit

# =========================
# Read Excel
# =========================
try:
    df = pd.read_excel(file_path, dtype=str)  # try to keep as strings
except Exception as e:
    print(f"Error reading Excel file: {e}")
    raise SystemExit

# =========================
# Detect required columns
# =========================
house_col  = find_exact(df, EXACT_NAMES["house"])  or find_regex(df, REQUIRED_PATTERNS["house"])
street_col = find_exact(df, EXACT_NAMES["street"]) or find_regex(df, REQUIRED_PATTERNS["street"])
zip_col    = find_exact(df, EXACT_NAMES["zip"])    or find_regex(df, REQUIRED_PATTERNS["zip"])

print("\nDetected columns:")
print("  House:  ", house_col)
print("  Street: ", street_col)
print("  Zip:    ", zip_col)

if not all([house_col, street_col, zip_col]):
    print("\n❌ Could not match required columns. Here are all headers I see:")
    for c in df.columns:
        print("  -", c)
    raise SystemExit

# Ensure output columns exist (added at the END if missing)
for col in OUTPUT_COLS:
    if col not in df.columns:
        df[col] = ""

# =========================
# Set up Edge WebDriver
# =========================
options = webdriver.EdgeOptions()
options.add_argument("window-size=1000,800")
driver = webdriver.Edge(service=Service(EDGE_DRIVER_PATH), options=options)
driver.set_window_position(100, 100)

invalid_rows = []

# =========================
# Process each row
# =========================
for index, row in df.iterrows():
    house_number, street_name, zip_code = "", "", ""
    try:
        # Clean inputs safely (prevents 'float' has no attribute 'strip')
        house_number = clean_house(row.get(house_col, ""))
        street_name  = clean_street(row.get(street_col, ""))
        zip_code     = clean_zip(row.get(zip_col, ""))

        if not all([house_number, street_name, zip_code]):
            print(f"[Row {index}] Skipped — missing address data.")
            continue

        print(f"\n[Row {index}] Checking: {house_number} {street_name}, {zip_code}")

        # Fresh page
        driver.get("about:blank")
        time.sleep(0.5)
        driver.get(SITE_URL)
        wait = WebDriverWait(driver, 12)
        time.sleep(1)

        # Fill form
        hn = wait.until(EC.presence_of_element_located((By.ID, "txtHouseNumber")))
        hn.clear(); hn.send_keys(house_number)

        st = wait.until(EC.presence_of_element_located((By.ID, "txtStreetName")))
        st.clear(); st.send_keys(street_name)

        zp = wait.until(EC.presence_of_element_located((By.ID, "txtZipcode")))
        zp.clear(); zp.send_keys(zip_code)

        time.sleep(0.4)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Find My Site']"))).click()
        time.sleep(2)

        # Invalid address message?
        error_msg_element = driver.find_elements(By.ID, "divMessage")
        if error_msg_element and error_msg_element[0].text.strip():
            error_text = error_msg_element[0].text.strip()
            print(f"   ❌ Invalid or unrecognized address: {error_text}")
            invalid_rows.append({
                "row": index,
                "house": house_number,
                "street": street_name,
                "zip": zip_code,
                "reason": error_text
            })
            continue

        # Scrape results
        site_data = {
            "AD":         safe_get_text(driver, By.ID, "assembly_district"),
            "ED":         safe_get_text(driver, By.ID, "election_district").split("/")[0],
            "Cong D":     safe_get_text(driver, By.ID, "congress_district"),
            "SD":         safe_get_text(driver, By.ID, "senate_district"),
            "Council D":  safe_get_text(driver, By.ID, "council_district"),
            "JD":         safe_get_text(driver, By.ID, "judicial_district"),
        }

        if all(not v for v in site_data.values()):
            print("   ❌ No polling information returned — possibly invalid address.")
            invalid_rows.append({
                "row": index,
                "house": house_number,
                "street": street_name,
                "zip": zip_code,
                "reason": "No polling info returned"
            })
            continue

        # Update output columns
        all_match = True
        for field, site_val in site_data.items():
            excel_val = safe_text(row.get(field, "")).strip()
            if excel_val != site_val:
                df.at[index, field] = site_val
                print(f"   Updated {field}: {excel_val or '∅'} → {site_val}")
                all_match = False

        if all_match:
            print("   ✅ All fields correct.")

    except Exception as e:
        print(f"[Row {index}] ❌ Error: {e}")
        invalid_rows.append({
            "row": index,
            "house": house_number,
            "street": street_name,
            "zip": zip_code,
            "reason": f"Script error: {e}"
        })
        continue

# =========================
# Save output
# =========================
base_name = os.path.splitext(file_path)[0]
output_path = f"{base_name}_with_districts.xlsx"
i = 1
while os.path.exists(output_path):
    output_path = f"{base_name}_with_districts({i}).xlsx"
    i += 1

df.to_excel(output_path, index=False)
print(f"\n✅ Done! File saved to: {output_path}")

# =========================
# Summary of invalid rows
# =========================
if invalid_rows:
    print("\n⚠️ Addresses that could not be processed:")
    for r in invalid_rows:
        print(f"   [Row {r['row']}] {r['house']} {r['street']}, {r['zip']} — {r['reason']}")

driver.quit()
