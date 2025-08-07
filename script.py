import pandas as pd
import time
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# === Step 1: Select Excel file ===
Tk().withdraw()
file_path = askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
if not file_path:
    print("No file selected. Exiting.")
    exit()

# === Step 2: Read the Excel sheet ===
try:
    df = pd.read_excel(file_path, dtype=str)  # Read all cells as strings
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()

# === Step 3: Set up Edge WebDriver ===
EDGE_DRIVER_PATH = "msedgedriver.exe"  # Adjust this path if needed
options = webdriver.EdgeOptions()
options.add_argument("window-size=1000,800")
driver = webdriver.Edge(service=Service(EDGE_DRIVER_PATH), options=options)
driver.set_window_position(100, 100)

# === Step 4: Define helper to safely extract text ===
def safe_get_text(by, value, fallback=""):
    try:
        elem = driver.find_element(by, value)
        return elem.text.strip() if elem and elem.text else fallback
    except Exception:
        return fallback

# Track invalid rows
invalid_rows = []

# === Step 5: Process each row and update values ===
for index, row in df.iterrows():
    try:
        house_number = str(row['HOUSE #']).strip()
        street_name = str(row['STREET NAME']).strip()
        zip_code = str(row['ZIP CODE']).strip()

        if not all([house_number, street_name, zip_code]):
            print(f"[Row {index}] Skipped — missing address data.")
            continue

        print(f"\n[Row {index}] Checking: {house_number} {street_name}, {zip_code}")

        # Load site fresh
        driver.get("about:blank")
        time.sleep(0.5)
        driver.get("https://findmypollsite.vote.nyc/")
        wait = WebDriverWait(driver, 10)
        time.sleep(1)

        # Fill out form
        wait.until(EC.presence_of_element_located((By.ID, "txtHouseNumber"))).send_keys(house_number)
        wait.until(EC.presence_of_element_located((By.ID, "txtStreetName"))).send_keys(street_name)
        wait.until(EC.presence_of_element_located((By.ID, "txtZipcode"))).send_keys(zip_code)

        time.sleep(0.5)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Find My Site']"))).click()

        time.sleep(2)

        # === Check for invalid address message ===
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

        # === Scrape data from results page ===
        site_data = {
            "AD": safe_get_text(By.ID, "assembly_district"),
            "ED": safe_get_text(By.ID, "election_district").split("/")[0],
            "Cong D": safe_get_text(By.ID, "congress_district"),
            "SD": safe_get_text(By.ID, "senate_district"),
            "Council D": safe_get_text(By.ID, "council_district"),
            "JD": safe_get_text(By.ID, "judicial_district")
        }

        # Check if all scraped fields are empty — treat as invalid
        if all(not val for val in site_data.values()):
            print("   ❌ No polling information returned — possibly invalid address.")
            invalid_rows.append({
                "row": index,
                "house": house_number,
                "street": street_name,
                "zip": zip_code,
                "reason": "No polling info returned"
            })
            continue

        # === Compare and update values ===
        all_match = True
        for field, site_val in site_data.items():
            excel_val = str(row[field]).zfill(len(site_val)) if pd.notna(row[field]) else ""
            if excel_val != site_val:
                df.at[index, field] = site_val
                print(f"   Updated {field}: {excel_val} → {site_val}")
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

# === Step 6: Save corrected Excel file with unique name ===
base_name = os.path.splitext(file_path)[0]
output_path = f"{base_name}_corrected.xlsx"

i = 1
while os.path.exists(output_path):
    output_path = f"{base_name}_corrected({i}).xlsx"
    i += 1

df.to_excel(output_path, index=False)
print(f"\n✅ Done! Corrected file saved to: {output_path}")

# === Step 7: Print summary of invalid rows ===
if invalid_rows:
    print("\n⚠️ The following addresses could not be processed:")
    for row in invalid_rows:
        print(f"   [Row {row['row']}] {row['house']} {row['street']}, {row['zip']} — {row['reason']}")

driver.quit()
