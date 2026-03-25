from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import re
import time
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from scripts.utils import make_driver

# ── Config ─────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent.parent
LOGINS_FILE = BASE_DIR / "data" / "directmarketlogins.xlsx"
OUTPUT_FILE = BASE_DIR / f"data/allocation_{int(time.time())}.xlsx"

HEADLESS = True
MAX_WORKERS = 3
STAGGER_SEC = 2
LOGIN_WAIT = 3
# ──────────────────────────────────────────────────────────────────────


# 🔥 WAIT FUNCTION (FIXED)
def wait_for_items(driver, timeout=25):
    WebDriverWait(driver, timeout).until(
        lambda d: len(find_all_elements_in_frames(
            d, By.CLASS_NAME, "catalauge-item-holder"
        )) > 0
    )


def find_element_in_frames(driver, by, selector):
    try:
        els = driver.find_elements(by, selector)
        if els:
            return els[0]
    except:
        pass

    for frame in driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe"):
        try:
            driver.switch_to.frame(frame)
            result = find_element_in_frames(driver, by, selector)
            if result:
                return result
            driver.switch_to.parent_frame()
        except:
            driver.switch_to.default_content()
    return None


def find_all_elements_in_frames(driver, by, selector):
    driver.switch_to.default_content()

    elements = driver.find_elements(by, selector)
    if elements:
        return elements

    for frame in driver.find_elements(By.TAG_NAME, "iframe") + driver.find_elements(By.TAG_NAME, "frame"):
        try:
            driver.switch_to.frame(frame)
            elements = driver.find_elements(by, selector)
            if elements:
                return elements

            for inner in driver.find_elements(By.TAG_NAME, "iframe") + driver.find_elements(By.TAG_NAME, "frame"):
                try:
                    driver.switch_to.frame(inner)
                    elements = driver.find_elements(by, selector)
                    if elements:
                        return elements
                    driver.switch_to.parent_frame()
                except:
                    pass

            driver.switch_to.parent_frame()
        except:
            pass

    driver.switch_to.default_content()
    return []


def parse_items(all_items, market):
    parsed = []

    for item in all_items:
        try:
            name = item.find_element(By.CLASS_NAME, "cat-prd-dsc").text.strip()
        except:
            name = "N/A"

        try:
            sku = item.find_element(By.CLASS_NAME, "cat-prd-id").text.strip()
        except:
            sku = "N/A"

        try:
            alloc_text = item.find_element(By.CLASS_NAME, "cat-prd-qty").text.strip()
            nums = re.findall(r'\d+', alloc_text)
            available = int(nums[0]) if len(nums) > 0 else 0
            total = int(nums[1]) if len(nums) > 1 else 0
        except:
            available, total = 0, 0

        if available > 0:
            parsed.append({
                "Market": market,
                "SKU": sku,
                "Name": name,
                "Available": available,
                "Total": total,
            })

    return parsed


def scrape_market(market, username, password):
    driver = make_driver(headless=HEADLESS)
    wait = WebDriverWait(driver, 25)

    try:
        driver.get("https://www.t-mobiledealerordering.com/")

        wait.until(EC.element_to_be_clickable((By.ID, "userid"))).send_keys(username)
        wait.until(EC.element_to_be_clickable((By.ID, "password"))).send_keys(password)

        wait.until(EC.element_to_be_clickable((By.NAME, "AgreeTerms"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@name='login']"))).click()

        time.sleep(LOGIN_WAIT)

        catalog_btn = find_element_in_frames(driver, By.XPATH, '//a[@onclick="show_catalog_view()"]')
        if not catalog_btn:
            return []

        catalog_btn.click()

        # 🔥 WAIT FIX
        wait_for_items(driver)

        devices = []

        # -------- PHONES --------
        driver.switch_to.default_content()

        all_items = find_all_elements_in_frames(
            driver, By.CLASS_NAME, "catalauge-item-holder"
        )

        print(f"[{market}] Phones:", len(all_items))
        devices.extend(parse_items(all_items, market))

        # -------- CPO --------
        driver.switch_to.default_content()

        cpo_link = find_element_in_frames(
            driver,
            By.XPATH,
            '//span[contains(text(),"CPO")]/ancestor::a'
        )

        if cpo_link:
            print(f"[{market}] → Opening CPO")

            cpo_link.click()
            wait_for_items(driver)

            driver.switch_to.default_content()

            cpo_items = find_all_elements_in_frames(
                driver, By.CLASS_NAME, "catalauge-item-holder"
            )

            print(f"[{market}] CPO:", len(cpo_items))
            devices.extend(parse_items(cpo_items, market))

        print(f"[{market}] ✅ {len(devices)} devices")
        return devices

    except Exception as e:
        print(f"[{market}] ❌ Error: {e}")
        return []

    finally:
        driver.quit()


def run(selected_markets=None):
    logins_df = pd.read_excel(LOGINS_FILE)
    logins_df.columns = logins_df.columns.str.strip()

    # 🔥 FILTER MARKETS
    if selected_markets:
        logins_df["Market"] = logins_df["Market"].astype(str).str.strip().str.lower()
        selected_markets = [m.strip().lower() for m in selected_markets]
        logins_df = logins_df[logins_df["Market"].isin(selected_markets)]

    rows = [row for _, row in logins_df.iterrows()]

    results_map = {}

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {
            executor.submit(
                scrape_market,
                str(row["Market"]).strip(),
                str(row["Username"]).strip(),
                str(row["Password"]).strip()
            ): row["Market"]
            for row in rows
        }

        for future in as_completed(futures):
            market = futures[future]
            try:
                results_map[market] = future.result()
            except:
                results_map[market] = []

    # 🔥 STYLED EXCEL
    wb = Workbook()
    ws = wb.active
    ws.title = "Allocation"

    header_fill = PatternFill("solid", fgColor="C40000")
    header_font = Font(bold=True, color="FFFFFF")
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    headers = ["Market", "SKU", "Name", "Available", "Total"]

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15

    row_num = 2
    for market, devices in results_map.items():
        for d in devices:
            ws.cell(row=row_num, column=1, value=d["Market"]).alignment = center
            ws.cell(row=row_num, column=2, value=d["SKU"]).alignment = center
            ws.cell(row=row_num, column=3, value=d["Name"]).alignment = left
            ws.cell(row=row_num, column=4, value=d["Available"]).alignment = center
            ws.cell(row=row_num, column=5, value=d["Total"]).alignment = center

            if row_num % 2 == 0:
                for col in range(1, 6):
                    ws.cell(row=row_num, column=col).fill = PatternFill("solid", fgColor="F2F2F2")

            row_num += 1

    wb.save(OUTPUT_FILE)

    print(f"Saved: {OUTPUT_FILE}")

    return str(OUTPUT_FILE)


if __name__ == "__main__":
    run()