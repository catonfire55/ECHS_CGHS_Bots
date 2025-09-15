import os
import sys
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import requests
import json
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
from collections import defaultdict
from pywinauto.application import Application
import shutil
import re

def search_test_code_and_rate(driver, wait, row_num, sub_category):
    try:
        # Step 1: Click the '?' button next to test code
        qbtn_xpath = f'//*[@id="table_proc"]/tbody/tr[{row_num}]/td[1]/button'
        wait.until(EC.element_to_be_clickable((By.XPATH, qbtn_xpath))).click()
        time.sleep(1)

        # Step 2: Enter sub_category into popup search
        search_input = wait.until(EC.presence_of_element_located((By.ID, 'searchText')))
        search_input.clear()
        search_input.send_keys(sub_category)
        driver.find_element(By.ID, 'searchbtn').click()

        # Step 3: Read first result (code + rate)
        test_code = driver.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[1]/td[2]').text.strip()
        rate = driver.find_element(By.XPATH, '//*[@id="example"]/tbody/tr[1]/td[3]').text.strip()

        # Step 4: Close popup
        try:
            driver.find_element(By.XPATH, '//*[@id="myModal"]/div/div/div[1]/button').click()
        except: pass

        return test_code, rate

    except Exception as e:
        print(f"❌ Search failed for {sub_category}: {e}")
        return None, None

def get_local_file_path(admission_number, file_name):
    src_path = os.path.join(r"\\172.16.1.205\machine3\IPD DSC", file_name)
    dest_dir = fr"C:\Users\{os.getlogin()}\Documents\upload_temp"
    os.makedirs(dest_dir, exist_ok=True)
    dest_path = os.path.join(dest_dir, file_name)
    clean_adm = re.sub(r'\D', '', str(admission_number))

    print(f"📁 Trying to copy:\nFROM: {src_path}\nTO:   {dest_path}")

    try:
        shutil.copy(src_path, dest_path)
        print(f"✅ Copied file locally: {file_name}")
        return dest_path
    except Exception as e:
        print(f"❌ Failed to copy file {file_name}: {e}")
        return None
def close_moa_popup_if_exists(driver):
    try:
        WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="ws_info_dialog"]'))
        )
        close_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="mainheader"]/div[9]/div[3]/div/button'))
        )
        driver.execute_script("arguments[0].click();", close_button)
        print("✅ MoA popup closed.")
        time.sleep(1)
    except Exception as e:
        print("ℹ️ MoA popup not found. Continuing...")

def build_override_map(excel_path):
    # Read from row 2 (index=1) because row 1 is merged header
    df = pd.read_excel(excel_path, sheet_name="Sheet1", header=1, dtype=str)
    df.columns = df.columns.str.strip()  # Clean column names

    # Ensure correct columns exist
    if not {'Wrong code', 'Right code', 'RATE'}.issubset(df.columns):
        raise Exception(f"❌ Columns missing. Found columns: {df.columns.tolist()}")

    df = df[['Wrong code', 'Right code', 'RATE']].dropna()

    override_map = {}
    for _, row in df.iterrows():
        wrong_codes = str(row['Wrong code']).split(',')
        right_code = str(row['Right code']).strip()
        rate = str(int(float(row['RATE'])))
        for code in wrong_codes:
            clean_code = code.strip().split("(")[0]
            override_map[clean_code] = {"correct_code": right_code, "rate": rate}
    return override_map


def override_test_codes(test_data_list, override_map):
    for test in test_data_list:
        raw_code = str(test.get("TEST_CODE", "")).strip()
        test_code = raw_code.upper().split("(")[0].strip()  # clean brackets and spaces

        if test_code in override_map:
            test['TEST_CODE'] = override_map[test_code]["correct_code"]
            test['AMOUNT'] = override_map[test_code]["rate"]
        else:
            test['TEST_CODE'] = raw_code
            test['AMOUNT'] = str(int(float(test.get("AMOUNT", "0"))))
    return test_data_list


def fetch_data_from_api(claim_no):
    url = "https://app.housepital.in/sarvodaya_api/api/echs_uploading_bot"
    headers = {
        "facilityGUID": "p09678-236-243244-afbb-5b87576da352",
        "userID": "bot@copperjam.com",
        "source": "Bot",
        "userKey": "iop90-789636-480b-34345-d5928b5bb7",
        "Content-Type": "application/json"
    }
    payload = {"data": {"CLAIM_NO": claim_no}}

    try:
        response = requests.post(url, headers=headers, json=payload, verify=False, timeout=10)
        if response.status_code == 200:
            data = response.json()
            records = data.get("result", []) 
            if records:
                first = records[0]
                return {
                    "date": first["DISCHARGE_ONLY_DATE"],
                    "time": first["DISCHARGE_ONLY_TIME"],
                    "type": first["DISCHARGE_TYPE"],
                    "doctor": first.get("PRIMARY_DOCTOR") ,
                    "net_amount": first.get("NET_AMOUNT", "0.00"),
                    "bill_no" :first.get("BILL_NO"),
                    "tests": records
                }
        else:
            print(f"❌ API call failed for {claim_no}, Status: {response.status_code}")
    except Exception as e:
        print(f"⚠️ Error for {claim_no}: {e}")
    return None

from selenium.webdriver.common.keys import Keys
from collections import defaultdict

def wait_until_overlay_disappears(driver, timeout=0.1):
    try:
        WebDriverWait(driver, timeout).until_not(
            EC.presence_of_element_located((By.CLASS_NAME, "ui-widget-overlay"))
        )
        print("✅ Overlay cleared.")
    except:
        print("⚠️ Overlay may still be present.")

def get_existing_test_rows(driver):
    existing_tests = {}
    rows = driver.find_elements(By.XPATH, '//*[@id="table_proc"]/tbody/tr')

    for index, row in enumerate(rows, start=2):  # Start from 2, as per your table logic
        try:
            code_inputs = row.find_elements(By.XPATH, './td[1]/input')
            if len(code_inputs) > 1:
                value = code_inputs[1].get_attribute("value").strip().upper()
                if value:
                    existing_tests[value] = index
        except Exception as e:
            print(f"⚠️ Skipped row {index}: {e}")
            continue

    print("📋 Existing Test Rows:", existing_tests)
    return existing_tests


def fill_expense_tests(driver, test_data_list):
    try:

        wait = WebDriverWait(driver, 30)

        added_keys = set()
        unlisted_tests = []
        final_unlisted_tests = []
        unlisted_tests_rows = 2
        final_tests = []
        final_packages = [] # {"TEST_CODE": "", "QUANTITY": "", "ROW": ""}
        roomrowscount = 3
        pkgtotalrows = 4
        testrows = 2
        medication_total = 0
        consultationdaystotal = 0
        room_categories = ["room rent"] 
        consultation_keywords = {"morning doctor visit", "evening doctor visit", "cross consultation charges", "doctor visit charges"}
        code_map = {
            1825: 1781,
            1826: 1782,
            1827: 1783,
            1829: 1784, 
            1824: 1785,
            1821: 1786,
            1822: 1787,
            1823: 1788,
            1781: 1789,
            1782: 1790,
            1783: 1791,
            1784: 1792,
            1785: 1793,
            1786: 1794,
            1787: 1795,
            1788: 1796,
            1789: 1797,
            1790: 1798,
            1792: 1799,
            1791: 1800,
            1793: 1801,
            1796: 1802,
            1797: 1803,
            1798: 1804,
            1799: 1805,
            1800: 1806,
            1801: 1807,
            1802: 1808,
            1803: 1809,
            1804: 1810,
            1805: 1811,
            1806: 1812,
            1807: 1813,
            1808: 1814,
            1809: 1815,
            1810: 1816,
            1811: 1817,
            1812: 1818,
            1813: 1819,
            1814: 1820,
            1815: 1821,
            1816: 1822,
            1817: 1823,
            1818: 1824,
            1819: 1825,
            1820: 1826,
        }


        def is_unlisted_code(test_code):
            forced_unlisted = {"1826"}
            if not test_code and txn_sub_category.strip().lower() not in room_categories and txn_sub_category.strip().lower() not in consultation_keywords:
                return True
            if (txn_sub_category.strip().lower() not in room_categories and txn_sub_category.strip().lower() not in consultation_keywords):
                test_code = str(test_code).strip()
            if not test_code.isdigit() and txn_sub_category.strip().lower() not in room_categories and txn_sub_category.strip().lower() not in consultation_keywords:
                return True
            if test_code in forced_unlisted:
                return True
            return False

        for test in test_data_list:
            
            test_code = str(test.get("TEST_CODE", "") or "").strip().upper()

            if (test_code in code_map):
                print(f"### Wrong Code {test_code} found replacing with right code {code_map[test_code]}")
                test_code = code_map[test_code]

            txn_sub_category = str(test.get("TXN_SUB_CATEGORY", "") or "").strip().lower()
            txn_category = str(test.get("TXN_CATEGORY", "") or "").strip().lower()
            quantity = int(float(test.get("QUANTITY", "1")))
            amount = int(float(test.get("MRP", "0")))  # Unit rate

            # ✅ FOR IP PACKAGES
            if ("ip package" in txn_category):
                pack_found = False
                for packs in final_packages:
                    if packs["TEST_CODE"] == test_code:
                        packs["QUANTITY"] += quantity
                        foundrow = packs["ROW"]
                        days_path = f'//*[@id="table_major"]/tbody/tr[{foundrow}]/td[3]/input[9]'
                        code_path = f'//*[@id="table_major"]/tbody/tr[{foundrow}]/td[1]/input'
                        mrp_path = f'//*[@id="table_major"]/tbody/tr[{foundrow}]/td[5]/input'
                        days_input = driver.find_element(By.XPATH, days_path)
                        days_input.clear()
                        time.sleep(0.2)
                        days_input.send_keys(packs["QUANTITY"])
                        print(f"✅ Found Existing Package again, changing quantity at row={foundrow} ...")
                        pack_found = True
                        continue

                if not pack_found:
                    final_packages.append({
                        "TEST_CODE" : test_code,
                        "QUANTITY": quantity,
                        "ROW" : pkgtotalrows
                    })
                    days_path = f'//*[@id="table_major"]/tbody/tr[{pkgtotalrows}]/td[3]/input[9]'
                    code_path = f'//*[@id="table_major"]/tbody/tr[{pkgtotalrows}]/td[1]/input'
                    mrp_path = f'//*[@id="table_major"]/tbody/tr[{pkgtotalrows}]/td[5]/input'
                    code_input = driver.find_element(By.XPATH, code_path)
                    time.sleep(0.1)
                    code_input.clear()
                    time.sleep(0.2)
                    code_input.send_keys(test_code)
                    time.sleep(0.2)
                    days_input = driver.find_element(By.XPATH, days_path)
                    time.sleep(0.1)
                    days_input.clear()

                    days_input.send_keys(quantity)
                    time.sleep(0.2)
                    mrp_input = driver.find_element(By.XPATH, mrp_path)
                    time.sleep(0.1)
                    mrp_input.clear()
                    time.sleep(0.2)
                    mrp_input.send_keys(amount)
                    time.sleep(0.1)
                    pkgrow_add_btn = driver.find_element(By.XPATH, '//*[@id="table_major"]/tbody/tr[3]/td[2]/input')
                    pkgrow_add_btn.click()
                    print(f"✅ Found New Package, Added to final_packages, and added to sheet row={pkgtotalrows-3} ...")
                    pkgtotalrows+=1
                    time.sleep(0.5)
                    continue


            # ✅ FILL ROOM CHARGES → SEPARATELY !!
            if (txn_sub_category.strip().lower() in room_categories):
                try:
                    if (roomrowscount > 3):
                        driver.find_element(By.XPATH, room_add_btn).click()
                        
                    room_type_drop_path = f'/html/body/div[13]/div[2]/table[2]/tbody/tr[{roomrowscount}]/td[1]/select'
                    if (amount <= 1500):
                        Select(driver.find_element(By.XPATH, room_type_drop_path)).select_by_visible_text("General")
                        print("🏠 Selected General Room Type...")
                    if (amount > 1500 and amount <=3000):
                        Select(driver.find_element(By.XPATH, room_type_drop_path)).select_by_visible_text("Semi-Private")
                        print("🏠 Selected Semi Private Room Type...")
                    if (amount > 3000 and amount<=4500):
                        Select(driver.find_element(By.XPATH, room_type_drop_path)).select_by_visible_text("Private")
                        print("🏠 Selected Private Room Type...")
                    if (amount > 4500):
                        Select(driver.find_element(By.XPATH, room_type_drop_path)).select_by_visible_text("Intensive Care Unit (ICU)")
                        print("🏠 Selected ICU/MICU Room Type...")
                    
                    room_input_xpath = f'//*[@id="table_room"]/tbody/tr[{roomrowscount}]/td[2]/input'  # Assumed XPath for room charges
                    room_mrp_xpath = f'//*[@id="table_room"]/tbody/tr[{roomrowscount}]/td[4]/input'
                    room_add_btn = '//*[@id="table_room"]/tbody/tr[2]/td[2]/input'
                    room_input = driver.find_element(By.XPATH, room_input_xpath)
                    room_mrp = driver.find_element(By.XPATH, room_mrp_xpath)
                    room_input.clear()
                    time.sleep(0.1)
                    room_input.send_keys(quantity)
                    time.sleep(0.1)
                    time.sleep(0.2)
                    room_mrp.clear()
                    time.sleep(0.1)
                    amount = str(int(float(test.get("MRP", "0"))))
                    room_mrp.send_keys(amount)
                    print(f"🏠 Filled Room Charges: {quantity} for ₹{amount} in Room Data")
                    roomrowscount+=1
                    continue
                except Exception as e:
                    print(f"❌ Failed to fill Room Data in Room Charges: {e}")
                    continue
            
            if txn_sub_category.strip().lower() in consultation_keywords:
                try:
                    consultationdaysxpath = '//*[@id="table_consult"]/tbody/tr[2]/td[2]/input[12]'
                    consultationcharges = '//*[@id="table_consult"]/tbody/tr[2]/td[4]/input'
                    consultationdaystotal += quantity
                    consultationdaysinput = driver.find_element(By.XPATH, consultationdaysxpath)
                    consultationdaysinput.clear()
                    time.sleep(0.1)
                    consultationdaysinput.send_keys(consultationdaystotal)
                    time.sleep(0.1)
                    consultationchargesinput = driver.find_element(By.XPATH, consultationcharges)
                    consultationchargesinput.clear()
                    time.sleep(0.1)
                    consultationchargesinput.send_keys(amount)
                    time.sleep(0.1)
                    continue

                except:
                    print("❌ Failed To Process Consultation Charges...\n")
                    continue

            # ✅ SKIP MEDICATION → total them

            if "medicine" in txn_sub_category.strip().lower() or "medication" in txn_category.strip().lower():
                try:
                    medication_total += int(float(amount)) * quantity
                    print(f"💊 Skipping medication: {txn_sub_category} → ₹{amount} x {quantity}")
                except Exception as e:
                    print(f"⚠️ Error calculating medication total for {txn_sub_category}: {e}")
                continue

            if test_code == "1812 & 1638":
                split_rates = {"1818": "6000", "1638": "1553"}
                for split_code in ["1818", "1638"]:
                    key = (split_code, txn_sub_category)
                    if key not in added_keys:
                        added_keys.add(key)
                        final_tests.append({
                            "TEST_CODE": split_code,
                            "TXN_SUB_CATEGORY": txn_sub_category,
                            "QUANTITY": quantity,
                            "AMOUNT": split_rates[split_code]
                        })
                    else:
                        for entry in final_tests:
                            if entry["TEST_CODE"] == split_code:
                                entry["QUANTITY"] += quantity
                                break
                continue

            key = (test_code, txn_sub_category)

            # ✅ UNLISTED CODES (after skipping room/med)
            if is_unlisted_code(test_code):
                unlisted_tests.append((txn_sub_category, quantity, amount))
                continue

            # ✅ DUPLICATE CHECK && TESTS ADDING...
            found = False
            for entry in final_tests:
                if (txn_sub_category.strip().lower() not in room_categories and txn_sub_category.strip().lower() not in consultation_keywords and "ip package" not in txn_category):
                    if entry["TEST_CODE"] == test_code:
                        entry["QUANTITY"] += quantity
                        row = entry["ROW"]
                        quantity_xpath = f'//*[@id="table_proc"]/tbody/tr[{row}]/td[3]/input[9]'
                        quantity_input = driver.find_element(By.XPATH, quantity_xpath)
                        time.sleep(0.1)
                        quantity_input.clear()
                        quantity_input.send_keys(entry["QUANTITY"])
                        found = True
                        print(f"✅ Found Test Code Pre Existing... Updated Values... at row {row}")
                        time.sleep(0.2)
                        continue
                        

            if not found and txn_sub_category.strip().lower() not in room_categories and txn_sub_category.strip().lower() not in consultation_keywords and "ip package" not in txn_category:
                added_keys.add(key)
                final_tests.append({
                    "TEST_CODE": test_code,
                    "QUANTITY": quantity,
                    "AMOUNT": amount,
                    "ROW":testrows
                })
                print(f"✅ Found New Test Code {test_code}, Added To List...")
                code_path = f'//*[@id="table_proc"]/tbody/tr[{testrows}]/td[1]/input[2]'
                quantity_xpath = f'//*[@id="table_proc"]/tbody/tr[{testrows}]/td[3]/input[9]'
                mrp_path = f'//*[@id="table_proc"]/tbody/tr[{testrows}]/td[5]/input'
                code_input = driver.find_element(By.XPATH, code_path)
                time.sleep(0.1)
                code_input.clear()
                time.sleep(0.2)
                code_input.send_keys(test_code)
                time.sleep(0.2)
                quantity_input = driver.find_element(By.XPATH, quantity_xpath)
                time.sleep(0.1)
                quantity_input.clear()
                quantity_input.send_keys(quantity)
                time.sleep(0.2)
                try:
                    mrp_input = driver.find_element(By.XPATH, mrp_path)
                    time.sleep(0.1)
                    mrp_input.clear()
                    time.sleep(0.2)
                    mrp_input.send_keys(amount)
                    time.sleep(0.1)
                except:
                    print(f"❌ Skipping Test {test_code}")
                row_add_btn_path = '//*[@id="table_proc"]/tbody/tr[1]/td[2]/input'
                row_add_btn = driver.find_element(By.XPATH, row_add_btn_path)
                row_add_btn.click()
                testrows+=1
                print(f"✅ Found New Test Code:{test_code} filled at row={testrows-2}")
                time.sleep(0.5)
                continue
            

        
        # Unlisted tests - excluding medication, consultation, and now room entries (already skipped)
        for txn_sub_category, quantity, amount in unlisted_tests:
            if "medicine" in txn_sub_category.strip().lower() or "medication" in txn_category.strip().lower() or txn_sub_category.strip().lower() in consultation_keywords and txn_sub_category.strip().lower() not in room_categories:
                continue
            try:
                found_unlisted = False
                for entry in final_unlisted_tests:
                    if (entry["TXN_SUBCAT"] == txn_sub_category.strip().lower()):
                        found_unlisted = True
                        net_amt = quantity*amount
                        unlisted_row = entry["ROW"]
                        entry["QUANTITY"] += quantity
                        entry["NET_AMT"] += net_amt
                        net_amt_path = f'//*[@id="table_unlist"]/tbody/tr[{unlisted_row}]/td[2]/input'
                        net_amt_input = driver.find_element(By.XPATH, net_amt_path)
                        net_amt_input.clear()
                        time.sleep(0.1)
                        net_amt_input.send_keys(entry["NET_AMT"])
                        print(f"📝 Updated Existing Unlisted Code At Row - {unlisted_row}")
                        

                if (not found_unlisted):
                    net_amt = amount*quantity
                    final_unlisted_tests.append({
                        "TXN_SUBCAT" : txn_sub_category.strip().lower(),
                        "ROW" : unlisted_tests_rows,
                        "QUANTITY" : quantity,
                        "AMOUNT" : amount,
                        "NET_AMT" : net_amt
                    })
                    txn_subcat_path = f'//*[@id="table_unlist"]/tbody/tr[{unlisted_tests_rows}]/td[1]/textarea'
                    net_amt_path = f'//*[@id="table_unlist"]/tbody/tr[{unlisted_tests_rows}]/td[2]/input'
                    txn_subcat_input = driver.find_element(By.XPATH, txn_subcat_path)
                    txn_subcat_input.clear()
                    time.sleep(0.1)
                    txn_subcat_input.send_keys(txn_sub_category)
                    time.sleep(0.1)
                    net_amt_input = driver.find_element(By.XPATH, net_amt_path)
                    net_amt_input.clear()
                    time.sleep(0.1)
                    net_amt_input.send_keys(net_amt)
                    time.sleep(0.1)
                    add_btn = driver.find_element(By.XPATH, '//*[@id="table_unlist"]/tbody/tr[1]/td[2]/input')
                    add_btn.click()
                    print(f"📝 Added Unlisted Code At Row - {unlisted_tests_rows-1}")
                    unlisted_tests_rows+=1
                    time.sleep(0.2)

                
            except Exception as ue:
                print(f"❌ Unlisted entry failed: {txn_sub_category} | {ue}")

        #Medication total filled in Other Charges
        if medication_total > 0:
            try:
                med_input_xpath = '//*[@id="table_other"]/tbody/tr[2]/td[2]/input'
                med_input = driver.find_element(By.XPATH, med_input_xpath)
                # (driver, med_input)
                med_input.clear()
                med_input.send_keys(str(medication_total))
                print(f"💊 Filled total Medication Charges: ₹{medication_total} in Pharmacy & Consumable Charges")
            except Exception as e:
                print(f"❌ Failed to fill Medication Total in Other Charges: {e}")

        consultation_tests = [
            (txn_sub_category, quantity, amount)
            for txn_sub_category, quantity, amount in unlisted_tests
            if txn_sub_category.strip().lower() in consultation_keywords
        ]

        for txn_sub_category, quantity, amount in consultation_tests:
            try:
                quantity_xpath = '//*[@id="table_cons"]/tbody/tr[1]/td[3]/input'
                pass
            except Exception as e:
                print(f"❌ Failed to add consultation {txn_sub_category}: {e}")
    except:
        print(f"❌ Failed to add test code {test_code}, CONTINUING...")



def round_minutes_to_5(minute_str):
    minute = int(minute_str)
    rounded = 5 * round(minute / 5)
    if rounded == 60:
        rounded = 0
    return str(rounded).zfill(2)
def clear_existing_test_rows(driver):
    print("🔄 Cleaning existing test rows from #table_proc...")

    while True:
        try:
            rows = driver.find_element(By.XPATH, '//*[@id="table_proc"]/tbody/tr')
            deleted_any = False

            for i in range(2, len(rows) + 2):
                try:
                    code_xpath = f'//*[@id="table_proc"]/tbody/tr[{i}]/td[1]/input[2]'
                    delete_btn_xpath = f'//*[@id="table_proc"]/tbody/tr[{i}]/td[8]/img'

                    code_input = driver.find_element(By.XPATH, code_xpath)
                    value = code_input.get_attribute("value").strip()
                    if value:
                        delete_btn = driver.find_element(By.XPATH, delete_btn_xpath)
                        # (driver, delete_btn)

                        try:
                            delete_btn.click()
                        except:
                            driver.execute_script("arguments[0].click();", delete_btn)

                        # 🔐 Accept popup
                        WebDriverWait(driver, 5).until(EC.alert_is_present())
                        alert = driver.switch_to.alert
                        print(f"⚠️ Alert: {alert.text}")
                        alert.accept()

                        print(f"🗑️ Deleted test row with code: {value}")
                        deleted_any = True
                        break  # break inner loop to re-fetch rows
                except Exception as e:
                    print(f"⚠️ Could not process row {i}: {e}")
                    continue

            if not deleted_any:
                break

        except Exception as e:
            print(f"❌ Error during test row cleanup: {e}")
            break

def clear_existing_unlisted_procedures(driver):
    wait = WebDriverWait(driver, 10)
    try:
        rows = driver.find_elements(By.XPATH, '//*[@id="table_unlist"]/tbody/tr')
        print(f"🧹 Found {len(rows)} unlisted rows")

        for i in range(len(rows), 1, -1):  # Reverse loop: last to 2
            try:
                textarea_xpath = f'//*[@id="table_unlist"]/tbody/tr[{i}]/td[1]/textarea'
                delete_btn_xpath = f'//*[@id="table_unlist"]/tbody/tr[{i}]/td[4]/img'

                # Check if textarea has any content
                textarea = driver.find_element(By.XPATH, textarea_xpath)
                if textarea.get_attribute("value").strip():
                    delete_btn = wait.until(EC.element_to_be_clickable((By.XPATH, delete_btn_xpath)))
                    # (driver, delete_btn)
                    delete_btn.click()

                    WebDriverWait(driver, 5).until(EC.alert_is_present())
                    driver.switch_to.alert.accept()

                    print(f"🗑️ Deleted unlisted row {i}")
                    time.sleep(0.1)
            except Exception as e:
                print(f"⚠️ Could not delete row {i}: {e}")

    except Exception as outer:
        print(f"❌ Outer error in clearing unlisted procedures: {outer}")



def get_chrome_driver_path():
    base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(__file__)
    return os.path.join(base_path, "chromedriver.exe")

def navigate_to_new_ip_submission(driver):
    wait = WebDriverWait(driver, 60)
    try:
        checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ihaveseennmi"]')))
        checkbox.click()
        print("✅ Checked the 'Need More Info' checkbox.")

        bill_submission_menu = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="li_1"]')))

        bill_submission_menu.click()
        print("✅ Clicked on 'Bill Submission'.")

        new_ip_submission = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="1"]/li[1]/a')))

        new_ip_submission.click()
        print("✅ Clicked on 'New IP Submission'.")

    except Exception as e:
        print(f"⚠️ Error in navigating to New IP Submission: {e}")
        # Function to search for admission number in File Explorer
def search_admission_number_in_explorer(admission_number):
       # Define the folder where the files are stored
    clean_adm = re.sub(r'\D', '', str(admission_number))
    network_path = r"\\172.16.1.205\machine3\IPD DSC"
    
    try:
        # List all files in the specified directory
        all_files = os.listdir(network_path)
        matched_files = [file for file in all_files if clean_adm in file]

        # Loop through the files and match the admission number with the filenames
        for file in all_files:
            if str(admission_number) in file:
                matched_files.append(file)

        # If files are found, print them
        if matched_files:
            print(f"✅ Found the following files for Admission Number {admission_number}:")
            for matched_file in matched_files:
                print(f"- {matched_file}")
        else:
            print(f"❌ No files found for Admission Number {admission_number} in the directory.")
        
    except Exception as e:
        print(f"❌ Error during file search: {e}")

# Example usage:

def fill_claim_id_and_submit(driver, claim_id, discharge_data,test_data_list):
    wait = WebDriverWait(driver, 5)
    try:
        # 🔁 Ensure we are on the right page by waiting for the Claim ID input field
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="claimId"]'))).send_keys(claim_id)

        print(f"✅ Entered Claim ID: {claim_id}")

        submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="submitbutton"]/td/input')))

        submit_btn.click()
        print("✅ Submitted claim ID form")

        # Click on the matching Claim ID in the table
        claim_link = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="claimtable"]/tbody/tr[2]/td[2]/a')))
        
        claim_link.click()
        print("✅ Clicked on claim ID to open details page")


        # Click on Admission Tab
        admission_tab = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mytabber"]/ul/li[2]/a')))
        admission_tab.click()
        print("✅ Clicked on 'Admission' tab")
        WebDriverWait(driver, 10)
       
        date = discharge_data['date']  # e.g., 2025-05-28
        time_parts = discharge_data['time'].split(":")  # e.g., ['14', '02', '49']
        hour_24 = time_parts[0].zfill(2)
        minutes = round_minutes_to_5(time_parts[1])

        driver.execute_script(f"document.getElementById('dateOfDischarge').value = '{date[8:10]}/{date[5:7]}/{date[:4]}'")
        driver.execute_script(f"document.getElementById('dateOfDischargeHours').value = '{hour_24}'")
        driver.execute_script(f"document.getElementById('dateOfDischargeMins').value = '{minutes}'")

        print(f"✅ Filled dynamic discharge: {date} {hour_24}:{minutes}")
        # --- Dynamic Discharge Type Mapping ---

        type_mapping = {
            "regular discharge": "R",
            "death discharge": "D",
            "lama": "L",
            "transfer to other hospital": "T"
        }
        try:
            discharge_type_text = discharge_data.get("type", "").strip().lower()
            discharge_code = type_mapping.get(discharge_type_text, "R")  # Default to Normal
        except:
            discharge_code = "N"
    

        # --- Select Discharge Type via JS ---
        discharge_dropdown = driver.find_element(By.XPATH, '//*[@id="dischargetype"]')
        #  (driver, discharge_dropdown)
        #  (driver, discharge_dropdown)
        driver.execute_script(f"document.getElementById('dischargetype').value = '{discharge_code}'")
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'))", discharge_dropdown)
        print(f"✅ Selected Discharge Type: → {discharge_code}")


        # Fill treating doctor name
        treating_doctor = discharge_data.get("doctor", "").strip()
        doctor_input = driver.find_element(By.XPATH, '//*[@id="treatingDoctors"]')
        #  (driver, doctor_input)
        #  (driver, doctor_input)
        doctor_input.clear()
        doctor_input.send_keys(treating_doctor)
        print(f"✅ Filled Treating Doctor: {treating_doctor}")





        # Fill net amount
        # net_amt = discharge_data.get("net_amount", "0.00")
        raw_net_amt = discharge_data.get("net_amount", "0")
        try:
            net_amt = str(int(float(raw_net_amt)))  # Convert to int first to remove decimal
        except (ValueError, TypeError):
            print(f"⚠️ Invalid net amount from API: {raw_net_amt}, defaulting to 0")
            net_amt = "0"

        # 🧾 Fill Net Amount field
        try:
            net_amt_input = driver.find_element(By.XPATH, '//*[@id="confirmNetClaimAmt"]')
            #  (driver, net_amt_input)
            #  (driver, net_amt_input)
            net_amt_input.clear()
            print(f"✅ Filling Net Amount: {net_amt}")
            net_amt_input.send_keys(net_amt)
            print(f"✅ Filled Net Amount: ₹{net_amt}")
        except Exception as e:
            print(f"❌ Failed to fill Net Amount: {e}")
        # Copy ailment and paste into hospital remark
        # Copy ailment and paste into Diagnosis field (not Hospital Remark)
        try:
            ailment_text = driver.find_element(By.XPATH, '//*[@id="admissionAilment"]').get_attribute("value")
            diagnosis_box = driver.find_element(By.XPATH, '//*[@id="diagnosis"]')
            diagnosis_box.clear()
            diagnosis_box.send_keys(ailment_text)
            print("✅ Copied admission ailment into Diagnosis field")
        except Exception as e:
            print(f"❌ Failed to copy ailment to Diagnosis field: {e}")


        # Click on next tab
        next_tab = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mytabber"]/ul/li[3]/a')))
        #  (driver, next_tab)
        next_tab.click()
        print("✅ Moved to 'Bill Details' tab") 
        # Fill Bill Number
        # 📦 Step: Get bill number from discharge_data
        bill_number = discharge_data.get("bill_no").strip()

        # 🧾 Step: Fill in the form
        try:
            bill_no_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="billNo"]')))
            #  (driver, bill_no_input)
            #  (driver, bill_no_input)
            bill_no_input.clear()
            bill_no_input.send_keys(bill_number)
            print(f"✅ Filled Bill Number: {bill_number}")
        except Exception as e:
            print(f"❌ Failed to fill Bill Number: {e}")
        # Click on the Expense link (worksheet)
        worksheet_link = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="a_worksheet"]')))
        #  (driver, worksheet_link)
        #  (driver, worksheet_link)
        worksheet_link.click()
        print("✅ Clicked on 'Expense' worksheet hyperlink")
        # Allow worksheet to load
        # 👇 Inspect API data
        # 🧹 Clear any pre-filled unlisted procedures
        # 🧹 Clear test rows.
        time.sleep(0.1)

        clear_existing_test_rows(driver)
        clear_existing_unlisted_procedures(driver)
        print(json.dumps(test_data_list[0], indent=2))
        fill_expense_tests(driver, test_data_list)
        # ✅ Expense worksheet close karo
        try:
            close_expense_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mainheader"]/div[13]/div[3]/div/button')))
            #  (driver, close_expense_btn)
            #  (driver, close_expense_btn)
            close_expense_btn.click()
            print("✅ Closed Expense popup")
        except Exception as e:
            print(f"⚠️ Failed to close Expense popup: {e}")

        # ✅ Line of Treatment tab 
        try:
            line_of_treatment_tab = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mytabber"]/ul/li[4]/a')))
            #  (driver, line_of_treatment_tab)
            #  (driver, line_of_treatment_tab)
            line_of_treatment_tab.click()
            print("✅ Opened Line of Treatment tab")
        except Exception as e:
            print(f"⚠️ Could not open Line of Treatment tab: {e}")

        # ✅ Investigation mein 'Yes' select karo aur "Investigations" likho
      # ✅ Click Line of Treatment tab
        try:
            line_of_treatment_tab = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mytabber"]/ul/li[4]/a')))
            #  (driver, line_of_treatment_tab)
            #  (driver, line_of_treatment_tab)
            line_of_treatment_tab.click()
            print("✅ Opened Line of Treatment tab")
        except Exception as e:
            print(f"⚠️ Could not open Line of Treatment tab: {e}")

        # ✅ Select 'Yes' for Investigations
        try:
            yes_radio = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//input[@name="proposedtreatment[0].anyTreatment" and @value="Y"]')
            ))
            #  (driver, yes_radio)
            #  (driver, yes_radio)
            yes_radio.click()
            print("✅ Clicked 'Yes' for Investigations")

            # ✅ Fill textarea
            textarea = wait.until(EC.element_to_be_clickable((By.ID, "treatmentDetail")))
            #  (driver, textarea)
            #  (driver, textarea)
            textarea.clear()
            textarea.send_keys("Investigations")
            print("📝 Filled 'Investigations' in textarea")

        except Exception as e:
            print(f"❌ Failed to handle Line of Treatment: {e}")
        # ✅ Click Clinical Finding Tab
                # ✅ Click on 'Clinical Finding' tab
        try:
            # Step 1: Click Clinical Finding tab
            clinical_tab = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mytabber"]/ul/li[5]/a')))
            #  (driver, clinical_tab)
            driver.execute_script("arguments[0].click();", clinical_tab)
            print("✅ Clicked on 'Clinical Finding' tab")

            # Step 2: Click Yes radio using JS (more reliable)
            try:
                yes_radio_xpath = "//input[@name='clinicalfinding[6].anyTest' and @value='Y']"
                text_input_xpath = "(//input[@id='testDetail'])[last()]"

                yes_radio = wait.until(EC.element_to_be_clickable((By.XPATH, yes_radio_xpath)))
                #  (driver, yes_radio)
                yes_radio.click()
                print("✅ Selected 'Yes' for Clinical Finding → Other")

                # Wait for correct textbox
                WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, text_input_xpath)))
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, text_input_xpath)))

                input_field = driver.find_element(By.XPATH, text_input_xpath)
                #  (driver, input_field)

                # Force fill with JS
                driver.execute_script("""
                    arguments[0].style.display = 'inline';
                    arguments[0].style.visibility = 'visible';
                    arguments[0].value = 'As Per Requirments';
                    arguments[0].dispatchEvent(new Event('input'));
                    arguments[0].dispatchEvent(new Event('blur'));
                """, input_field)

                print("📝 Filled correct textbox for Clinical Finding → Other")

            except Exception as e:
                print(f"❌ Final fix failed for Clinical Finding → Other: {e}")

            


        except Exception as e:
            print(f"❌ Final fallback failed to fill Clinical Finding → Other: {e}")

        # try:
        #     upload_tab = WebDriverWait(driver, 10).until(
        #         EC.element_to_be_clickable((By.XPATH, '//*[@id="mytabber"]/ul/li[6]/a'))
        #     )
        #      (driver, upload_tab)
        #      (driver, upload_tab)
        #     upload_tab.click()
        #     print("✅ Opened 'Upload Documents' tab")
        #     time.sleep(2)
        # except Exception as e:
        #     print(f"❌ Failed to open 'Upload Documents' tab: {e}")
        # Step 1: Navigate to the Admission Tab
        
        admission_tab_xpath = '//*[@id="mytabber"]/ul/li[7]/a'  # XPath for the Admission tab
        admission_tab = wait.until(EC.element_to_be_clickable((By.XPATH, admission_tab_xpath)))
        #  (driver, admission_tab)
        #  (driver, admission_tab)
        admission_tab.click()
        print("✅ Clicked on 'Admission' tab")

        # Step 2: Extract the Admission Number
        admission_number_xpath = '//*[@id="name"]/table/tbody/tr[1]/td/table[2]/tbody/tr[4]/td[2]'
        admission_number = driver.find_element(By.XPATH, admission_number_xpath).text
        print(f"✅ Extracted Admission Number: {admission_number}")
        

                 
       # Step 3: Navigate to the Upload Documents Tab
        upload_documents_tab_xpath = '//*[@id="mytabber"]/ul/li[6]/a'  # XPath for the Upload Documents tab
        upload_documents_tab = wait.until(EC.element_to_be_clickable((By.XPATH, upload_documents_tab_xpath)))
        #  (driver, upload_documents_tab)
        #  (driver, upload_documents_tab)
        upload_documents_tab.click()
        print("✅ Opened 'Upload Documents' tab")
        upload_documents_for_admission(driver, admission_number)
        # Access the network share using the address
        # network_path = r"\\172.16.1.205\machine3\IPD DSC"
        # os.startfile(network_path)  # This will open the network share in File Explorer
        search_admission_number_in_explorer(admission_number)
        time.sleep(0.1)  # Wait for the folder to open


      




    except Exception as e:
        print(f"⚠️ Error during claim processing: {e}")
                # 🚪 Close the Expense popup
def upload_documents_for_admission(driver, admission_number):
    wait = WebDriverWait(driver, 20)
    upload_path = r"\\172.16.1.205\\machine3\\IPD DSC"
    local_temp_dir = fr"C:\\Users\{os.getlogin()}\\Documents\\upload_temp"
    os.makedirs(local_temp_dir, exist_ok=True)

    # Clean admission number
    clean_adm = re.sub(r'\D', '', str(admission_number))

    try:
        # Step 1: Match files
        files = os.listdir(upload_path)
        matched_files = [f for f in files if clean_adm in f.lower()]
        if not matched_files:
            print(f"❌ No files found for Admission Number {admission_number}")
            return

        print(f"✅ Found files: {matched_files}")

        # Step 2: Copy locally
        local_file_map = {}
        for file_name in matched_files:
            src = os.path.join(upload_path, file_name)
            dest = os.path.join(local_temp_dir, file_name)
            try:
                shutil.copy2(src, dest)
                local_file_map[file_name] = dest
                print(f"📥 Copied: {file_name}")
            except Exception as e:
                print(f"❌ Failed to copy {file_name}: {e}")

        # Step 3: Define keywords
        document_keywords = {
            "Bill Details": ["bill"],
            "Death Certificate": ["death"],
            "Delay Condonation Letter": ["waiver"],
            "Discharge Summary": ["disc"],
            "Drug Certificate": ["drug"],
            "Lab Reports": ["rep"],
            "Unlisted Procedure Approval": ["appx"],
            "Feedback Form": ["feed"],
            "Others": ["pouch", "cardpre", "pro", "chemo", "ot", "other", "dispho", "haemo", "repo", "prbc", "aadharcard", "pre", "procedure", "refferal", "transfer"],
            "Extended Stay Certificate": ["extention"]
        }

        # Step 4: Read already uploaded filenames
        uploaded_filenames = set()
        try:
            rows = driver.find_elements(By.XPATH, '//*[@id="uploadtable"]/tbody/tr')
            for row in rows:
                try:
                    fname = row.find_element(By.XPATH, './td[2]').text.strip()
                    uploaded_filenames.add(fname)
                except:
                    continue
        except:
            pass

        # Step 5: Match and upload
        dropdown = wait.until(EC.presence_of_element_located((By.ID, 'uploadDocType')))
        options = dropdown.find_elements(By.TAG_NAME, 'option')

        for option in options:
            doc_type = option.text.strip()
            if doc_type not in document_keywords:
                continue

            keywords = document_keywords[doc_type]
            # matching_files = [f for f in local_file_map if any(k in f.lower() for k in keywords) and f not in uploaded_filenames]
            matching_files = [
    f for f in local_file_map
    if any(k in f.lower() for k in keywords)
    and f not in uploaded_filenames
    and "ds" in f.lower()  # 🔒 Only allow files with 'DS' in the name
]
            if not matching_files:
                print(f"⏭️ No files matched for: {doc_type}")
                continue

            for file_name in matching_files:
                try:
                    # Re-select dropdown every time
                    dropdown = wait.until(EC.element_to_be_clickable((By.ID, 'uploadDocType')))
                    for opt in dropdown.find_elements(By.TAG_NAME, 'option'):
                        if opt.text.strip().lower() == doc_type.lower():
                            opt.click()
                            print(f"📄 Selected: {doc_type}")
                            break

                    # Send file path directly to <input type="file">
                    file_input = wait.until(EC.presence_of_element_located((By.ID, 'medicalReport')))
                    file_path = local_file_map[file_name]
                    driver.execute_script("arguments[0].style.display = 'block';", file_input)

                    file_input.send_keys(file_path)
                    print(f"📎 Selected file: {file_name}")

                    # upload_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//input[@value="Upload"]')))
                    #  (driver, upload_btn)
                    # upload_btn.click()
                    try:
                        WebDriverWait(driver, 10).until_not(
                            EC.presence_of_element_located((By.CLASS_NAME, 'ui-widget-overlay'))
                        )
                        print("✅ Overlay cleared before upload.")
                    except:
                        print("⚠️ Overlay might still be present.")

                    # Proceed to upload
                    upload_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//input[@value="Upload"]')))
                    # (driver, upload_btn)
                    upload_btn.click()
                    print(f"✅ Uploaded: {file_name} as {doc_type}")

                except Exception as e:
                    print(f"❌ Failed to upload {file_name} under {doc_type}: {e}")
                    continue
                try:
                    for file in os.listdir(local_temp_dir):
                        file_path = os.path.join(local_temp_dir, file)
                        if os.path.isfile(file_path):
                            os.remove(file_path)
                    print("🧹 Cleaned up local temp folder after upload.")
                except Exception as e:
                    print(f"⚠️ Failed to clean temp folder: {e}")

    except Exception as e:
        print(f"❌ Outer error: {e}")



      
def main():

    claim_ids = []
    claimid = str(input("####### ENTER THE CLAIM ID: "))
    claim_ids.append(claimid)

    options = Options()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    service = Service(get_chrome_driver_path())
    driver = webdriver.Chrome()

    driver.get("https://www.echsbpa.utiitsl.com/ECHS")
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="username"]'))).send_keys("shrcechs")
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys("care@321")
    print("⏳ Please fill captcha manually within 30 sec...")

    try:
        WebDriverWait(driver, 120).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ihaveseennmi"]'))
        )
        close_moa_popup_if_exists(driver)
        print("✅ Login detected. Proceeding with navigation...")
        navigate_to_new_ip_submission(driver)
        
        for raw_claim_id in claim_ids:
            claim_id = str(int(float(raw_claim_id.strip())))
            print(f"🔁 Claim ID from Excel: '{claim_id}'")

            result = fetch_data_from_api(claim_id)
            if not result:
                print(f"❌ Skipping Claim ID {claim_id} as no discharge data found from API.")
                continue
            print("Fetched data from API...")
            discharge_data = {
                "date": result["date"],
                "time": result["time"],
                "type": result["type"],
                "doctor": result["doctor"],
                "net_amount": result["net_amount"],
                "bill_no": result["bill_no"]
            }
            test_data_list = result.get("tests", [])
            # test_data_list = override_test_codes(test_data_list, override_map)

    # ✅ Round minutes to nearest 5
            time_parts = discharge_data['time'].split(":")
            hour_24 = time_parts[0].zfill(2)
            rounded_minutes = round_minutes_to_5(time_parts[1])
            discharge_data['time'] = f"{hour_24}:{rounded_minutes}:00"
            
                # Fill claim
            try:
                fill_claim_id_and_submit(driver, claim_id, discharge_data,test_data_list)
            except Exception as e:
                print(f"❌ Error while processing claim {claim_id}: {e}")
                continue

            print("-" * 60)

    except Exception as e:
        print(f"❌ Login timeout or navigation error: {e}")
    finally:
        time.sleep(3800)
        driver.quit()


if __name__ == "__main__":
    main()