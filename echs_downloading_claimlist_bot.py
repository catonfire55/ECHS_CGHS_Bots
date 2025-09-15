import win32gui
import win32con
import shutil
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from openpyxl import load_workbook
import time, os, psutil, subprocess, sys
import threading
from openpyxl.utils import get_column_letter


# ------------------ Helper Functions ------------------
def maximize_terminal():
    hwnd = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

def minimize_terminal():
    hwnd = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)

def wait_for_rename_completion(file_path, timeout=60):
    end_time = time.time() + timeout
    while time.time() < end_time:
        if os.path.exists(file_path):
            time.sleep(2)  # Allow buffer for rename to finish
            return True
        time.sleep(1)
    print(f"❌ File rename not detected: {file_path}")
    print(f"✅ Rename confirmed: {file_path}")

    return False

def get_chrome_driver_path():
    base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(__file__)
    return os.path.join(base_path, "chromedriver.exe")

# Create folder and ensure it's used correctly
def create_user_folder(user_id):
    today = time.strftime("%Y-%m-%d")
    base_path = os.path.join(os.path.expanduser("~"), "Documents", "EchsData")
    user_path = os.path.join(base_path, today, user_id)
    
    print(f"Attempting to create folder for User ID: {user_id} at {user_path}")
    
    if not os.path.exists(user_path):
        os.makedirs(user_path)  # Ensure directory is created if it doesn't exist
        print(f"✅ Created folder for User ID: {user_id} at {user_path}")
    else:
        print(f"⚠️ Folder for User ID: {user_id} already exists at {user_path}")
    
    return user_path



def close_moa_popup_if_exists(driver):
    try:
        WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="ui-id-1"]'))
        )
        close_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="mainheader"]/div[9]/div[3]/div/button'))
        )
        # scroll_into_view(driver, close_button)
        close_button.click()
        print("✅ Closed MoA expiry popup.")
        time.sleep(2)
    except:
        print("ℹ️ No MoA popup found. Proceeding normally...")

def handle_post_download_navigation(driver):
    try:
        dashboard_link = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="breadcrumbs"]/li[1]/a')))
        # scroll_into_view(driver, dashboard_link)
        #highlight_element(driver, dashboard_link)
        dashboard_link.click()
        print("↩️ Returned to main dashboard.")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="dashtable2"]/tbody/tr/td[2]')))
    except Exception as e:
        print(f"⚠️ Could not click Dashboard link after download: {e}")

def handle_post_login_navigation(driver):
    try:
        close_moa_popup_if_exists(driver)

        checkbox = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="ihaveseennmi"]'))
        )
        # scroll_into_view(driver, checkbox)
        #highlight_element(driver, checkbox)
        checkbox.click()
        print("✅ Checked the 'Need More Info' checkbox.")

        other_tab = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="li_7"]'))
        )
        # scroll_into_view(driver, other_tab)
        #highlight_element(driver, other_tab)
        other_tab.click()
        print("✅ Expanded 'Other' menu.")

        dashboard_link = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="7"]/li[3]/a'))
        )
        # scroll_into_view(driver, dashboard_link)
        #highlight_element(driver, dashboard_link)
        dashboard_link.click()
        print("✅ Clicked 'Dashboard' link.")
        time.sleep(4)

    except Exception as e:
        print(f"⚠️ Error in post-login navigation: {e}")

# def highlight_element(driver, element, effect_time=1):
#     original_style = element.get_attribute('style')
#     driver.execute_script("arguments[0].setAttribute('style', arguments[1]);",
#                          element, "border: 2px solid red; background: yellow;")
#     time.sleep(effect_time)
#     driver.execute_script("arguments[0].setAttribute('style', arguments[1]);",
#                          element, original_style)

# def scroll_into_view(driver, element):
#     driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
def timer_check_for_user_login(start_time, timeout=60):
    if time.time() - start_time > timeout:
        return True  # 5 minutes have passed
    return False

start_time = time.time()  # Start the timer when the first user logs in

def wait_for_download_to_finish(download_dir, timeout=100):
    end_time = time.time() + timeout
    while time.time() < end_time:
        if not [f for f in os.listdir(download_dir) if f.endswith('.crdownload')]:
            return True
        time.sleep(1)
    return False

def click_element_when_ready(driver, element, count_id, retries=3):
    for attempt in range(retries):
        try:
            WebDriverWait(driver, 20).until(EC.visibility_of(element))
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(element))
            print(f"🔹 Clicking element with ID: {count_id} (attempt {attempt + 1})")
            #highlight_element(driver, element)
            ActionChains(driver).move_to_element(element).click().perform()
            print(f"✅ Clicked on count ID: {count_id}")
            return True
        except Exception as ex:
            print(f"⚠️ Could not click element '{count_id}' on attempt {attempt + 1}: {ex}")
            time.sleep(2)
    return False

def export_to_excel(driver, download_dir, retries=2, timeout=100):
    for attempt in range(1, retries + 1):
        try:
            select_display = Select(WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, "displayMode"))))
            select_display.select_by_visible_text("Export To Excel")
            print(f"✅ Export option selected (Attempt {attempt})")

            submit_btn = driver.find_element(By.XPATH, "//input[@value='Submit']")
            # scroll_into_view(driver, submit_btn)f
            #highlight_element(driver, submit_btn)
            submit_btn.click()

            if wait_for_download_to_finish(download_dir, timeout=timeout):
                print(f"✅ Excel file downloaded (attempt {attempt}).")
                return True
            else:
                print(f"⚠️ Attempt {attempt} download timeout. Retrying...")
                time.sleep(2)

        except Exception as ex:
            print(f"⚠️ Attempt {attempt} failed: {ex}")
            time.sleep(2)

    print("❌ All export attempts failed.")
    return False






def rename_file_to_status(download_dir, status_text, latest_filename, force_use_status=False, suffix=""):
    import datetime
    import xlrd
    import re

    def clean_status_text(raw_text, skip_cleaning=False):
        if not raw_text:
            return "Unknown_Status"

        if skip_cleaning:
            return raw_text.strip().title()

        patterns = [
            r"(?i)\bas on\b.*", 
            r"\bprocessed on\b.*",
            r"\bsubmitted on\b.*",
            r"\bon\b.*",
            r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}",  # dates like 01/05/2025
            r"\d{8}",  # dates like 20250501
        ]
        for pattern in patterns:
            raw_text = re.sub(pattern, '', raw_text, flags=re.IGNORECASE)

        return raw_text.strip().title()

    old_path = os.path.join(download_dir, latest_filename)
    new_status = status_text  # default

    if not force_use_status:
        try:
            wb = xlrd.open_workbook(old_path)
            sheet = wb.sheet_by_index(0)
            raw = sheet.cell_value(6, 0).strip()
            print(f"📄 Raw status text from row 7: '{raw}'")  # Show raw extracted line
            if raw:
                skip_cleaning = "Claim Settled" in status_text  # 13ST case
                new_status = clean_status_text(raw, skip_cleaning=skip_cleaning)
                print(f"✅ Cleaned status used for rename: '{new_status}'")
        except Exception as e:
            print(f"⚠️ Could not extract status from Excel content: {e}")

    # Final rename
    safe_status = new_status.replace("_", " ").replace("/", "").strip()
    timestamp = datetime.datetime.now().strftime("%H%M%S")
    new_filename = f"{safe_status}{suffix}_{timestamp}.xls"
    new_path = os.path.join(download_dir, new_filename)

    if os.path.exists(new_path):
        os.remove(new_path)

    shutil.move(old_path, new_path)
    print(f"✅ Renamed '{latest_filename}' ➔ '{new_filename}'")
    return new_filename



def wait_for_specific_file(download_dir, prefix, timeout=300):
    print(f"\u23F3 Waiting for file: {prefix}.xls (up to {timeout} seconds)")
    end_time = time.time() + timeout
    while time.time() < end_time:
        for f in os.listdir(download_dir):
            if f.startswith(prefix) and f.endswith(".xls") and not f.endswith(".crdownload"):
                print(f"\u2705 File detected: {f}")
                return f
        time.sleep(1)
    print(f"\u274C File not found after waiting: {prefix}.xls")
    raise FileNotFoundError(f"File {prefix}.xls not found")






def get_status_text(count_id):
    status_mapping = {
        "13PT": "Processed for Settlement",
        "99HI" : "Hold Claim",
        "99PT": "Rejected Claims Processed", "99CT": "Cancel Claim", "99XT": "Referral Validity Expired",
        "99YT": "Inactive Intimations"
        , "1PT": "Admission Intimation Pending",
        "1ST": "Admission Intimation Submitted", "1YT": "Emergency Intimation To Polyclinic",
        "2MT": "Admission Intimated - With More Info", "1XT": "Emergency Intimation Rejected by Polyclinic",
        "2ST": "Intimation Acknowledged", "2NT": "Need More Information [Int]",
        "2XT": "Not Entitled", "3PT": "Claim Submission Pending", "4ST": "Claim Documents Received & Verified",
        "5ST": "Scrutinizer Verified", "6ST": "Claim Authorized [BPA]",
        "6NT": "Need More Information [Val]", "6XT": "Recommended for Rejection [BPA]",
        "7ST": "Recommended for Approval [RC]", "7RT": "Review By Validator [App]",
        "7NT": "Need More Information [App]", "7XT": "Recommended for Rejection [RC]",
        "12ST": "Recommended for Approval [COrg]", "12RT": "Review By Validator [COrg]",
        "12NT": "Need More Information [App]", "8ST": "Approved [RC. Dir]",
        "8RT": "Review Required by MO [RC. Dir]", 
        "10RT": "Review Required by MO (MD)",
         "9ST": "Approved for Payment [Dy.MD]",
         "3ST": "Claim Submitted Electronically",
        "9XT" : "Rejected (Dy.MD)"
        
        
    }
    return status_mapping.get(count_id, "Unknown_Status")

    

def merge_raw_append(download_dir):
    excel_files = [f for f in os.listdir(download_dir) if f.endswith('.xls')]
    if not excel_files:
        print("⚠️ No Excel files found to merge.")
        return

    sheets_by_colcount = {}
    total_rows_written = 0  # Track total row count
    for file_name in excel_files:
        path = os.path.join(download_dir, file_name)
        print(f"🔹 Reading file: {path}")
        try:
            df = pd.read_excel(path, engine='xlrd', header=None, dtype=str, skiprows=range(0, 9), keep_default_na=False)
            try:
                headers = pd.read_excel(path, engine='xlrd', header=None, dtype=str,
                                        skiprows=8, nrows=1, keep_default_na=False).iloc[0].tolist()
            except:
                headers = [f"Col{i + 1}" for i in range(df.shape[1])]

            df.iloc[:, 0] = df.iloc[:, 0].astype(str)

            if any(key in file_name.lower() for key in ["claim settled", "hold claim"]):
                df = df[~df.iloc[:, 0].str.lower().str.strip().eq("claim id")]

            for word in ["grand total", "sub total", "sub-total", "accepted since days", "total"]:
                df = df[~df.iloc[:, 0].str.lower().str.contains(word)]

            if df.empty:
                print(f"⚠️ Skipped: {file_name} (no valid rows after filtering)")
                continue

            status_label = os.path.splitext(file_name)[0].rsplit("_", 1)[0]  # Remove timestamp part
            df["Status"] = status_label
            headers.append("Status")

            col_count = df.shape[1]
            if col_count not in sheets_by_colcount:
                sheets_by_colcount[col_count] = {"data": [], "headers": headers}
            sheets_by_colcount[col_count]["data"].append(df)

        except Exception as e:
            print(f"⚠️ Could not read '{file_name}': {e}")


    merged_path = os.path.join(download_dir, "Final_Merged.xlsx")
    with pd.ExcelWriter(merged_path, engine="openpyxl") as writer:
        for count, info in sheets_by_colcount.items():
            if not info["data"]:
                continue
            df_final = pd.concat(info["data"], ignore_index=True)
            df_final.columns = info["headers"]
            df_final.to_excel(writer, sheet_name=f"{count}_cols"[:31], index=False)
            row_count = len(df_final)
            total_rows_written += row_count
            print(f"📝 Sheet '{count}_cols': {row_count} rows written.")

    print(f"✅ Merged all files into: {merged_path}")
    print(f"📊 Total rows written in Final_Merged.xlsx: {total_rows_written}")
def merge_all_final_files(base_path):
    today = time.strftime("%Y-%m-%d")
    today_path = os.path.join(base_path, today)
    all_merged_folder = os.path.join(today_path, "All_Merged_Files")
    os.makedirs(all_merged_folder, exist_ok=True)

    final_files = []

    for user_folder in os.listdir(today_path):
        user_folder_path = os.path.join(today_path, user_folder)
        if os.path.isdir(user_folder_path):
            for root, dirs, files in os.walk(user_folder_path):
                for file in files:
                    if file.lower() == "final_merged.xlsx":
                        final_files.append(os.path.join(root, file))

    if not final_files:
        print("\u26a0\ufe0f No.xlsx files found in today's folders.")
        return

    sheets_by_colcount = {}

    print("\n📘 Reading each Final_Merged.xlsx file:")
    for file_path in final_files:
        try:
            data = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
            for sheet, df in data.items():
                col_count = df.shape[1]
                sheets_by_colcount.setdefault(col_count, []).append(df)
                print(f"  ✅ {os.path.basename(file_path)} -> Sheet '{sheet}': {len(df)} rows")
        except Exception as e:
            print(f"  ❌ Error reading {file_path}: {e}")

    combined_path = os.path.join(all_merged_folder, "All_Final_Merged.xlsx")
    print(f"\n📦 Merging into: {combined_path}")

    merged_total = 0
    with pd.ExcelWriter(combined_path, engine="openpyxl") as writer:
        for col_count, dfs in sheets_by_colcount.items():
            combined_df = pd.concat(dfs, ignore_index=True)
            sheet_name = f"{col_count}_cols"
            combined_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            row_count = len(combined_df)
            merged_total += row_count
            print(f"  ✅ Written sheet '{sheet_name}': {row_count} rows")

    print("\n📊 Verifying total row count from all original Final_Merged.xlsx files:")
    original_total = sum(len(df) for dfs in sheets_by_colcount.values() for df in dfs)
    print(f"  📄 Total expected rows: {original_total}")
    print(f"  📄 Total written in All_Final_Merged.xlsx: {merged_total}")

    if original_total == merged_total:
        print("\n🎉 Success: All rows were correctly merged. No data skipped.")
    else:
        print("\n⚠️ Row count mismatch! Some rows may be missing.")

# ------------------ Main Logic ------------------
def wait_for_file_rename(directory, expected_filename, timeout=200):
    import time
    import os
    
    print(f"⏳ Waiting for file: {expected_filename} (up to {timeout} seconds)")
    expected_path = os.path.join(directory, expected_filename)
    end_time = time.time() + timeout
    while time.time() < end_time:
        if os.path.exists(expected_path):
            print(f"✅ File found: {expected_filename}")
            return True
        time.sleep(1)
    print(f"❌ File not found after waiting: {expected_filename}")
    return False
def main():
    today = time.strftime("%Y-%m-%d")
    base_path = os.path.join(os.path.expanduser("~"), "Documents", "EchsData")

    count_ids = ["13PT", "99HI"
                 ,"99PT", "99CT", "99XT", "99YT",
    "1PT", "1ST", "1YT", "2MT", "1XT",
    "2ST", "2NT", "2XT", "3PT", "4ST",
    "5ST", "6ST", "6NT", "6XT", "7ST",
    "7RT", "7NT", "7XT", "12ST", "12RT",
    "12NT", "8ST", "8RT", "10RT","9ST","9XT"
    ]

    while True:
        # options = Options()
        # options.add_experimental_option("prefs", {
        #     "download.prompt_for_download": False,
        #     "plugins.always_open_pdf_externally": True
        # })
        # options.add_argument('--ignore-certificate-errors')
        # options.add_argument('--start-maximized')

        driver = webdriver.Chrome() # service=Service(get_chrome_driver_path()), options=options
        driver.get("https://echsbpa.utiitsl.com/ECHS/....do")
        print("\u23F3 Waiting for user to enter User ID and manually log in...")

        WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//*[@id="username"]')))

        user_id_fake_dropped = ""
        try:
            start_time = time.time()
            while not user_id_fake_dropped and time.time() - start_time < 28800:
                user_id_element = WebDriverWait(driver, 20).until(
                    EC.visibility_of_element_located((By.XPATH, '//*[@id="username"]')))
                user_id_fake_dropped = user_id_element.get_attribute("value").strip()  
                if not user_id_fake_dropped:
                    print("⚠️ User ID is empty! Please log in.")
                    time.sleep(3)
            if not user_id_fake_dropped:
                print("\u23F3 8 Hours have passed without user login. Proceeding with the final merge.")
                merge_all_final_files(base_path)
                driver.quit()
                return
            else:
                print(f"👤 Detected User ID....")
        except Exception as e:
            print(f"⚠️ Error detecting user ID: {e}")
            driver.quit()
            return
        
        user_id = "SHRCECHS"
        user_path = create_user_folder(user_id)

        WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ihaveseennmi"]')))
        handle_post_login_navigation(driver)

        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": user_path
        })
        print(f"Download folder is set to: {user_path}")

        for count_id in count_ids:
            try:
                print(f"\n🔹 Checking count ID: {count_id}")
                element = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.ID, count_id)))
                status_text = get_status_text(count_id)

                if not click_element_when_ready(driver, element, count_id):
                    print(f"❌ Skipping {count_id} due to click failure.")
                    continue
                time.sleep(1)

                # if count_id == "13ST":
                #     try:
                #         print("🔍 Special handling for 13ST with SHRCECHS...")

                #         patient_type_xpath = '//*[@id="newfilter"]/td[2]/select'
                #         WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, patient_type_xpath)))
                #         Select(driver.find_element(By.XPATH, patient_type_xpath)).select_by_visible_text("In-Patient")
                #         print("➡️ Selected In-Patient")
                #         time.sleep(2)

                #         if export_to_excel(driver, user_path, timeout=200):
                #             latest_file = wait_for_specific_file(user_path, "CLAIMLIST_" + user_id.lower(), timeout=200)
                #             renamed_file = rename_file_to_status(user_path, status_text, latest_file, force_use_status=True, suffix="_InPatient")
                #             wait_for_rename_completion(os.path.join(user_path, renamed_file))
                #         else:
                #             raise Exception("In-Patient download failed")

                #         handle_post_download_navigation(driver)

                #         # --- Out-Patient ---
                #         print("🔁 Re-clicking 13ST for Out-Patient...")
                #         element = WebDriverWait(driver, 20).until(
                #             EC.element_to_be_clickable((By.ID, count_id)))
                #         if not click_element_when_ready(driver, element, count_id):
                #             raise Exception("13ST re-click failed")
                #         time.sleep(2)

                #         Select(driver.find_element(By.XPATH, patient_type_xpath)).select_by_visible_text("Out-Patient")
                #         print("➡️ Selected Out-Patient")
                #         time.sleep(2)

                #         fin_year_xpath = '//*[@id="financialYear"]'
                #         WebDriverWait(driver, 10).until(
                #             EC.presence_of_element_located((By.XPATH, fin_year_xpath)))

                #         year_dropdown = Select(driver.find_element(By.XPATH, fin_year_xpath))
                #         total_years = len(year_dropdown.options)

                #         for i in range(1, total_years):
                #             year_text = year_dropdown.options[i].text.strip()
                #             year_dropdown.select_by_index(i)
                #             print(f"📅 Selected Financial Year: {year_text}")
                #             time.sleep(1)

                #             WebDriverWait(driver, 10).until(
                #                 EC.presence_of_element_located((By.NAME, "displayMode")))

                #             Select(driver.find_element(By.NAME, "displayMode")).select_by_visible_text("Export To Excel")
                #             print("📦 Display Mode set to 'Export To Excel'")
                #             time.sleep(1)

                #             submit_btn = WebDriverWait(driver, 10).until(
                #                 EC.element_to_be_clickable((By.XPATH, "//input[@value='Submit']")))
                #             scroll_into_view(driver, submit_btn)
                #             highlight_element(driver, submit_btn)
                #             submit_btn.click()
                #             print("🚀 Clicked Submit to trigger export")
                #             time.sleep(3)

                #             if export_to_excel(driver, user_path, timeout=180):
                #                 latest_file = wait_for_specific_file(user_path, "CLAIMLIST_" + user_id.lower(), timeout=300)
                #                 renamed_file = rename_file_to_status(user_path, status_text, latest_file, force_use_status=True, suffix=f"_OutPatient_{year_text}")
                #                 wait_for_rename_completion(os.path.join(user_path, renamed_file))
                #             else:
                #                 print(f"⚠️ Download failed for year: {year_text}")

                #         handle_post_download_navigation(driver)

                #     except Exception as e:
                #         print(f"⚠️ Special 13ST handling failed: {e}")
                if count_id == "13PT":
                    try:
                        print("📌 Special handling for 13PT – forcing status for renaming...")
                        if not export_to_excel(driver, user_path):
                            raise Exception("Export failed for 13PT")

                        latest_file = wait_for_specific_file(user_path, "CLAIMLIST_" + user_id.lower(), timeout=300)
                        renamed_file = rename_file_to_status(user_path, status_text, latest_file, force_use_status=True)
                        wait_for_rename_completion(os.path.join(user_path, renamed_file))
                        handle_post_download_navigation(driver)

                    except Exception as e:
                        print(f"⚠️ Special 13PT handling failed: {e}")


                else:
                    if not export_to_excel(driver, user_path):
                        print(f"❌ Skipping {count_id} due to export failure.")
                        continue

                    latest_file = wait_for_specific_file(user_path, "CLAIMLIST_" + user_id.lower(), timeout=300)
                    renamed_file = rename_file_to_status(user_path, status_text, latest_file)
                    wait_for_rename_completion(os.path.join(user_path, renamed_file))
                    handle_post_download_navigation(driver)

            except Exception as e:
                print(f"⚠️ Could not handle count ID '{count_id}': {e}")
                

        print(f"\n✅ Done for User: {user_id}")
        print(f"📂 Files saved at: {user_path}")
        merge_raw_append(user_path)

        print("🛑 No further user login. Generating All Final Merged file now...")
        merge_all_final_files(base_path)
        print("✅ Final All Merged File created.")
        driver.quit()
        break

print(">>>Checking for previous junk data...")
pcuser = os.getlogin()
try:
    if os.path.exists(f"C:/Users/{pcuser}/Documents/EchsData"):
        shutil.rmtree(f"C:/Users/{pcuser}/Documents/EchsData")
        print(">>> Old Junk Files Found... Deleted..!!")
except Exception as e:
    print(">>> No Previous Junk Files Found....")
try:
    if os.path.exists(f"C:/Users/{pcuser}/Documents/temp_dir"):
        shutil.rmtree(f"C:/Users/{pcuser}/Documents/temp_dir")
        print(">>> Old Junk Files Found... Deleted..!!")
except Exception as e:
    print(">>> No Previous Junk Files Found....")
main()
maximize_terminal()
time.sleep(1)
print("\n\n\n\n>>> Now See Saksham Jain's Magic 😉 ")
print(">>> Making Claim Settled Merged File...")
print(">>> You Will Need to Login To ECHS Again....")
print(f"\n\n\n>>>Start Time => {time.ctime()}\n\n")
time.sleep(2)
minimize_terminal()


# CLAIM_SETTLED_BOT_SAKSHAM STARTS HERE !!!!



pcuser = os.getlogin()
download_path_og = os.path.join(os.path.expanduser("~"), "Documents", "temp_dir")
rows_del = [1,2,3,4,5,6,7]
fix_columns_indi = ["Claim ID","Region","Hospital Name","Card ID","Name Of ESM","Patient Name","Patient Type","Admit Type","Accept Date","Net Claim Amt.","Approved Amt.","Processed On"]
fix_columns = ["Claim ID","Region","Hospital Name","Card ID","Name Of ESM","Patient Name","Patient Type","Admit Type","Accept Date","Net Claim Amt.","Approved Amt.","Processed On","Status"]
fix_col_width = {
    'A': 13,
    'B': 10,
    'C': 49.4,
    'D': 20,
    'E': 22,
    'F': 22.5,
    'G': 10.5,
    'H': 10,
    'I': 18,
    'J': 15.5,
    'K': 17.2,
    'L': 22,
    'M': 37,
}
if os.path.exists(download_path_og):
    shutil.rmtree(download_path_og)
    os.makedirs(download_path_og)
    print(">>> Found Previous Files...Cleaned")
else:
    os.makedirs(download_path_og)
    print(">>> Made the download dir...")
# options = Options()
# options.add_experimental_option("prefs", {
#     "download.prompt_for_download": False,
#     "plugins.always_open_pdf_externally": True
# })
# options.add_argument('--ignore-certificate-errors')
# options.add_argument('--start-maximized')

driver = webdriver.Chrome() # service=Service(get_chrome_driver_path()), options=options
driver.execute_cdp_cmd("Page.setDownloadBehavior", {"behavior": "allow", "downloadPath": download_path_og})
print(">>> Changed the download path to Documents/temp_dir")

def check_download(directory, prefix, timeout=120):
    #print(f"/u23F3 Waiting for file: {prefix}.xls (up to {timeout} seconds)")
    end_time = time.time() + timeout
    while time.time() < end_time:
        for filename in os.listdir(directory):
            if filename.startswith(prefix) and filename.endswith(('.xls', '.xlsx')) and not filename.endswith('.crdownload'):
                print(f">>>Download complete: {filename}")
                return filename
        time.sleep(1)  # Check every second
    print(f">>>Download (failed) timed out after {timeout} seconds for {prefix}")
    return None

def rename_og_to_yearwise(year_text):
    old_path = (f"C:/Users/{pcuser}/Documents/temp_dir")
    new_path = (f"C:/Users/{pcuser}/Documents/SHRCECHS_ClaimBot/CLAIM_SETTLED_YEARWISE_CLEANED")
    print(f">>>Renaming Og File to Yearwise - {year_text}")
    old_file = os.path.join(old_path,"CLAIMLIST_shrcechs.xls")
    new_file = os.path.join(new_path,f"claim_settled_by_year_{year_text}.xlsx")
    os.rename(old_file,new_file)

def clean_data(year_text):
    print(">>> Cleaning Data...")
    new_path = (f"C:/Users/{pcuser}/Documents/SHRCECHS_ClaimBot/CLAIM_SETTLED_YEARWISE_CLEANED")
    excel_file_to_clean = os.path.join(new_path,f"claim_settled_by_year_{year_text}.xlsx")
    df = pd.read_excel(excel_file_to_clean)
    df = df.drop(rows_del)
    df.columns = fix_columns_indi
    df.to_excel(excel_file_to_clean, index=False)

    print(">>>Data Cleaning Successful")
    
def merge_all_claim_settled():
        merged_file_claimsettled_loc = f"C:/Users/{pcuser}/Documents/SHRCECHS_ClaimBot/merged_claim_settled.xlsx"
        dfs = (pd.read_excel(f"C:/Users/{pcuser}/Documents/SHRCECHS_ClaimBot/CLAIM_SETTLED_YEARWISE_CLEANED/claim_settled_by_year_{2012+i}.xlsx") for i in range(1,total_financial_years))
        merged_file = pd.concat(dfs, ignore_index=True)
        merged_file_claimsettled = merged_file
        merged_file_claimsettled["Status"] = "Claim Settled"
        merged_file_claimsettled.columns = fix_columns
        merged_file_claimsettled.drop(0)
        merged_file_claimsettled = merged_file_claimsettled.style.set_properties(**{'text-align': 'left'})
        merged_file_claimsettled.to_excel(merged_file_claimsettled_loc, index=False)
        print(">>> Merging Files Success... ")

        print(">>> Arranging Columns...")
        book = load_workbook(merged_file_claimsettled_loc)
        sheet = book.active
        for col, width in fix_col_width.items():
            sheet.column_dimensions[col].width = width
        book.save(merged_file_claimsettled_loc)
        print(">>> Arranged Column Widths...")

print("!!!!! Welcome... !!!!!")
if os.path.exists(f"C:/Users/{pcuser}/Documents/SHRCECHS_ClaimBot"):
    print(">>>Previous Files Detected....Cleaning Junk....")
    shutil.rmtree(f"C:/Users/{pcuser}/Documents/SHRCECHS_ClaimBot")
    os.makedirs(f"C:/Users/{pcuser}/Documents/SHRCECHS_ClaimBot/CLAIM_SETTLED_YEARWISE_CLEANED", exist_ok=False)
else:
    os.makedirs(f"C:/Users/{pcuser}/Documents/SHRCECHS_ClaimBot/CLAIM_SETTLED_YEARWISE_CLEANED", exist_ok=False)
    print(">>>Clean Directory Made...")

driver.get("https://www.echsbpa.utiitsl.com/ECHS")
driver.implicitly_wait(10)
print(">>>enter the captcha...")
WebDriverWait(driver, 28800).until(EC.visibility_of_element_located((By.ID, 'infopara')))

driver.implicitly_wait(2)
time.sleep(1)

try:
    moa_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mainheader"]/div[9]/div[3]/div/button')))
    driver.find_element(By.XPATH,'//*[@id="mainheader"]/div[9]/div[3]/div/button').click()
    print(">>>MoA Popup Closed...")
except:
    print(">>>No MoA Popup Found ;)")
driver.implicitly_wait(1)

ihaveseennmi_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ihaveseennmi"]')))
ihaveseennmi_button.click()
print(">>>clicked on ihaveseennmi")

(driver.find_element(By.ID, 'li_7')).click()
driver.implicitly_wait(1)
driver.find_element(By.XPATH, '//*[@id="7"]/li[3]/a').click()
driver.implicitly_wait(2)
WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "13ST"))).click()
print(">>>Clicked On Claim Settled...")

WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="newfilter"]/td[2]/select')))
Select(driver.find_element(By.XPATH, '//*[@id="newfilter"]/td[2]/select')).select_by_visible_text("Both")
driver.implicitly_wait(1)

financial_years = Select((driver.find_element(By.ID, "financialYear")))
total_financial_years = len(financial_years.options)
for i in range(1,total_financial_years):
    financial_years.select_by_index(i)
    selected_year_text_for_print = financial_years.options[i].text.strip()
    selected_year_text_for_name = f"{2012+i}"
    print(f">>>Selected Year -> {selected_year_text_for_print}")
    select_excel_export = Select(driver.find_element(By.XPATH, '//*[@id="filterTable"]/tbody/tr[14]/td[2]/select'))
    select_excel_export.select_by_visible_text("Export To Excel")
    submit_button = driver.find_element(By.XPATH, '//*[@id="submitbutton"]/td/input').click()
    time.sleep(0.2)
    download_result=check_download(download_path_og, "CLAIMLIST_shrcechs")
    if download_result:
        print(f">>>Found File...")
        rename_og_to_yearwise(selected_year_text_for_name)
        print(f">>>Renamed File to claim_settled_by_year_{selected_year_text_for_name}.xlsx successfully..")
        clean_data(selected_year_text_for_name)
    else:
        print(">>> Download failed, stopping execution.")
        break

print(">>>All Indivisual Files are downloaded and cleaned...")
driver.quit()
merge_all_claim_settled()
maximize_terminal()
time.sleep(1)
print("\n\n\n>>> Now Merging Bot Starts 😉\n\n\n")
print(">>> It would take 2-3 mins... Kindly have patience ❤️...")



# MERGING OF FINAL FILE STARTS HERE



pcuser = os.getlogin()
today = time.strftime("%Y-%m-%d")
both_final_files = [
    f"C:/Users/{pcuser}/Documents/SHRCECHS_ClaimBot/merged_claim_settled.xlsx",
    f"C:/Users/{pcuser}/Documents/EchsData/{today}/All_Merged_Files/All_Final_Merged.xlsx"
]
final_output = f"C:/Users/{pcuser}/Documents/Full_Dashboard_ECHS_dated_{today}.xlsx"
df1 = pd.read_excel(both_final_files[0], sheet_name=None)
df1_combined = pd.concat(df1.values(), ignore_index=True)
print(">>> Read ClaimSettled File...")
df2 = pd.read_excel(both_final_files[1], sheet_name=None)
df2_combined = pd.concat(df2.values(), ignore_index=True)
print(">>> Read OtherFilesMerged...")
final_df = pd.concat([df1_combined, df2_combined], ignore_index=True)
final_df.to_excel(final_output, index=False)
book = load_workbook(final_output)
sheet = book.active
for col, width in fix_col_width.items():
        sheet.column_dimensions[col].width = width
book.save(final_output)
print("\n!!!!! SUCCESS !!!!!!! ")
print(f">>> The Final File Is Stored In Documents Itself...\n\n")

choice = input(">>> Would you like to clear old junk files?? (enter 'y' for yes or any key for exiting program.....) > ")
if choice == "y":
    print(">>> I Respect Your Decision And Concern For Your Storage Efficiency....")
    print(">>> Removing Old Trash Files...")
    shutil.rmtree(f"C:/Users/{pcuser}/Documents/EchsData")
    shutil.rmtree(f"C:/Users/{pcuser}/Documents/temp_dir")
    try:
        shutil.rmtree(f"C:/Users/{pcuser}/Documents/SHRCECHS_ClaimBot")
    except Exception as e:
        print("..")   
    print("--------------------Thankyou For Using This Tool 😊 BYEEE!!!!--------------------")
    time.sleep(2)
    os.startfile(final_output)
    exit
else:
    print(">>> I Respect Your Decision....")
    print("--------------------Thankyou For Using This Tool 😊 BYEEE!!!!--------------------")
    print(f"\n\n>>> End Time => {time.ctime()}")
    os.startfile(final_output)
    time.sleep(2)
    exit