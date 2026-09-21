import time, json
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd

excel_path = r"C:\Users\saksh\OneDrive\Desktop\ECHS_Intimation\ECHS_Intimation.xlsx"
df = pd.read_excel(excel_path)
data = pd.DataFrame(df)
chrome_options = Options()
chrome_options.set_capability("goog:loggingPrefs", {"performance": "ALL"})
driver = webdriver.Chrome(service=Service(), options=chrome_options)
driver.execute_cdp_cmd("Network.enable", {})

def open_echs():
    driver.get("https://www.echsbpa.utiitsl.com/ECHS/")
    print(">> ECHS Site Opened")

    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,('/html/body/table/tbody/tr[2]/td[2]/form/table/tbody/tr[1]/td[3]/table/tbody/tr[2]/td[2]/input[1]')))).send_keys('hos.6966')
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,('/html/body/table/tbody/tr[2]/td[2]/form/table/tbody/tr[1]/td[3]/table/tbody/tr[4]/td[2]/input')))).send_keys('ECHS@1111')
    print(">> Entered Credentials")

# Fetching Captcha Text for ECHS
def fill_captcha_echs():
    global captcha_text
    captcha_text = ""
    try:
        for entry in driver.get_log("performance"):

            message = json.loads(entry["message"])["message"]

            if message["method"] != "Network.responseReceived":
                continue

            params = message["params"]
            response = params["response"]

            url = response["url"]

            if "type=g" not in url:
                continue

            request_id = params["requestId"]

            print("Found request:")
            print("URL = ",url)
            print("request id = ", request_id)

            result = driver.execute_cdp_cmd(
                "Network.getResponseBody",
                {"requestId": request_id}
            )

            captcha_text = result["body"][23:29]
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/table/tbody/tr[1]/td[3]/table/tbody/tr[10]/td[2]/input"))).send_keys(captcha_text)
            print(">> Captcha Filled > Text = ",captcha_text)
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/table/tbody/tr[1]/td[3]/table/tbody/tr[11]/td/input[1]"))).click()
    except:
        print(">> Couldn't Fill Captcha")
        while True:
            if len(driver.find_element(By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/table/tbody/tr[1]/td[3]/table/tbody/tr[10]/td[2]/input").get_attribute("value")) == 6:
                time.sleep(1)
            continue

def mou_nmi():
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[10]/div[3]/div/button"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ihaveseennmi"))).click()
    print(">> Closed MOU, Checked NMI Button")

def gen_claimid(i):
    # Take Excel Data in Variables
    card_num = str(data.loc[i,'card no'])
    room_type = str(data.loc[i,'room type']).strip().title()
    opd_num = str(data.loc[i,'opd num'])
    diagnosis = str(data.loc[i,'diagnosis'])
    polyclinic = str(data.loc[i,'polyclinic'])

    print(">> Fetched Excel Data")

    TOA = str(data.loc[i,'TOA'])
    TOAHR, TOAMIN, extra = TOA.split(":")
    if len(TOAHR) == 1:
        TOAHR = f"0{TOAHR}"
    if len(TOAMIN) == 1:
        TOAMIN = f"0{TOAMIN}"
    TOAMIN = round(int(TOAMIN) / 5) * 5
    TOAMIN = str(TOAMIN)

    DOA = str(data.loc[i,'DOA'])
    year, month, day = DOA.split("-")
    day = day.split(" ")[0]
    if "0" in month and month != "10":
        month = month[1:]
    if "0" in day and day != "10":
        day = day[1:]

    date_obj = datetime.strptime(f"{day}/{month}/{year}", "%d/%m/%Y")
    discharge_date = date_obj + timedelta(days=2)
    DOD = discharge_date.strftime("%d/%m/%Y")
    daydischarge, monthdischarge, yeardischarge = DOD.split("/")
    monthdischarge = str(int(monthdischarge))
    if "0" in monthdischarge and monthdischarge != "10":
        monthdischarge = monthdischarge[1:]
    if "0" in daydischarge and daydischarge != "10":
        daydischarge = daydischarge[1:]

    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[1]/form/div/ul/li[1]"))).click()
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[1]/form/div/ul/li[1]/ul/li[1]/a"))).click()

    # Enter Card Number
    Select(WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/table/tbody/tr[4]/td[2]/select")))).select_by_value("N")
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/table/tbody/tr[7]/td[2]/input"))).send_keys(card_num)
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/table/tbody/tr[8]/td/input"))).click()

    # Patient Details Tab
    Select(WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[1]/table/tbody/tr[9]/td[2]/select")))).select_by_value("3")

    # Admission Tab
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/ul/li[2]"))).click()
    Select(WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[2]/table[1]/tbody/tr[3]/td[2]/select")))).select_by_visible_text(polyclinic.title())
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[2]/table[1]/tbody/tr[6]/td[2]/input"))).send_keys(opd_num)
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[2]/table[1]/tbody/tr[7]/td[2]/input"))).send_keys("75000")
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[2]/table[1]/tbody/tr[8]/td[4]/input"))).send_keys("1")
    if "day" in room_type:
        Select(WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[2]/table[1]/tbody/tr[8]/td[2]/select")))).select_by_visible_text("Day Care")
    if "gen" in room_type:
        Select(WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[2]/table[1]/tbody/tr[8]/td[2]/select")))).select_by_visible_text("General")
    if "pri" in room_type and "semi" not in room_type:
        Select(WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[2]/table[1]/tbody/tr[8]/td[2]/select")))).select_by_visible_text("Private")
    if "semi" in room_type:
        Select(WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[2]/table[1]/tbody/tr[8]/td[2]/select")))).select_by_visible_text("Semi-Private")
    Select(WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[2]/table[1]/tbody/tr[5]/td[2]/select[1]")))).select_by_visible_text(TOAHR)
    Select(WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[2]/table[1]/tbody/tr[5]/td[2]/select[2]")))).select_by_visible_text(TOAMIN)
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[2]/table[1]/tbody/tr[5]/td[2]/img"))).click()
    Select(WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[13]/div/div/select[1]")))).select_by_value(str(int(month)-1))
    Select(WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[13]/div/div/select[2]")))).select_by_visible_text(year)
    WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.XPATH, f"//div[@id='ui-datepicker-div']//table/tbody//a[text()='{day}']"))).click()
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[2]/table[1]/tbody/tr[6]/td[4]/img"))).click()
    Select(WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[13]/div/div/select[1]")))).select_by_value(str(int(monthdischarge)-1))
    Select(WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[13]/div/div/select[2]")))).select_by_visible_text(yeardischarge)
    WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.XPATH, f"//div[@id='ui-datepicker-div']//table/tbody//a[text()='{daydischarge}']"))).click()
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[3]/table/tbody/tr[2]/td[2]/textarea"))).send_keys("Intimation")

    # Patient History Tab
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/ul/li[3]"))).click()
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[3]/table/tbody/tr[9]/td[2]/input[1]"))).click()
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[3]/table/tbody/tr[9]/td[2]/input[4]"))).send_keys("......")

    # Clinical Findings Tab
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,"/html/body/table/tbody/tr[2]/td[2]/form/div[2]/ul/li[4]"))).click()
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,"/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[4]/table/tbody/tr[2]/td[2]/input[1]"))).click()
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,"/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[4]/table/tbody/tr[2]/td[2]/input[4]"))).send_keys("70/-")

    # Diagnosis Tab
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/ul/li[5]"))).click()
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[5]/table[1]/tbody/tr[5]/td[2]/textarea"))).send_keys(diagnosis)

    # Proposed Treatment Tab
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/ul/li[6]"))).click()
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[6]/table/tbody/tr[4]/td[2]/input[1]"))).click()
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[2]/div[6]/table/tbody/tr[4]/td[2]/textarea"))).send_keys("Medical Management")

    # Fetch ClaimID
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/div[3]/table/tbody/tr[3]/td[1]/input[1]"))).click()
        time.sleep(1)
        success_text = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "succpara"))).text
        claimid = success_text.split(":")[-1].strip()
        data.loc[i, "claimid"] = claimid
        data.to_excel(excel_path,index=False)
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[12]/div[3]/div/button"))).click()
    except:
        print(">> ERROR: COULDN'T FETCH CLAIM ID KINDLY NOTE DOWN FROM SCREEN !!!")
        time.sleep(10000)



if __name__ == '__main__':
    open_echs()
    time.sleep(1)
    fill_captcha_echs()
    mou_nmi()
    #if 
    n = len(df)
    print(f">> Running Intimation Bot for {n} Patients")
    for i in range(n):
        if pd.isna(data.loc[i,'claimid']):
            gen_claimid(i)


    time.sleep(100000)
