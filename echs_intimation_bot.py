import time, json
# import os, psutil, subprocess, base64
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
# import pandas as pd

chrome_options = Options()
chrome_options.set_capability("goog:loggingPrefs", {"performance": "ALL"})
driver = webdriver.Chrome(service=Service(), options=chrome_options)
driver.execute_cdp_cmd("Network.enable", {})

driver.get("https://www.echsbpa.utiitsl.com/ECHS/")
print(">> ECHS Site Opened...")
time.sleep(1)

WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,('/html/body/table/tbody/tr[2]/td[2]/form/table/tbody/tr[1]/td[3]/table/tbody/tr[2]/td[2]/input[1]')))).send_keys('ECHS@1111')
WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,('/html/body/table/tbody/tr[2]/td[2]/form/table/tbody/tr[1]/td[3]/table/tbody/tr[4]/td[2]/input')))).send_keys('hos.6966')
print(">> Entered Credentials...")

# Fetching Captcha Text for ECHS
captcha_text = ""
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
        {
            "requestId": request_id
        }
    )

    captcha_text = result["body"][23:29]
    print(">> Captcha Text = ",captcha_text)

WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td[2]/form/table/tbody/tr[1]/td[3]/table/tbody/tr[10]/td[2]/input"))).send_keys(captcha_text)

time.sleep(10)