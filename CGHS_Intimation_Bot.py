# ================================================================ WORKING PLAN =====================================================================
# AN AI GENERATED BOT REVIEW AND SUMMARY FOR NEXT DEVELOPERS - ( I haven't checked the summary dear it's just a simple code overview )
#
# HIGH-LEVEL PROCESS FLOW:
#
# PHASE 1: DATA INITIALIZATION
# ────────────────────────────────────────────────────────────────────────────
# 1. Load main patient Excel file from user's Documents folder
# 2. Load doctor details database from local backup
# 3. Load MedNet credentials from secure storage
# 4. Add placeholder columns for contact number, address, doctor name, and ward type
# 5. Initialize separate tracking lists for emergency and referral cases
#
#
# PHASE 2: MedNet DATA EXTRACTION
# ────────────────────────────────────────────────────────────────────────────
# 1. LOGIN TO MedNet PORTAL (Hospital management system)
#    - Enter credentials from loaded Excel file
#    - Navigate to Inpatient/Reception section
#
# 2. PROCESS EACH PATIENT IN EXCEL
#    - Filter patients with either CARD NO (emergency) or REFERRAL NO (referral cases)
#    - For each valid patient:
#      * Extract registration number
#      * Parse bill date and convert to date components (day, month)
#      * Open patient bill record using date and registration number
#      * Scrape patient information:
#        - Contact number
#        - Address
#        - Primary doctor name(s)
#        - Ward type (critical/ward/single/twin)
#      * Store extracted data back to Excel file
#
# 3. CATEGORIZE CASES
#    - Add index to emergency cases list (has CARD NO)
#    - Add index to referral cases list (has REFERRAL NO)
#
#
# PHASE 3: EMERGENCY CASE PROCESSING (EMG FORM GENERATION)
# ────────────────────────────────────────────────────────────────────────────
# For each emergency case:
#
# 1. CREATE EMG INTIMATION DOCUMENT
#    - Load template DOCX file
#    - Fill form fields with patient data (11 specific line items):
#      * Line 6: Patient name
#      * Line 7: Insurance card number
#      * Line 8: Contact number
#      * Line 9: Address
#      * Line 10: IP (Inpatient) number
#      * Line 11: Bill date
#      * Line 12: Bill time
#      * Line 13: Chief complaints
#      * Line 15: Diagnosis
#      * Line 24: Doctor name
#      * Line 25: Doctor designation
#      * Line 26: Current date
#    - Save as DOCX file
#    - Convert to PDF
#    - Delete temporary DOCX file
#
#
# PHASE 4: EMERGENCY CASE SUBMISSION (NHA PORTAL)
# ────────────────────────────────────────────────────────────────────────────
# For each emergency case, perform automated form filling on NHA pre-authorization portal:
#
# 1. ACCESS NHA PRE-AUTHORIZATION PORTAL
#    - Open provider.nha.gov.in
#    - Close any popup modals if present
#
# 2. SEARCH PATIENT RECORD
#    - Enter claim number in search field
#    - Click search button
#    - Click on patient record from results
#    - Click "Medical Info" button
#    - Click "Personal Info" button
#
# 3. FILL PERSONAL INFORMATION SECTION
#    - Check button state (EDIT or SAVE)
#    - If EDIT mode: click to enable editing
#    - Select "NO" for two yes/no fields (e.g., pre-existing conditions)
#
# 4. FILL ADMISSION DETAILS SECTION
#    - Open next section (Admission Information)
#    - Click "Admission Details" expandable section
#    - Check and enable EDIT mode if needed
#    - Set admission date (calendar picker)
#    - Set admission time (hour, minute, AM/PM dropdowns)
#    - Set surgery/discharge date
#    - Select "NO" for MLC (Medico-Legal Case)
#    - Select admission type: "Emergency"
#
# 5. FILL TREATMENT SECTION
#    - Open Treatment column
#    - Click "Referral Manual" subsection
#    - Enter IP number
#    - Select wellness center: "D70"
#    - Set referral date
#    - Enter treatment details:
#      * Code: "emg"
#      * Facility: "shrc"
#      * Remarks: "Intimation"
#    - Upload EMG form PDF (file dialog)
#
# 6. FILL DIAGNOSIS INFORMATION
#    - Open Diagnosis subsection
#    - Enter first diagnosis code and select from dropdown
#    - Click "Add Another"
#    - Enter second diagnosis code and select from dropdown
#    - Add as many diagnosis entries as needed
#
# 7. FILL TREATMENT PLAN
#    - Open Treatment Plan section
#    - Add ward type entry: "critical", "ICU", duration "7"
#    - Add consultation entry: "consult", "Inpatient", duration "18"
#
# 8. UPLOAD SUPPORTING DOCUMENTS
#    - Open Investigations/Documents section
#    - Upload insurance card PDF
#    - Upload patient photo JPEG
#
# 9. FILL CARE TEAM DETAILS
#    - Open Care Team section
#    - Select practitioner type: "Other"
#    - Enter doctor name
#    - Enter doctor registration number
#    - Enter doctor qualification
#    - Enter doctor contact number
#    - Click "Add" to save care team member
#
# 10. SUBMIT FOR PRE-AUTHORIZATION
#     - Click "Preview & Validate" button
#     - Click "Validate" button
#     - Click "Initiate Pre-Authorization" button
#     - Confirm "YES" in confirmation dialog
#     - Wait for submission confirmation
#
#
# PHASE 5: REFERRAL CASE PROCESSING
# ────────────────────────────────────────────────────────────────────────────
# Repeat Phase 4 steps with these modifications:
#
# 1. SKIP EMG FORM GENERATION (uses existing referral form)
#
# 2. FILL REFERRAL DETAILS SECTION (instead of referral manual):
#    - Select referral radio button
#    - Enter referral number
#    - Click search button
#    - Enter intimation remarks
#    - Upload referral form PDF
#
# 3. All other steps follow same pattern as emergency cases
#
#
# KEY DATA STRUCTURES:
# ────────────────────────────────────────────────────────────────────────────
#
# MAIN EXCEL FILE COLUMNS USED:
#   - Regn. No. (Registration/IP number) - unique identifier
#   - CARD NO (Insurance card - emergency cases)
#   - REFERRAL NO (Referral ID - referral cases)
#   - Bill Date (format: YYYY-MM-DD or DD-MM-YYYY)
#   - Bill Time (format: HH:MM AM/PM)
#   - Patient Name
#   - Primary Doctor (may contain multiple names separated by "/")
#   - COMPLAINTS (chief complaints)
#   - DIAGNOSE (diagnosis code/description)
#   - Claim No (claim identifier)
#
# OUTPUT COLUMNS (Auto-populated):
#   - Contact No
#   - Address
#   - Doctor Name
#   - Ward Type
#
# GENERATED FILES:
#   - {IP_NUMBER}_emg.pdf - Emergency intimation form
#   - {IP_NUMBER}_card.pdf - Insurance card (pre-existing)
#   - {IP_NUMBER}_photo.jpeg - Patient photo (pre-existing)
#   - {IP_NUMBER}_reffphoto.pdf - Referral document (pre-existing)
#
#
# ERROR HANDLING STRATEGY:
# ────────────────────────────────────────────────────────────────────────────
#   - MedNet extraction errors: Skip patient and log, continue with next
#   - Form filling errors: Try alternate XPATH (DOM structure variations), skip field if both fail
#   - Critical failures: Sleep 24 hours to allow manual intervention
#   - File operations: Attempt deletion, log if fails (non-blocking)
#
#
# TECHNICAL IMPLEMENTATION DETAILS:
# ────────────────────────────────────────────────────────────────────────────
#
# WEB AUTOMATION TOOLS:
#   - Selenium WebDriver for browser automation
#   - WebDriverWait for element synchronization
#   - PyAutoGUI for file dialog input
#
# DATA PROCESSING:
#   - Pandas for Excel operations
#   - Python-DOCX for Word document manipulation
#   - Docx2PDF for document conversion
#
# TIMING CONSIDERATIONS:
#   - 5-10 second waits between major actions
#   - 0.5-2 second waits for page transitions
#   - 1-2 second pauses for file operations
#   - 24-hour (86400s) timeout on critical errors
#
# ELEMENT INTERACTION STRATEGY:
#   - Dynamic XPath selection with multiple fallbacks
#   - Try/except blocks for each field (graceful degradation)
#   - Date/time parsing from multiple format variations
#   - Dropdown selection via keyboard input when possible
#
#
# PROCESS STATISTICS:
# ────────────────────────────────────────────────────────────────────────────
#   - Per Emergency Case: ~5-6 minutes (form filling + submission)
#   - Per Referral Case: ~4-5 minutes (submission)
#   - Batch Processing: Sequential (one case at a time, waits between cases)
#
#
# ================================================================ SOURCE CODE ================================================================

from docx import Document
from docx2pdf import convert
from datetime import date
from docx.shared import Pt
import pandas as pd
import time, os, pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import win32gui
import win32con

driver = webdriver.Chrome()

pcuser = os.getlogin()
df_main = pd.read_excel(f"c:/Users/{pcuser}/Documents/CGHS_INTIMATION_EXCEL.xlsx")
main_excel = pd.DataFrame(df_main)
df_doc = pd.read_excel(f"C:/Users/{pcuser}/Documents/INTIMATION_BOT_DATA_DONT_EDIT/Doctors Details.xlsx")
doc_details = pd.DataFrame(df_doc)
df_cred = pd.read_excel(f"C:/Users/{pcuser}/Documents/INTIMATION_BOT_DATA_DONT_EDIT/mednet_cred.xlsx")
mednet_cred = pd.DataFrame(df_cred)
print(">>> Loaded Main Excels !!!")
main_excel['Contact No']=None
main_excel['Address']=None
main_excel['Doctor Name']=None
main_excel['Ward Type']=None # critical / ward
print(">>> Added the new columns for contact num, address, doctor name, ward type\n\n")

final_emg_cases_index = [0,1]
final_ref_cases_index = []


def main():

    try:
        #fetchdatamednet()
        OpenCGHS()
        for i in range(len(final_emg_cases_index)):
            temp_index = final_emg_cases_index[i]
            print(F"\n\n>>> WORKING FOR INDEX - [{temp_index}] - [{main_excel.loc[temp_index, "Regn. No.          "]}]\n\n")
            emg_case(temp_index)
        print("\n\n NOW INTIMATION FOR REFERRAL CASES \n\n")
        for i in range(len(final_ref_cases_index)):
            temp_index = final_ref_cases_index[i]
            print(F"\n\n>>> WORKING FOR INDEX - [{temp_index}] - [{main_excel.loc[temp_index, "Regn. No.          "]}]\n\n")
            ref_case(temp_index)
    except Exception as e:
        print(e)
        
        time.sleep(86400)


# ======= EMG FORM FILLING PART ========


def fetchdatamednet():
    driver.get("http://10.150.65.40/Login.jsp")
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="j_username"]')))
    enter_id_btn = driver.find_element(By.XPATH, '//*[@id="j_username"]')
    enter_id_btn.send_keys(mednet_cred.loc[0]['ID'])
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="j_password"]')))
    enter_pass_btn = driver.find_element(By.XPATH, '//*[@id="j_password"]')
    enter_pass_btn.send_keys(mednet_cred.loc[0]['pass'])
    login_btn = driver.find_element(By.XPATH, '//*[@id="submitButton"]')
    login_btn.click()
    print(">>> Login Succesfull")
    time.sleep(5)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menuDept"]'))).click()
    time.sleep(1)
    if ("activesublabel" in driver.find_element(By.XPATH, '//*[@id="RECEPTION"]').get_attribute("class").lower()):
        print(">>> Inpatient displayed...")
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[1]/div[3]/div[2]/div/div[1]/div[2]'))).click()
        inpatientbtn = WebDriverWait(driver , 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="leftNavigation"]/li[2]/a/div[2]')))
        inpatientbtn.click()
        print(">>> Opened Inpatient")
    else:
        print(">>> Inpatient Not Displayed...")
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="RECEPTION"]'))).click()
        print(">>> Opened Reception")
        time.sleep(0.5)
        #WebDriverWait(driver , 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="leftNavigation"]/li[2]/a/div[2]')))
        WebDriverWait(driver , 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="leftNavigation"]/li[2]/a/div[2]'))).click()
        print(">>> Opened Inpatient")
        time.sleep(5)

    for i in range(len(main_excel)):
        if (pd.isna(main_excel.loc[i, 'CARD NO']) and pd.isna(main_excel.loc[i, 'REFERRAL NO'])):
            main_excel.loc[i, 'CARD NO'] = None
            print(f">>> Skipping Patient {main_excel.loc[i,'Regn. No.          ']}")
        else:
            print(f">>> Fetching Data For Patient {main_excel.loc[i,'Regn. No.          ']}")

            if not pd.isna(main_excel.loc[i,'CARD NO']):
                final_emg_cases_index.append(i)
            if not pd.isna(main_excel.loc[i,'REFERRAL NO']):
                final_ref_cases_index.append(i)

            ip = main_excel.loc[i,'Regn. No.          ']
            datesplit = str(main_excel.loc[i," Bill Date           "]).split("-")
            if(len(datesplit[0]) == 4):
                # datesplit = datesplit[2].split(" ")
                date = str(int(datesplit[2]))
                month = str(int(datesplit[1])-1)
            else:
                date = str(int(datesplit[0]))
                month = str(int(datesplit[1])-1)

            WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[2]/div[2]/div[3]/div/div[2]/div[1]"))).click()
            Select(driver.find_element(By.XPATH,"/html/body/div[7]/div/div/label[1]/select")).select_by_value(month)
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"//td[normalize-space()='{date}']"))).click()
            

            ipinputbox = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="searchByRegnNo"]')))
            ipinputbox.clear()
            ipinputbox.send_keys(ip)
            # WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[2]/div[2]/div[2]/div/div[6]/div[1]/div/div/input'))).clear()
            print(">>> Opening Bill...")
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[2]/div[2]/div[4]/div[2]/div/div[1]/div[1]/div[1]/div[1]/div'))).click()
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[2]/div[2]/div[3]/div[1]/div/div[1]/div[6]/div[2]/div[5]/div'))).click()
            time.sleep(2)
            print(">>> Opened Bill...")
            # fetching contact num
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="patientHeaderDIV"]/div[2]/div[3]/div[2]')))
            main_excel.loc[i,'Contact No'] = driver.find_element(By.XPATH, '//*[@id="patientHeaderDIV"]/div[2]/div[3]/div[2]').text
            # fetching address 
            main_excel.loc[i,'Address'] = driver.find_element(By.XPATH, '//*[@id="patientHeaderDIV"]/div[2]/div[1]/div[2]/div[2]').text
            # fetching Doctor Name 
            main_excel.loc[i,'Doctor Name'] = (driver.find_element(By.XPATH, '//*[@id="consultantHeader"]').text).lower().title()
            # fetching Ward Type
            if any(x in driver.find_element(By.XPATH, '//*[@id="consultantHeader"]').text.lower() for x in ['ward', 'twin', 'single']):
                main_excel.loc[i,'Ward Type'] = 'ward'
            else:
                main_excel.loc[i,'Ward Type'] = 'critical'
            WebDriverWait(driver , 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="leftNavigation"]/li[2]/a/div[2]'))).click()
            
            if not pd.isna(main_excel.loc[i,'CARD NO']):
                makeemgform(i)
                print(f"\n>>> MADE EMG FORM FOR {main_excel.loc[i, "Regn. No.          "]}\n")

    main_excel.to_excel(f'c:/Users/{pcuser}/Documents/CGHS_INTIMATION_EXCEL.xlsx', index=False)


def makeemgform(index): # name-line6, cardnum-line7, contact-line8, address-line9, IPnum-line10, date=line11, 
                   # time-line12, complaints-line13, diagnosis-line15, docname-line24, doc_designation-line25, date_today-line26
    line = 1
    doc = Document(f"C:/Users/{pcuser}/Documents/INTIMATION_BOT_DATA_DONT_EDIT/cghs_emg_form.docx")
    for para in doc.paragraphs:
        #print(para.text,"\n", line,"\n\n")
        if line==6:
            para.clear()
            run = para.add_run(f"Patient Name: {main_excel.loc[index, " Patient Name          "]}")
            run.bold = True
            run.font.size = Pt(13)
        if line==7:
            para.clear()
            run = para.add_run(f"Insurance No. : {int(main_excel.loc[index, "CARD NO"])}")
            run.bold = True
            run.font.size = Pt(13)
        if line==8:
            para.clear()
            run = para.add_run(f"Contact No. : {main_excel.loc[index, "Contact No"]}")
            run.bold = True
            run.font.size = Pt(13)
        if line==9:
            para.clear()
            run = para.add_run(f"Address : {main_excel.loc[index, "Address"]}")
            run.bold = True
            run.font.size = Pt(13)
        if line==10:
            para.clear()
            run = para.add_run(f"IP No. : {main_excel.loc[index, "Regn. No.          "]}")
            run.bold = True
            run.font.size = Pt(13)
        if line==11:
            para.clear()
            run = para.add_run(f"Date : {main_excel.loc[index, " Bill Date           "]}")
            run.bold = True
            run.font.size = Pt(13)
        if line==12:
            para.clear()
            run = para.add_run(f"Time : {main_excel.loc[index, " Bill Time           "]}")
            run.bold = True
            run.font.size = Pt(13)
        if line==13:
            para.clear()
            run = para.add_run(f"Chief Complaints : {main_excel.loc[index, "COMPLAINTS"]}")
            run.bold = True
            run.font.size = Pt(13)
        if line==15:
            para.clear()
            run = para.add_run(f"Diagnosis : {main_excel.loc[index, "DIAGNOSE"]}")
            run.bold = True
            run.font.size = Pt(13)
        if line==24:
            para.clear()
            run = para.add_run(f"Name : {main_excel.loc[index, " Primary Doctor          "]}")
            run.bold = True
            run.font.size = Pt(13)
        if line==25:
            para.clear()
            doctors = main_excel.loc[index, " Primary Doctor          "].split("/")
            print(doctors[0])
            run = para.add_run(f"Desgination : {find_designation_doc(doctors[0].lower().title())}")
            run.bold = True
            run.font.size = Pt(13)
        if line==26:
            para.clear()
            run = para.add_run(f"Date : {date.today()}")
            run.bold = True
            run.font.size = Pt(13)

        line+=1
    doc.save(f"C:/Users/{pcuser}/Documents/{main_excel.loc[index, "Regn. No.          "]}_emg.docx")
    convert(f"C:/Users/{pcuser}/Documents/{main_excel.loc[index, "Regn. No.          "]}_emg.docx", f"C:/Users/{pcuser}/Documents/{main_excel.loc[index, "Regn. No.          "]}_emg.pdf")
    try:
        os.remove(f"C:/Users/{pcuser}/Documents/{main_excel.loc[index, "Regn. No.          "]}_emg.docx")
    except:
        print(f"Couldn't remove {main_excel.loc[index, "Regn. No.          "]}_emg.docx file..")

# ======== EMERGENCY INTIMATION PART =========


def find_designation_doc(doc_name):
    docs = doc_name.split("/")
    for doc in docs:
        row = doc_details.index[doc_details["Employee Name"].str.contains(doc, case=False, na=False)].to_list()
        if (row == []):
            continue
        else:
            return row[0]

def OpenCGHS():
    try:
        driver.get("https://provider.nha.gov.in")
        try:
            modal = WebDriverWait(driver, 500).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//div[contains(@class,'modal-content')]")
                )
            )
            modal.find_element(
                By.XPATH, ".//button[normalize-space()='CLOSE']"
            ).click()
        except Exception as e:
            print(">>> No Popup Found...")
    except:
        print(" !!! BOT STOPPED DUE TO ERROR !!!")
        time.sleep(86400)

def emg_case(index):
    cardloc = os.path.join("C:\\", "Users", f"{pcuser}", "Documents", f"{main_excel.loc[index, "Regn. No.          "]}_card.pdf")
    patientphotoloc = os.path.join("C:\\", "Users", f"{pcuser}", "Documents", f"{main_excel.loc[index, "Regn. No.          "]}_photo.jpeg")
    diag = str(main_excel.loc[index, "DIAGNOSE"])
    emgformloc = os.path.join("C:\\", "Users", f"{pcuser}", "Documents", f"{main_excel.loc[index, "Regn. No.          "]}_emg.pdf")
    docs = main_excel.loc[index, " Primary Doctor          "].split("/")
    for i in range(len(docs)):
        docname = docs[i].strip().title()
        docname = docname.replace("Dr. ", "").replace("Dr ", "")
        row = find_designation_doc(docname)
        if (row != None):
            break

    print(f">> Selected [{docname}] > row number - {row}")
    docregn = doc_details.loc[row, "Registration No."]
    docqualification = doc_details.loc[row, "Educational Qualification"]
    doccontact = str(doc_details.loc[row, "Contact No."])
    claimno = str(int(main_excel.loc[index, " Claim No           "]))
    ipnum = str(main_excel.loc[index, "Regn. No.          "])
    datesplit = str(main_excel.loc[index," Bill Date           "]).split("-")
    if(len(datesplit[0]) == 4):
        # datesplit = datesplit[2].split(" ")
        date = str(int(datesplit[2]))
    else:
        date = str(int(datesplit[0]))
    fulltimesplit = str(main_excel.loc[index, " Bill Time           "]).split(" ")
    ampm = fulltimesplit[1]
    timesplit = fulltimesplit[0].split(":")
    hr = timesplit[0]
    min = timesplit[1]


    try:
        element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div/div[2]/div[3]/div[1]/div[2]/div/div[4]/div/input')))
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
        element.click()
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div/div[2]/div[3]/div[1]/div[2]/div/div[4]/div/input'))).click()
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div/div[2]/div[3]/div[1]/div[2]/div/div[4]/div/input'))).send_keys(claimno)
        print(">> Entered ClaimNo.")
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Search']/preceding-sibling::span[contains(@class,'icon')]"))).click()
        time.sleep(0.5)
        print(">> Clicked On Search")
        time.sleep(1)
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[2]/div[3]/div[2]/div/div/div/div/div[3]/div/div[2]/p/something'))).click()
        print(">> Clicked On Patient")
        time.sleep(1)
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[1]/div/button"))).click()
            print(">> Opened Medical Info")
            time.sleep(1)
        except:
            print("First failed")
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[1]/div/button"))).click()
            print(">> Opened Medical Info")
    
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[1]/h2/button"))).click()
            print(">> Opened Personal Info")
        except:
            print("failed")
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[1]/h2/button"))).click()
            print(">> Opened Personal Info")
        
        # FILL PERSONAL INFO COLUMN :-
        try:
            # get save button state in personal info
            try:
                save_btn_state = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[1]/div/button'))).text
            except:
                save_btn_state = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[1]/div/button'))).text
            if ("EDIT" in str(save_btn_state)):
              # click edit button to start doing tasks
              try:
                WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
              except:
                WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
            print(f"current button state is: {save_btn_state}")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[3]/div[1]/fieldset/div[2]/label'))).click()
                print(">> Clicked NO")
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[3]/div[2]/fieldset/div[2]/label'))).click()
                print(">> Clicked NO")
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[3]/div[1]/fieldset/div[2]/label'))).click()
                print(">> Clicked NO")
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[3]/div[2]/fieldset/div[2]/label'))).click()
                print(">> Clicked NO")
        except:
            print("failed")

        # OPEN ADMISSION DETAILS COLUMN :-
        try:
            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[1]/div/button"))).click()
            except:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[1]/div/button'))).click()
            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/button'))).click()
            except:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/button'))).click()
            print(">> Opened Admission Information")
        except:
            print("!! Failed to Open Admission Info")
        
        try:
            WebDriverWait(driver, 0.5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[1]/h2/button'))).click()
            print("Opened Admission Details")
            time.sleep(1)
        except:
            print("Couldn't Open Admission Details")
            try:
                WebDriverWait(driver, 0.5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[1]/h2/button'))).click()
                print("Opened Admission Details")
                time.sleep(1)
            except:
                print("Couldn't Open Admission Details")

        # CHECK ADMSSION DETAILS SAVE BUTTON STATE:-
        try:
            save_btn_state = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).text
            if ("EDIT" in str(save_btn_state)):
                WebDriverWait(driver,15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
            print(f"current button state is: {save_btn_state}")
        except:
            print("Couldn't get button state")
            try:
                save_btn_state = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).text
                if("EDIT" in str(save_btn_state)):
                    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
                print(f"current button state is: {save_btn_state}")

            except:
                print("Couldn't get button state...")
                save_btn_state = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).text
                if("EDIT" in str(save_btn_state)):
                    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
                print(f"current button state is: {save_btn_state}")

        # FILL ADMISSION DETAILS COLUMN :-
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div/label'))).click()
            print("Opened Admission Date Bar")
            time.sleep(1)
            #print(WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div[2]/div[1]/p/text()[1]'))).text)
        except:
            print("Couldn't Open Admission Date Bar")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div/label'))).click()
                print("Opened Admission Date Bar")
                time.sleep(1)
                #print(WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div[2]/div[1]/p/text()[1]'))).text)
            except:
                print("Couldn't Open Admission Date Bar")

        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div[2]')))
            # date_drop_box = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div[2]')
            # Select(date_drop_box).select_by_value("2")
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"//button[normalize-space()={date} and not(@tabindex='-1')]"))).click()
        except:
            print("Couldn't set date")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div[2]')))
                # date_drop_box = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div[2]')
                # Select(date_drop_box).select_by_value("2")
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"//button[normalize-space()={date} and not(@tabindex='-1')]"))).click()
            except:
                print("Couldn't set date")
        
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[2]/div/div[1]/button'))).click()
            WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()={hr}]"))).click()
        except:
            print("Couldn't set hr")
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[2]/div/div[1]/button'))).click()
            WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()={hr}]"))).click()
        
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/button'))).click()
            WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()={min}]"))).click()
        except:
            print("Couldn't set min")
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/button'))).click()
            WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()={min}]"))).click()
        
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[2]/div/div[3]/button'))).click()
            if(ampm == "AM"):
                WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()='AM']"))).click()
            else:
                WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()='PM']"))).click()
            print("Set AM/PM")
        except:
            print("couldn't set am/pm")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[2]/div/div[3]/button'))).click()
                if(ampm == "AM"):
                    WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()='AM']"))).click()
                else:
                    WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()='PM']"))).click()
                print("Set AM/PM")
            except:
                print("Couldn't set AM/PM")

        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[2]/div[2]/div/div[1]/label'))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[2]/div[2]/div/div[1]/label'))).click()
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"//button[normalize-space()={date} and not(@tabindex='-1')]"))).click()
        except:
            print("Couldn't set surgery date")

        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[3]/fieldset/div[2]/label'))).click()
            print("Clicked No MLC")
        except:
            print("failed")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[3]/fieldset/div[2]/label'))).click()
                print("Clicked No MLC")
            except:
                print("Failed...")
        
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[4]/div/div/div/div[1]/div[2]'))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[4]/div/div/div/div[1]/div[2]'))).click()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[4]/div/div/div/div[1]/div[2]/input"))).send_keys("emer"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[4]/div/div/div/div[1]/div[2]/input"))).send_keys("emer"+Keys.ENTER)
        except:
            print("Couldn't select admission type")


        # OPEN TREATMENT COLUMN :-
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[2]/button"))).click()
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[2]/button"))).click()
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/button"))).click()
            print("Opened TREATMENT")
        except:
            print("Couldn't open TREATMENT")
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/button"))).click()
            print("Opened TREATMENT")

        # FILL REFERRAL MANUAL :-
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[1]/div/h2/button"))).click()
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[1]/div/h2/button"))).click()
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[1]/div/div[2]/div/div/div/div/form/div/div[2]/input"))).click()
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[1]/div/div[2]/div/div/div/div/form/div/div[2]/input"))).click()
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[1]/div[1]/div/input"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[1]/div[1]/div/input"))).click()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[1]/div[1]/div/input"))).send_keys(ipnum)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[1]/div[1]/div/input"))).send_keys(ipnum) 
        except:
            print("Couldn't input referal num")
        
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]"))).click()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("D70"+Keys.ENTER)
                #WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("D70"+Keys.ENTER)
                #WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
        except:
            print("Couldn't set wellness center")
        
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[1]/div[3]/div[2]/div/div/label"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[1]/div[3]/div[2]/div/div/label"))).click()
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"//button[normalize-space()='{date}' and not(@tabindex='-1')]"))).click()
        except:
            print("Couldn't set referral date")

        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[2]/div[1]/div/input"))).send_keys("emg")
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[2]/div[1]/div/input"))).send_keys("emg")

        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[2]/div[2]/div/input"))).send_keys("shrc")
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[2]/div[2]/div/input"))).send_keys("shrc")

        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[2]/div[3]/div/input"))).send_keys("Intimation")
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[2]/div[3]/div/input"))).send_keys("Intimation")
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[3]/div/fieldset/div/div[1]/button"))).click()
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[3]/div/fieldset/div/div[1]/button"))).click()
        # try:
        #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[3]/div/fieldset/div/div[1]/input"))).click()
        # except:
        #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[3]/div/fieldset/div/div[1]/input"))).click()
        time.sleep(1)
        pyautogui.write(emgformloc)
        time.sleep(1)
        pyautogui.press('enter')
        

        # FILL DIAGNOSIS
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/h2/button"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/h2/button"))).click()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[1]/div/input"))).send_keys(diag)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[1]/div/input"))).send_keys(diag)
            try:
                index0disease = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div/div/div[2]/div[2]/div/div/div/ul/li[1]/div/div[1]/div[2]/span[1]/span"))).click()
                #ActionChains(driver).move_to_element(index0disease).click(index0disease).perform()
            except:
                index1disease = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div/div/div[2]/div[2]/div/div/div/ul/li[1]/div/div[1]/div[2]/span[1]/span"))).click()
                #ActionChains(driver).move_to_element(index1disease).click(index1disease).perform()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[4]/button"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[4]/button"))).click()

            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[1]/div/input"))).send_keys(diag)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[1]/div/input"))).send_keys(diag)

            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div/div/div[2]/div[2]/div/div/div/ul/li[2]/div/div[1]/div[2]/span[1]/span"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div/div/div[2]/div[2]/div/div/div/ul/li[2]/div/div[1]/div[2]/span[1]/span"))).click()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[4]/button"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[4]/button"))).click()

        except:
            print("Couldn't add diagnosis")

        # OPEN TREATMENT PLAN COLUMN
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[1]/div[2]/h2/button"))).click()
            print("Opened Treatment Plan")
        except:
            print("Couldn't open treatment plan")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[1]/div[2]/h2/button"))).click()
                print("Opened Treatment Plan")
            except:
                print("Couldn't Open Treatment Plan")
                try:
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[1]/div[2]/h2/button"))).click()
                except:
                    print("Couldn't open treatment plan")
        
        # ADD WARD TYPE PLAN
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("critical"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("critical"+Keys.ENTER)
            # try:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            # except:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("icu"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("icu"+Keys.ENTER)
            # try:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            # except:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).clear()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).clear()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).send_keys("7")
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).send_keys("7")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[4]/div[3]/img"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[4]/div[3]/img"))).click()
        except:
            print("Couldn't add treatment ward plan")

        # ADD CONSULTATION :-
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]"))).click()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("consult"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("consult"+Keys.ENTER)
            # try:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            # except:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]"))).click()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("Inp"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("Inp"+Keys.ENTER)
            # try:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            # except:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).clear()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).clear()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).send_keys("18")
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).send_keys("18")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[4]/div[3]/img"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[4]/div[3]/img"))).click()
        except:
            print("Couldn't add treatment consultation plan")

        
        # ENTER CARD AND PHOTO FILES :-
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[4]/div/div[1]/div/h2/button"))).click()
                print("Opened investigation/documents")
            except:
                print("Couldn't Open Investigations/documents")
                try:
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[4]/div/div[1]/div/h2/button"))).click()
                    print("Opened Investigation/documents")
                except:
                    print("Couldn't Open Investigations/documents")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[4]/div/div[2]/div/table/tbody/tr[1]/td[4]/fieldset/div/div[1]/button"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[4]/div/div[2]/div/table/tbody/tr[1]/td[4]/fieldset/div/div[1]/button"))).click()
            time.sleep(1)
            pyautogui.write(cardloc)
            time.sleep(1)
            pyautogui.press('enter')
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[4]/div/div[2]/div/table/tbody/tr[2]/td[4]/fieldset/div/div[1]/button"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[4]/div/div[2]/div/table/tbody/tr[2]/td[4]/fieldset/div/div[1]/button"))).click()
            time.sleep(1)
            pyautogui.write(patientphotoloc)
            time.sleep(1)
            pyautogui.press('enter')
        except:
            print("Couldn't Fill Investigations/Documents")
        
        
        # FILL CARE TEAM DETAILS :-
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[1]/div/h2/button"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[1]/div/h2/button"))).click()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[1]/div/div[2]/div/div/div/div/div[1]/div/div/div/div[1]/div[2]/input"))).send_keys("other"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[1]/div/div[2]/div/div/div/div/div[1]/div/div/div/div[1]/div[2]/input"))).send_keys("other"+Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[1]/div/input"))).send_keys(docname+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[1]/div/input"))).send_keys(docname+Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[2]/div/input"))).send_keys(docregn+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[2]/div/input"))).send_keys(docregn+Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[3]/div/div/div/div[1]/div[2]/input"))).send_keys(docqualification+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[3]/div/div/div/div[1]/div[2]/input"))).send_keys(docqualification+Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[4]/div"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[4]/div"))).click()
            time.sleep(1.5)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[4]/div/input"))).send_keys(doccontact+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[4]/div/input"))).send_keys(doccontact+Keys.ENTER)
            try:
                WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[1]/div/div[2]/div/div/div/div/div[2]/button"))).click()
            except:
                WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[1]/div/div[2]/div/div/div/div/div[2]/button"))).click()
        
        except:
            print("Couldn't fill care team details")

        # PREVIEW AND VALIDATE...
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Preview & Validate']"))).click()
        time.sleep(1)
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='validate']"))).click()
            time.sleep(1)
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Initiate Pre-Authorization']"))).click()
            time.sleep(1)
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='YES']"))).click()
            time.sleep(1)
            print(f"\n>>> Patient {ipnum} process completed... !!!\n")
            time.sleep(8)
        except:
            print("!!! Patient IP358516 process couldn't complete due to error... !!!")
            print("!!! Going to home\n")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[1]/div/div[1]/button[1]"))).click()
            except:
                try:
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[2]/div/div[1]/button[1]"))).click()
                except:    
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div/div[1]/button[1]"))).click()
            
            file = f"c:/Users/{pcuser}/Documents/CGHS_INTIMATION_EXCEL.xlsx"
            wb = load_workbook(file)
            ws = wb.active

            red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

            # Color full first row (Excel row 1)
            for cell in ws[index+2]:
                cell.fill = red_fill

            wb.save(file)

            print(f"!!! Coloured index {index} in Red")
            time.sleep(8)

    except Exception as e:
        print(">>>> BOT STOPPED FOR 24HRS... CHECK MANUALLY...!!!!")
        print(e)
        time.sleep(86400)

# ========= REFERRAL INTIMATION PART ==========

def ref_case(index):
    cardloc = os.path.join("C:\\", "Users", f"{pcuser}", "Documents", f"{main_excel.loc[index, "Regn. No.          "]}_card.pdf")
    patientphotoloc = os.path.join("C:\\", "Users", f"{pcuser}", "Documents", f"{main_excel.loc[index, "Regn. No.          "]}_photo.jpeg")
    diag = str(main_excel.loc[index, "DIAGNOSE"])
    reffphotoloc = os.path.join("C:\\", "Users", f"{pcuser}", "Documents", f"{main_excel.loc[index, "Regn. No.          "]}_reffphoto.pdf")
    docs = main_excel.loc[index, " Primary Doctor          "].split("/")
    for i in range(len(docs)):
        docname = docs[i].strip().title()
        docname = docname.replace("Dr. ", "").replace("Dr ", "")
        row = find_designation_doc(docname)
        if (row != None):
            break

    print(f">> Selected [{docname}] > row number - {row}")
    docregn = doc_details.loc[row, "Registration No."]
    docqualification = doc_details.loc[row, "Educational Qualification"]
    doccontact = str(doc_details.loc[row, "Contact No."])
    claimno = str(int(main_excel.loc[index, " Claim No           "]))
    ipnum = str(main_excel.loc[index, "Regn. No.          "])
    reffnum = str(main_excel.loc[index, "REFERRAL NO"])
    datesplit = str(main_excel.loc[index," Bill Date           "]).split("-")
    if(len(datesplit[0]) == 4):
        # datesplit = datesplit[2].split(" ")
        date = str(int(datesplit[2]))
    else:
        date = str(int(datesplit[0]))
    fulltimesplit = str(main_excel.loc[index, " Bill Time           "]).split(" ")
    ampm = fulltimesplit[1]
    timesplit = fulltimesplit[0].split(":")
    hr = timesplit[0]
    min = timesplit[1]

    try:
        element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div/div[2]/div[3]/div[1]/div[2]/div/div[4]/div/input')))
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
        element.click()
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div[4]/div/input'))).click()
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div[4]/div/input'))).send_keys(claimno)
        print(">> Entered ClaimNo.")
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Search']/preceding-sibling::span[contains(@class,'icon')]"))).click()
        time.sleep(0.5)
        print(">> Clicked On Search")
        time.sleep(1)
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[2]/div[3]/div[2]/div/div/div/div/div[3]/div/div[2]/p/something'))).click()
        print(">> Clicked On Patient")
        time.sleep(1)
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[1]/div/button"))).click()
            print(">> Opened Medical Info")
            time.sleep(1)
        except:
            print("First failed")
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[1]/div/button"))).click()
            print(">> Opened Medical Info")
    
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[1]/h2/button"))).click()
            print(">> Opened Personal Info")
        except:
            print("failed")
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[1]/h2/button"))).click()
            print(">> Opened Personal Info")
        
        # FILL PERSONAL INFO COLUMN :-
        try:
            # get save button state in personal info
            try:
                save_btn_state = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[1]/div/button'))).text
            except:
                save_btn_state = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[1]/div/button'))).text
            if ("EDIT" in str(save_btn_state)):
              # click edit button to start doing tasks
              try:
                WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
              except:
                WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
            print(f"current button state is: {save_btn_state}")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[3]/div[1]/fieldset/div[2]/label'))).click()
                print(">> Clicked NO")
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[3]/div[2]/fieldset/div[2]/label'))).click()
                print(">> Clicked NO")
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[3]/div[1]/fieldset/div[2]/label'))).click()
                print(">> Clicked NO")
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[1]/div/div/div/div/div/div/div/div/div/div[3]/div/div[2]/div/form/div/div/div[3]/div[2]/fieldset/div[2]/label'))).click()
                print(">> Clicked NO")
        except:
            print("failed")

        # OPEN ADMISSION DETAILS COLUMN :-
        try:
            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[1]/div/button"))).click()
            except:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[1]/div/button'))).click()
            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/button'))).click()
            except:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/button'))).click()
            print(">> Opened Admission Information")
        except:
            print("!! Failed to Open Admission Info")
        
        try:
            WebDriverWait(driver, 0.5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[1]/h2/button'))).click()
            print("Opened Admission Details")
            time.sleep(1)
        except:
            print("Couldn't Open Admission Details")
            try:
                WebDriverWait(driver, 0.5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[1]/h2/button'))).click()
                print("Opened Admission Details")
                time.sleep(1)
            except:
                print("Couldn't Open Admission Details")

        # CHECK ADMSSION DETAILS SAVE BUTTON STATE:-
        try:
            save_btn_state = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).text
            if ("EDIT" in str(save_btn_state)):
                WebDriverWait(driver,15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
            print(f"current button state is: {save_btn_state}")
        except:
            print("Couldn't get button state")
            try:
                save_btn_state = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).text
                if("EDIT" in str(save_btn_state)):
                    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
                print(f"current button state is: {save_btn_state}")

            except:
                print("Couldn't get button state...")
                save_btn_state = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).text
                if("EDIT" in str(save_btn_state)):
                    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
                print(f"current button state is: {save_btn_state}")

        # FILL ADMISSION DETAILS COLUMN :-
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div/label'))).click()
            print("Opened Admission Date Bar")
            time.sleep(1)
            #print(WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div[2]/div[1]/p/text()[1]'))).text)
        except:
            print("Couldn't Open Admission Date Bar")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div/label'))).click()
                print("Opened Admission Date Bar")
                time.sleep(1)
                #print(WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div[2]/div[1]/p/text()[1]'))).text)
            except:
                print("Couldn't Open Admission Date Bar")

        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div[2]')))
            # date_drop_box = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div[2]')
            # Select(date_drop_box).select_by_value("2")
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"//button[normalize-space()={date} and not(@tabindex='-1')]"))).click()
        except:
            print("Couldn't set date")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div[2]')))
                # date_drop_box = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[1]/div[2]')
                # Select(date_drop_box).select_by_value("2")
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"//button[normalize-space()={date} and not(@tabindex='-1')]"))).click()
            except:
                print("Couldn't set date")
        
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[2]/div/div[1]/button'))).click()
            WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()={hr}]"))).click()
        except:
            print("Couldn't set hr")
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[2]/div/div[1]/button'))).click()
            WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()={hr}]"))).click()
        
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/button'))).click()
            WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()={min}]"))).click()
        except:
            print("Couldn't set min")
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/button'))).click()
            WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()={min}]"))).click()
        
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[2]/div/div[3]/button'))).click()
            if(ampm == "AM"):
                WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()='AM']"))).click()
            else:
                WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()='PM']"))).click()
            print("Set AM/PM")
        except:
            print("couldn't set am/pm")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[1]/div[2]/div[2]/div/div[3]/button'))).click()
                if(ampm == "AM"):
                    WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()='AM']"))).click()
                else:
                    WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, f"//ul[contains(@class,'A0PbSUeMish')]//li[normalize-space()='PM']"))).click()
                print("Set AM/PM")
            except:
                print("Couldn't set AM/PM")

        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[2]/div[2]/div/div[1]/label'))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[2]/div[2]/div/div[1]/label'))).click()
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"//button[normalize-space()={date} and not(@tabindex='-1')]"))).click()
        except:
            print("Couldn't set surgery date")

        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[3]/fieldset/div[2]/label'))).click()
            print("Clicked No MLC")
        except:
            print("failed")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[3]/fieldset/div[2]/label'))).click()
                print("Clicked No MLC")
            except:
                print("Failed...")
        
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[4]/div/div/div/div[1]/div[2]'))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[4]/div/div/div/div[1]/div[2]'))).click()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[4]/div/div/div/div[1]/div[2]/input"))).send_keys("emer"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[2]/div[4]/div/div/div/div[1]/div[2]/input"))).send_keys("emer"+Keys.ENTER)
        except:
            print("Couldn't select admission type")


        # OPEN TREATMENT COLUMN :-
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[2]/button"))).click()
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[2]/button"))).click()
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/button"))).click()
            print("Opened TREATMENT")
        except:
            print("Couldn't open TREATMENT")
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/button"))).click()
            print("Opened TREATMENT")



        # FILL REFERRAL DETAILS
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[1]/div/h2/button"))).click()
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[1]/div/h2/button"))).click()
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[1]/div/div[2]/div/div/div/div/form/div/div[3]/input"))).click()
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[1]/div/div[2]/div/div/div/div/form/div/div[3]/input"))).click()
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div/div[1]/div/input"))).send_keys(reffnum)
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div/div[1]/div/input"))).send_keys(reffnum)
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div/div[2]/button"))).click()
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div/div[2]/button"))).click()
        time.sleep(1)
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div[1]/div[2]/div[3]/div/input"))).send_keys("intimation")
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div[1]/div[2]/div[3]/div/input"))).send_keys("intimation")
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div[1]/div[3]/div/fieldset/div/div[1]/button"))).click()
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div[1]/div[3]/div/fieldset/div/div[1]/button"))).click()

        time.sleep(1)
        pyautogui.write(reffphotoloc)
        time.sleep(1)
        pyautogui.press('enter')
       
       
        # FILL DIAGNOSIS
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/h2/button"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/h2/button"))).click()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[1]/div/input"))).send_keys(diag)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[1]/div/input"))).send_keys(diag)
            try:
                index0disease = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div/div/div[2]/div[2]/div/div/div/ul/li[1]/div/div[1]/div[2]/span[1]/span"))).click()
                #ActionChains(driver).move_to_element(index0disease).click(index0disease).perform()
            except:
                index1disease = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div/div/div[2]/div[2]/div/div/div/ul/li[1]/div/div[1]/div[2]/span[1]/span"))).click()
                #ActionChains(driver).move_to_element(index1disease).click(index1disease).perform()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[4]/button"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[4]/button"))).click()

            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[1]/div/input"))).send_keys(diag)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[1]/div/input"))).send_keys(diag)

            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div/div/div[2]/div[2]/div/div/div/ul/li[2]/div/div[1]/div[2]/span[1]/span"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div/div/div[2]/div[2]/div/div/div/ul/li[2]/div/div[1]/div[2]/span[1]/span"))).click()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[4]/button"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div[4]/button"))).click()

        except:
            print("Couldn't add diagnosis")

        # OPEN TREATMENT PLAN COLUMN
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[1]/div[2]/h2/button"))).click()
            print("Opened Treatment Plan")
        except:
            print("Couldn't open treatment plan")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[1]/div[2]/h2/button"))).click()
                print("Opened Treatment Plan")
            except:
                print("Couldn't Open Treatment Plan")
                try:
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[1]/div[2]/h2/button"))).click()
                except:
                    print("Couldn't open treatment plan")
        
        # ADD WARD TYPE PLAN
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("critical"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("critical"+Keys.ENTER)
            # try:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            # except:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("icu"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("icu"+Keys.ENTER)
            # try:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            # except:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).clear()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).clear()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).send_keys("7")
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).send_keys("7")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[4]/div[3]/img"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[4]/div[3]/img"))).click()
        except:
            print("Couldn't add treatment ward plan")

        # ADD CONSULTATION :-
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]"))).click()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("consult"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("consult"+Keys.ENTER)
            # try:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            # except:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]"))).click()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("Inp"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("Inp"+Keys.ENTER)
            # try:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            # except:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).clear()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).clear()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).send_keys("18")
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[3]/div[2]/div/input"))).send_keys("18")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[4]/div[3]/img"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[4]/div[3]/img"))).click()
        except:
            print("Couldn't add treatment consultation plan")

        
        # ENTER CARD AND PHOTO FILES :-
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[4]/div/div[1]/div/h2/button"))).click()
                print("Opened investigation/documents")
            except:
                print("Couldn't Open Investigations/documents")
                try:
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[4]/div/div[1]/div/h2/button"))).click()
                    print("Opened Investigation/documents")
                except:
                    print("Couldn't Open Investigations/documents")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[4]/div/div[2]/div/table/tbody/tr[1]/td[4]/fieldset/div/div[1]/button"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[4]/div/div[2]/div/table/tbody/tr[1]/td[4]/fieldset/div/div[1]/button"))).click()
            time.sleep(1)
            pyautogui.write(cardloc)
            time.sleep(1)
            pyautogui.press('enter')
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[4]/div/div[2]/div/table/tbody/tr[2]/td[4]/fieldset/div/div[1]/button"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[4]/div/div[2]/div/table/tbody/tr[2]/td[4]/fieldset/div/div[1]/button"))).click()
            time.sleep(1)
            pyautogui.write(patientphotoloc)
            time.sleep(1)
            pyautogui.press('enter')
        except:
            print("Couldn't Fill Investigations/Documents")
        
        
        # FILL CARE TEAM DETAILS :-
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[1]/div/h2/button"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[1]/div/h2/button"))).click()
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[1]/div/div[2]/div/div/div/div/div[1]/div/div/div/div[1]/div[2]/input"))).send_keys("other"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[1]/div/div[2]/div/div/div/div/div[1]/div/div/div/div[1]/div[2]/input"))).send_keys("other"+Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[1]/div/input"))).send_keys(docname+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[1]/div/input"))).send_keys(docname+Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[2]/div/input"))).send_keys(docregn+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[2]/div/input"))).send_keys(docregn+Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[3]/div/div/div/div[1]/div[2]/input"))).send_keys(docqualification+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[3]/div/div/div/div[1]/div[2]/input"))).send_keys(docqualification+Keys.ENTER)
            time.sleep(1)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[4]/div"))).click()
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[4]/div"))).click()
            time.sleep(1.5)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[4]/div/input"))).send_keys(doccontact+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[4]/div/input"))).send_keys(doccontact+Keys.ENTER)
            try:
                WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[1]/div/div[2]/div/div/div/div/div[2]/button"))).click()
            except:
                WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[1]/div/div[2]/div/div/div/div/div[2]/button"))).click()

        except:
            print("Couldn't fill care team details")

        # PREVIEW AND VALIDATE...
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Preview & Validate']"))).click()
        time.sleep(1)
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='validate']"))).click()
            time.sleep(1)
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Initiate Pre-Authorization']"))).click()
            time.sleep(1)
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='YES']"))).click()
            time.sleep(1)
            print(f"\n>>> Patient {ipnum} process completed... !!!\n")
            time.sleep(8)
        except:
            print("!!! Patient IP358516 process couldn't complete due to error... !!!")
            print("!!! Going to home\n")
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[1]/div/div[1]/button[1]"))).click()
            except:
                try:
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[2]/div/div[1]/button[1]"))).click()
                except:    
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div/div[1]/button[1]"))).click()
            
            file = f"c:/Users/{pcuser}/Documents/CGHS_INTIMATION_EXCEL.xlsx"
            wb = load_workbook(file)
            ws = wb.active

            red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

            # Color full first row (Excel row 1)
            for cell in ws[index+2]:
                cell.fill = red_fill

            wb.save(file)
            print(f"!!! Coloured index {index} in Red")
            time.sleep(8)

    except Exception as e:
        print(">>>> BOT STOPPED FOR 24HRS... CHECK MANUALLY...!!!!")
        print(e)
        time.sleep(86400)

def maximize_terminal():
    hwnd = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

if __name__ == "__main__":
    main()
    print("===================== SUCCESSSS!! ====================")
    time.sleep(1800)
    


# ================================================================ DUMP-YARD ===============================================================

# def takingclaimnosfrommednet():
#     df = pd.read_excel(excel_file_loc)
#     print(df.head(), "\n\n\n")
#     df["claimids"] = "pending"
#     driver.get("http://10.150.65.44/Login.jsp")
#     # wait for user login
#     loginbtn = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="submitButton"]')))
#     print("> Waiting for user to login....")
#     WebDriverWait(driver, 86400).until(EC.staleness_of(loginbtn))
#     print("> User Logged In !!")
#     # PLS Close Password Popup If Comes
#     time.sleep(5)
#     # try:
#     #     WebDriverWait(driver, 5).until(EC.alert_is_present())
#     #     alert = driver.switch_to.alert
#     #     print("Alert text:", alert.text)
#     #     alert.accept()   # or alert.dismiss()
#     #     print(">>> JS alert handled")
#     # except Exception as e:
#     #     print(">>> No JS alert")
#     # click on menu
#     WebDriverWait(driver , 1000).until(EC.presence_of_element_located((By.XPATH, '//*[@id="menuDept"]')))
#     menu = WebDriverWait(driver , 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menuDept"]')))
#     menu.click()
#     # click on reception
#     WebDriverWait(driver , 100).until(EC.presence_of_element_located((By.XPATH, '//*[@id="RECEPTION"]')))
#     receptionbtn = WebDriverWait(driver , 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[1]/div[3]/div[2]/input')))
#     receptionbtn.click()
#     # click on inpatient
#     WebDriverWait(driver , 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="leftNavigation"]/li[2]/a/div[2]')))
#     inpatientbtn = WebDriverWait(driver , 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="leftNavigation"]/li[2]/a/div[2]')))
#     inpatientbtn.click()
#     # Enter every IPNO.
#     for i in range(len(df)):
#         ipinputbox = driver.find_element(By.XPATH, '//*[@id="searchByRegnNo"]')
#         ipinputbox.clear()
#         ipinputbox.send_keys(df.iloc[i,0])
#         time.sleep(1)
#         profilebtn = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="tableData"]/div/div[1]/div[2]/a/div')))
#         ActionChains(driver).move_to_element(profilebtn).perform()
#         notesbtn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tableData"]/div/div[1]/div[2]/div[2]/div')))
#         notesbtn.click()
#         regnid = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div[2]/div[2]/div[13]/div[5]/div[2]/div[1]')))
#         tempclaimidvar = regnid.text.split("-")[-1].replace(" ", "")
#         print("\n",tempclaimidvar, "\n")
#         claimidslist.append(tempclaimidvar)
#         time.sleep(0.5)
#         closenotesbtn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="notesPopOPD"]/div[1]/div[2]/div')))
#         closenotesbtn.click()
#         time.sleep(1)
    
#     # df["claimids"] = claimidslist
#     # df.to_excel(excel_file_loc, index=False)
