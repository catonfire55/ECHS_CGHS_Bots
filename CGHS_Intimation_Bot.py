# ============================================================================================================================================
#                                                             _____WORKING PLAN____
# Check Emergency & Referral Cases
# Fill Emergency Form
# Opening Site
#
# TASKS: 1) Medical Info - Personal Details ( no no dono me )
#        2) Admission Information - Admission Details ( from excel ) ( Check Planned(ref) or Emergency from the files in documents Folder )
#                                   > If IPNO_emg file in documents - emergency case - select manual in Treatment(referal details),
#                                           remarks - "intimation" , patient ip is referal number , wellness center code "70" and pres enter,
#                                           endorsment by - "emg" , edorsment hosp - "shrc" , upload doc ( IPNO_emg )
#                                           date of referal - date of admission
#
#                                   > if IPNO_refphoto in documents - referal case - select emanual in treatment(referal details),
#                                           wait for user to enter referal Number ( User should enter c to continue )
#
#        3) Diagnosis & Treatment Plan ( speciality - micu/ccu/icu )
#        4) 
# emg medical cert - from sys - name of hosp, addr of hosp, patient name, contatct num, patient addr, digital sign of doc, hosp auth stamp
#                       manual -  card num, diagnosis, presenting complaint, treatment plan
# in excel format - mlc, 
#
# ================================================================ Source Code ================================================================

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
from selenium.webdriver.support import expected_conditions as EC
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
final_emg_cases_index = []
final_ref_cases_index = []


def main():

    try:
        fetchdatamednet()
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
    driver.get("http://10.150.65.44/Login.jsp")
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
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="megaMenuPopupDiv"]/div/div[2]/div[1]/div[2]'))).click()
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
        if (pd.isna(main_excel.loc[i,'CARD NO']) and pd.isna(main_excel.loc[i, 'REFERRAL NO'])):
            main_excel.loc[i,'CARD NO'] = None
            print(f">>> Skipping Patient {main_excel.loc[i,'Regn. No.          ']}")
        else:
            print(f">>> Fetching Data For Patient {main_excel.loc[i,'Regn. No.          ']}")

            if not pd.isna(main_excel.loc[i,'CARD NO']):
                final_emg_cases_index.append(i)
            if not pd.isna(main_excel.loc[i,'REFERRAL NO']):
                final_ref_cases_index.append(i)

            ip = main_excel.loc[i,'Regn. No.          ']
            ipinputbox = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="searchByRegnNo"]')))
            ipinputbox.clear()
            ipinputbox.send_keys(ip)
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[2]/div[2]/div[2]/div/div[6]/div[1]/div/div/input'))).clear()
            print(">>> Opening Bill...")
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[2]/div[2]/div[3]/div[2]/div/div[1]/div[1]/div[1]/div[1]/div'))).click()
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
        row = (doc_details.index[doc_details["Employee Name"] == doc].to_list())
        if (row == []):
            continue
        else:
            return row[0]

def OpenCGHS():
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

def emg_case(index):
    cardloc = os.path.join("C:\\", "Users", f"{pcuser}", "Documents", f"{main_excel.loc[index, "Regn. No.          "]}_card.pdf")
    patientphotoloc = os.path.join("C:\\", "Users", f"{pcuser}", "Documents", f"{main_excel.loc[index, "Regn. No.          "]}_photo.jpeg")
    diag = str(main_excel.loc[index, "DIAGNOSE"])
    emgformloc = os.path.join("C:\\", "Users", f"{pcuser}", "Documents", f"{main_excel.loc[index, "Regn. No.          "]}_emg.pdf")
    docname = main_excel.loc[index, " Primary Doctor          "].split("/")[0].strip().title()
    row = find_designation_doc(docname)
    docregn = doc_details.loc[row, "Registration No."]
    docqualification = doc_details.loc[row, "Educational Qualification"]
    doccontact = doc_details.loc[row, "Contact No."]
    claimno = str(main_excel.loc[index, " Claim No           "])
    ipnum = str(main_excel.loc[index, "Regn. No.          "])
    datesplit = str(main_excel.loc[index," Bill Date           "]).split("-")
    if(len(datesplit[0]) == 4):
        datesplit = datesplit[2].split(" ")
        date = str(int(datesplit[0]))
    else:
        date = str(int(datesplit[0]))
    fulltimesplit = str(main_excel.loc[index, " Bill Time           "]).split(" ")
    ampm = fulltimesplit[1]
    timesplit = fulltimesplit[0].split(":")
    hr = timesplit[0]
    min = timesplit[1]


    try:
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
            save_btn_state = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).text
            if ("EDIT" in str(save_btn_state)):
                WebDriverWait(driver,15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
            print(f"current button state is: {save_btn_state}")
        except:
            print("Couldn't get button state")
            try:
                save_btn_state = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).text
                if("EDIT" in str(save_btn_state)):
                    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
                print(f"current button state is: {save_btn_state}")

            except:
                print("Couldn't get button state...")
                save_btn_state = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).text
                if("EDIT" in str(save_btn_state)):
                    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
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
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("consult"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("consult"+Keys.ENTER)
            # try:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            # except:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("Inp"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("Inp"+Keys.ENTER)
            # try:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            # except:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
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
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[4]/div/input"))).send_keys(doccontact+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[4]/div/input"))).send_keys(doccontact+Keys.ENTER)

        except:
            print("Couldn't fill care team details")

        # GO BACK TO HOME PAGE :-
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[1]/div/div[1]/button[1]"))).click()
            print(">>> Went back to home page...")
        except:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[2]/div/div[1]/button[1]"))).click()
            print(">>> Went back to home page...")
                

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
    docname = main_excel.loc[index, " Primary Doctor          "].split("/")[0].strip().title()
    row = find_designation_doc(docname)
    docregn = doc_details.loc[row, "Registration No."]
    docqualification = doc_details.loc[row, "Educational Qualification"]
    doccontact = doc_details.loc[row, "Contact No."]
    claimno = str(main_excel.loc[index, " Claim No           "])
    ipnum = str(main_excel.loc[index, "Regn. No.          "])
    reffnum = str(main_excel.loc[index, "REFERRAL NO"])
    datesplit = str(main_excel.loc[index," Bill Date           "]).split("-")
    if(len(datesplit[0]) == 4):
        datesplit = datesplit[2].split(" ")
        date = str(int(datesplit[0]))
    else:
        date = str(int(datesplit[0]))
    fulltimesplit = str(main_excel.loc[index, " Bill Time           "]).split(" ")
    ampm = fulltimesplit[1]
    timesplit = fulltimesplit[0].split(":")
    hr = timesplit[0]
    min = timesplit[1]

    try:
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
            save_btn_state = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).text
            if ("EDIT" in str(save_btn_state)):
                WebDriverWait(driver,15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
            print(f"current button state is: {save_btn_state}")
        except:
            print("Couldn't get button state")
            try:
                save_btn_state = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).text
                if("EDIT" in str(save_btn_state)):
                    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[6]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
                print(f"current button state is: {save_btn_state}")

            except:
                print("Couldn't get button state...")
                save_btn_state = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).text
                if("EDIT" in str(save_btn_state)):
                    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[7]/div[2]/div/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div/div[1]/div/button'))).click()
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
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("consult"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("consult"+Keys.ENTER)
            # try:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            # except:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("Inp"+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys("Inp"+Keys.ENTER)
            # try:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
            # except:
            #     WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/div/div[1]/div[2]/input"))).send_keys(Keys.ENTER)
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
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[7]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[4]/div/input"))).send_keys(doccontact+Keys.ENTER)
            except:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[6]/div[3]/div/div/div/div/div/div/div/div/div[5]/div/div/div[2]/div/div/div/div/form/div[1]/div[4]/div/input"))).send_keys(doccontact+Keys.ENTER)

        except:
            print("Couldn't fill care team details")

        # GO BACK TO HOME PAGE :-
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[1]/div/div[1]/button[1]"))).click()
            print(">>> Went back to home page...")
        except:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[2]/div/div[1]/button[1]"))).click()
            print(">>> Went back to home page...")

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
