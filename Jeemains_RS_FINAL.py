import logging
import oracledb
from csv import writer
from PIL import Image
from selenium import webdriver
from paddleocr import PaddleOCR
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import requests
from bs4 import BeautifulSoup
import oracledb
import numpy as np
import pandas as pd

v_process_user = 'D5'

logging.getLogger('ppocr').setLevel(logging.ERROR)
def extract_text_with_paddleocr(image_path):
    print(image_path)
    """
    Extract text from an image using PaddleOCR.
    """
    try:
        # Initialize PaddleOCR model with advanced settings
        ocr = PaddleOCR(
            use_angle_cls=True,  # Correct text rotation (useful for skewed or rotated text)
            lang='en',           # Set OCR language (English)
            rec=True,            # Enable text recognition (extract actual text)
            cls=True,            # Enable text-line classification (helpful for fragmented text)
            use_space_char=True, # Recognize spaces as characters to preserve word separation
            det=True,            # Enable detection of all text regions
            box_thresh=0.5,      # Threshold for detecting boxes with text
            rec_batch_num=10     # Batch size for text recognition (for better accuracy)
        )

        # Perform OCR on the image
        results = ocr.ocr(image_path, cls=True, rec=True, det=True)
        # Extract and combine the text
        extracted_text = "\n".join([line[1][0] for line in results[0]])

        return extracted_text
    except Exception as e:
        return ""
    
max_attempts = 20
attempts = 0
login_successful = False
    
#oracledb.init_oracle_client()
oracledb.init_oracle_client(lib_dir=r"D:\app\udaykumard\product\instantclient_23_6")
conn = oracledb.connect(user='RESULT', password='LOCALDEV', dsn='192.168.15.208:1521/orcldev')
cur = conn.cursor()

# driver = webdriver.Chrome()
# driver.maximize_window()

str_dataslot = "SELECT PROCESS_USER, START_VAL, END_VAL FROM DATASLOTS_VAL_USER WHERE PROCESS_USER = '"+v_process_user+"'"
cur.execute(str_dataslot)
res_dataslot = cur.fetchall()

start_sno = res_dataslot[0][1]
end_sno = res_dataslot[0][2]

Sno = start_sno

# str_Jeeappno = "SELECT 1 as sno, APPNO, PASSWORD FROM I_JEEMAINS_ADMITCARD_06FEB25 WHERE PROCESS_STATUS = 'P' AND APPNO = '250310417958'"
# cur.execute(str_Jeeappno)
# res = cur.fetchall()

str_Jeeappno = "SELECT SLNO, APPNO, PASSWORD FROM \
( \
    SELECT SNO AS SLNO, TRIM(APPNO) APPNO, TRIM(PASSWORD) PASSWORD FROM I_JEEMAINS_ADMITCARD_06FEB25 \
    WHERE PROCESS_STATUS = 'P' AND PASSWORD IS NOT NULL AND LENGTH(APPNO) = 12 \
) \
WHERE SLNO >= '"+str(start_sno)+"' AND  SLNO <='"+str(end_sno)+"' ORDER BY SLNO"
cur.execute(str_Jeeappno)
res = cur.fetchall()

def QA_details():
    href_link = driver.find_element(By.ID, "ctl00_LoginContent_rptViewQuestionPaper_ctl01_lnkviewKey")
    url_link = href_link.get_attribute("href")

    driver_qa = webdriver.Chrome()
    driver_qa.maximize_window()
    driver_qa.get(url_link)

    payload = {}
    headers = {}

    response = requests.request("POST", url_link, headers=headers, data=payload)
    soup = BeautifulSoup(response.text, 'html.parser')
    div_text = soup.find("div", {"class": "main-info-pnl"})
    table_data = div_text.find('table').find_all('tr')

    data = []
    for row in table_data:
        row_data = []
        for cell in row.find_all('td'):
            row_data.append(cell.text)
        data.append(row_data)

    o_ApplicationNo = data[0][1]
    o_CandidateName = data[1][1]
    o_RollNo = data[2][1]
    o_TestDate = data[3][1]
    o_TestTime = data[4][1]
    o_Subject = data[5][1]

    sno = 1
    rows_list = []
    df = pd.DataFrame(columns=('QA_ID','APPNO', 'QUESTION_TYPE', 'QUESTION_ID', 'OPTION_1_ID', 'OPTION_2_ID', 'OPTION_3_ID', 'OPTION_4_ID', 'STATUS', 'CHOSEN_OPTION'))

    def append_row(df, row):
        return pd.concat([
                        df, 
                        pd.DataFrame([row], columns=row.index)
                        ]).reset_index(drop=True)
    
    #o_section_name = driver.find_element(By.XPATH, '/html/body/div/div[3]/div[1]/div[1]/span[2]').text 

    div_questions = soup.find_all("div", {"class": "grp-cntnr"})

    o_section_name = ""
    question_pnl = driver_qa.find_elements(By.XPATH, "//div[@class='question-pnl']")
    for qstns in question_pnl:

        data_arr = qstns.text.split('\n')

        if (sno >= 1 and sno <= 20):
            o_section_name = driver_qa.find_element(By.XPATH, '/html/body/div/div[3]/div[1]/div[1]/span[2]').text 
        elif (sno >= 21 and sno <= 25):
            o_section_name = driver_qa.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/div[1]/span[2]").text  
        elif (sno >= 26 and sno <= 45):
            o_section_name = driver_qa.find_element(By.XPATH, "/html/body/div/div[3]/div[3]/div[1]/span[2]").text      
        elif (sno >= 46 and sno <= 50):
            o_section_name = driver_qa.find_element(By.XPATH, "/html/body/div/div[3]/div[4]/div[1]/span[2]").text  
        elif (sno >= 51 and sno <= 70):
            o_section_name = driver_qa.find_element(By.XPATH, "/html/body/div/div[3]/div[5]/div[1]/span[2]").text    
        elif (sno >= 71 and sno <= 75):
            o_section_name = driver_qa.find_element(By.XPATH, "/html/body/div/div[3]/div[6]/div[1]/span[2]").text             

        if (sno >= 1 and sno <= 20)  or (sno >= 26 and sno <= 45) or (sno >= 51 and sno <= 70):
            
            o_qa_no = sno
            o_QUESTION_TYPE = data_arr[5].split(':')[1].strip()
            o_QUESTION_ID = data_arr[6].split(':')[1].strip()
            o_Option1_ID = data_arr[7].split(':')[1].strip()
            o_Option2_ID = data_arr[8].split(':')[1].strip()
            o_Option3_ID = data_arr[9].split(':')[1].strip()
            o_Option4_ID = data_arr[10].split(':')[1].strip()
            o_Status = data_arr[11].split(':')[1].strip()
            o_Chosen_Option = data_arr[12].split(':')[1].strip()

            new_row = pd.Series({'SECTION': o_section_name,'QA_ID': o_qa_no, 'APPNO': o_ApplicationNo,'QUESTION_TYPE': o_QUESTION_TYPE,'QUESTION_ID': o_QUESTION_ID,'OPTION_1_ID': o_Option1_ID,'OPTION_2_ID': o_Option2_ID,
                        'OPTION_3_ID': o_Option3_ID,'OPTION_4_ID': o_Option4_ID,'STATUS': o_Status,'CHOSEN_OPTION': o_Chosen_Option})
            
            df = append_row(df, new_row)
        elif (sno >= 21 and sno <=25) or (sno >= 46 and sno <=50) or (sno >= 71 and sno <= 75):
            
            o_qa_no = sno
            o_QUESTION_TYPE = data_arr[2].split(':')[1].strip()
            o_QUESTION_ID = data_arr[3].split(':')[1].strip()
            o_Option1_ID = ''
            o_Option2_ID = ''
            o_Option3_ID = ''
            o_Option4_ID = ''
            o_Status = data_arr[4].split(':')[1].strip()
            o_Chosen_Option = data_arr[1].split(':')[1].strip()

            new_row = pd.Series({'SECTION': o_section_name,'QA_ID': o_qa_no, 'APPNO': o_ApplicationNo,'QUESTION_TYPE': o_QUESTION_TYPE,'QUESTION_ID': o_QUESTION_ID,'OPTION_1_ID': o_Option1_ID,'OPTION_2_ID': o_Option2_ID,
                        'OPTION_3_ID': o_Option3_ID,'OPTION_4_ID': o_Option4_ID,'STATUS': o_Status,'CHOSEN_OPTION': o_Chosen_Option})
            
            df = append_row(df, new_row)

        sno +=1

    akclist = []
    akclist = df.values.tolist()

    #-----------START OF INSERTION OF STUDENT DETAILS---------------------------------------------------------------------
    insert_stu_dtls = \
                "INSERT INTO O_JEEMAINS_RESPONSSHEET_25 (APPLICATIONNO, ROLLNO, CANDIDATENAME, SUBJECT, PROCESS_STATUS, PROCESS_USER, TESTDATE, TESTTIME) \
                VALUES ('"+ str(o_ApplicationNo) +"', '"+ str(o_RollNo) +"','"+ o_CandidateName +"', \
                '"+ o_Subject +"', 'D', '"+ v_process_user +"', '"+ o_TestDate+"', '"+ o_TestTime+"')"
            
    cur.execute(insert_stu_dtls) # Execute an INSERT statement

    update_IpStatus = "UPDATE I_JEEMAINS_ADMITCARD_06FEB25 SET PROCESS_STATUS = 'D' WHERE APPNO = '"+ str(o_ApplicationNo) +"'"
    cur.execute(update_IpStatus) # Execute an UPDATE statement
    #-----------END OF INSERTION OF STUDENT DETAILS---------------------------------------------------------------------

    #-----------START OF INSERTION OF STUDENT QUESTION AND ANSWER DETAILS---------------------------------------------------------------------
    cur.setinputsizes(None, 20)
    insert_qa = """insert into O_JEEMAINS_RESPONSESHEET_QA1_25 (QA_ID, APPLICATIONNO, QUESTION_TYPE, QUESTION_ID, OPTION_1_ID, OPTION_2_ID, OPTION_3_ID, OPTION_4_ID, STATUS, CHOSEN_OPTION, SECTION)
            values (:1, :2, :3, :4, :5, :6, :7, :8, :9, :10, :11 )"""
    cur.executemany(insert_qa, akclist)
    conn.commit()
    #-----------END OF INSERTION OF STUDENT QUESTION AND ANSWER DETAILS---------------------------------------------------------------------    
    driver_qa.quit() 

for row in res:
    driver = webdriver.Chrome()
    driver.maximize_window()
    v_appno = row[1]     
    v_password = row[2]
            
    try:  
        driver.get("https://examinationservices.nic.in/JeeMain2025/root/CandidateLogin.aspx?enc=Ei4cajBkK1gZSfgr53ImFVj34FesvYg1WX45sPjGXBpvTjwcqEoJcZ5VnHgmpgmK")
        
        i_regno = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtRegno")      
        i_regno.send_keys(v_appno)

        i_password = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtPassword")
        i_password.send_keys(v_password)

        Captchaimg = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_captchaimage")
        driver.execute_script("arguments[0].scrollIntoView(true);", Captchaimg)
        Captchaimg.screenshot('Screenshotcaptcha.png')

        captcha_text = extract_text_with_paddleocr('Screenshotcaptcha.png')
        i_Captcha = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtsecpin")      
        i_Captcha.send_keys(captcha_text)
        
        #Button press
        Submit_button = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_btnsignin')
        driver.execute_script("arguments[0].click();", Submit_button)
            
        try:
            v_Captcha = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_lblerror1').text 
        except:
            v_Captcha = ""     
                          
        if v_Captcha in ["CAPTCHA did not match. Please Re-enter."]:
            while attempts < max_attempts and not login_successful:
                i_regno = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtRegno") 
                i_regno.clear()     
                i_regno.send_keys(v_appno)

                i_password = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtPassword")
                i_password.clear()
                i_password.send_keys(v_password)

                Captchaimg = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_captchaimage")
                driver.execute_script("arguments[0].scrollIntoView(true);", Captchaimg)
                Captchaimg.screenshot('Screenshotcaptcha.png')

                captcha_text = extract_text_with_paddleocr('Screenshotcaptcha.png')
                i_Captcha = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtsecpin")      
                i_Captcha.send_keys(captcha_text)
                
                #Button press
                Submit_button = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_btnsignin')
                driver.execute_script("arguments[0].click();", Submit_button)

                try:
                    QA_details()
                    driver.quit()
                    login_successful = True
                    
                except:
                    login_successful=False    

        elif v_Captcha == "Invalid Application No or Password.":        
            update_IpStatus = "UPDATE I_JEEMAINS_ADMITCARD_06FEB25 SET PROCESS_STATUS = 'INVALID', ERROR_MESSAGE = '"+ str(v_Captcha) +"' WHERE APPNO = '"+ str(v_appno) +"'"
            cur.execute(update_IpStatus) # Execute an UPDATE statement
            conn.commit()    
            driver.quit()
        else:
            QA_details()  
            driver.quit()

    except:
        #pass
        update_IpStatus = "UPDATE I_JEEMAINS_ADMITCARD_06FEB25 SET PROCESS_STATUS = 'D', ERROR_MESSAGE = 'NO BTECH BUTTON' WHERE APPNO = '"+ str(v_appno) +"'"
        cur.execute(update_IpStatus) # Execute an UPDATE statement
        conn.commit()
        driver.quit()

                                    


