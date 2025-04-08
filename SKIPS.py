import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import math
from time import sleep
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.options import Options
import re

class Start():
    def __init__(self):
        options = Options()
        options.add_experimental_option("excludeSwitches", ["enable-logging"])

        print("testing started")
        driver = webdriver.Chrome(options=options)
        
        self.driver = driver

        self.main_url = "https://forms.office.com/Pages/ResponsePage.aspx?id=YITPJFFnikKrGajrBN_kR14iqPry4HNOqTNpc3I2zHBUNE82NFg0SUtXMzBEOEZJRks3SzBDWlVCMy4u"
        

        self.driver.get(self.main_url)

        self.driver.implicitly_wait(2)

        self.file_path ="SKIPS.xlsx"
        self.df = pd.read_excel(self.file_path,)
        self.df = self.df.astype(str)

        self.total = len(self.df["STATUS"])


    def click_element_by_CSS(self, path):
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, path))).click()

    def click_element_by_XPATH(self, path):
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.XPATH, path))).click()

    def omtool(self):
        self.click_element_by_CSS('span[data-automation-value="External - Debt Collection Agencies"]')

        self.click_element_by_CSS('button[data-automation-id="nextButton"]')

    def external(self): # External - Debt Collection Agencies
        self.click_element_by_CSS('span[data-automation-value="I agree"]')

        self.click_element_by_CSS('button[data-automation-id="nextButton"]')

    def normalize_option_text(self,text):
        return re.sub(r"\s*-\s*", "-", text.strip())
    
    def click_option(self, text):
        print(text)
        self.click_element_by_CSS(f'div[role="listbox"] div span span[aria-label="{text}"]')

    def scan_and_select_option(self, text):
        wait = WebDriverWait(self.driver, 5)
        keywords = ['Residence', 'Business', 'Lead']
        normalized_text = None
        for keyword in keywords:
            if keyword in text:
                normalized_text = keyword
                break  # Stop once we find the first matching keyword
        if not normalized_text:
            print("No matching keyword found in the provided text.")
            return

        print(f"Normalized text: {normalized_text}")
        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[role="listbox"]')))
        options = self.driver.find_elements(By.CSS_SELECTOR, 'div[role="listbox"] div span span')

        for option in options:
            option_text = option.get_attribute("aria-label")
            # Check if the normalized text matches the option text
            if normalized_text in option_text:
                print(f"Found matching option: {option_text}")
                option.click()  # Click on the matching option
                return  # Exit after selecting the first match
        print(f"No matching option found for {normalized_text}.")

    def iat(self, agency, email, unit, acc_num, loan_acc_num, transac, report_sub): # Information and Transaction
        question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')

        for question_item in question_items:
            question_title = question_item.find_element(By.CSS_SELECTOR, 'div div span[data-automation-id="questionTitle"]').text.split('\n')[1]

            match question_title:
                case 'Debt Collection Agency':
                    question_item.find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                    self.click_option(agency)

                case 'Email':
                    input = question_item.find_element(By.CSS_SELECTOR, 'div div span input')
                    input.clear()
                    input.send_keys(email)

                case 'CTL4 / Unit':
                    question_item.find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                    self.click_option(unit)

                case 'Account Number - LAN/PAN':
                    input = question_item.find_element(By.CSS_SELECTOR, 'div div span input')
                    input.clear()
                    input.send_keys(acc_num)

                case 'Loan Account Name':
                    input = question_item.find_element(By.CSS_SELECTOR, 'div div span input')
                    input.clear()
                    input.send_keys(loan_acc_num)

                case 'Transaction':
                    self.click_element_by_CSS(f'span[data-automation-value="{transac}"]')

                    question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')

                    question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()

                    self.click_option(report_sub)
                    

        self.click_element_by_CSS('button[data-automation-id="nextButton"]')

    def skip_trace(self, endo_date, tracing_attempt, add_visit, fv_client, type_client, ptp_date, fv_unit, w_brgy,brgy_remark,fv_remarks,other_add,add_type):
        wait = WebDriverWait(self.driver, 5)
        question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')

        for question_item in question_items:
            question_title = question_item.find_element(By.CSS_SELECTOR, 'div div span[data-automation-id="questionTitle"]').text.split('\n')[1]
            
            match question_title:
                case 'Endorsement Date (FV)':
                    input = question_item.find_element(By.CSS_SELECTOR, 'div div div div div div input')
                    input.clear()
                    input.send_keys(endo_date)

                case 'Field Visit Attempt':
                    question_item.find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                    normalized_text = self.normalize_option_text(tracing_attempt) 
                    self.click_option(normalized_text)
                   
                    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                    question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                    question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                    self.scan_and_select_option(add_visit)

                    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                    question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                    
                    if add_visit != "Lead Address (not same address provided in endorsement list)":
                       
                        question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                        self.click_option(fv_client)
                        
                        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                        question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                        question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                        normalized_text = self.normalize_option_text(type_client) 
                        self.click_option(normalized_text)
                  
                        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                        question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                    
                        if type_client not in ["Promised-to-Pay", "Promised-to-Voluntary Surrender"]:
                            question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                            self.click_option(fv_unit)
                            
    
                            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                            question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                        
                            question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                            self.click_option(w_brgy)
             
                            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                            question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                            
                        
                            if w_brgy == 'Yes':
                                input_field = question_items[-2].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(brgy_remark)
                 
                                input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(fv_remarks)
                            else:
                                input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(fv_remarks)
                        else:
                            print(ptp_date)
                            input = question_items[-2].find_element(By.CSS_SELECTOR, 'div div div div div div input')
                            input.clear()
                            input.send_keys(ptp_date)
                            
                       
                            question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                            self.click_option(fv_unit)
                            
                            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                            question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                        
                            question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                            self.click_option(w_brgy)
                            
                            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                            question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                            
                        
                            if w_brgy == 'Yes':
                                input_field = question_items[-2].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(brgy_remark)
                                
                                input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(fv_remarks)
                            else:
                                input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(fv_remarks)
                    else:
                        input = question_items[-2].find_element(By.CSS_SELECTOR, 'div div div div div div input')
                        input.clear()
                        input.send_keys(other_add)
                      
                        question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                        self.click_option(add_type)
                        
                        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                        question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                        
                        question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                        self.click_option(fv_client)
                        
                        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                        question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                        question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                        self.click_option(type_client)
                        
                        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                        question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                    
                        if type_client not in ["Promised-to-Pay", "Promised-to-Voluntary Surrender"]:
                            question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                            self.click_option(fv_unit)
                            
                            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                            question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                        
                            question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                            self.click_option(w_brgy)
                            
                            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                            question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                            
                        
                            if w_brgy == 'Yes':
                                input_field = question_items[-2].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(brgy_remark)
                                
                                input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(fv_remarks)
                            else:
                                input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(fv_remarks)
                        else:
                            input = question_items[-2].find_element(By.CSS_SELECTOR, 'div div div div div div input')
                            input.clear()
                            input.send_keys(ptp_date)
                            
                            question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                            self.click_option(fv_unit)
                            
                            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                            question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                        
                            question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                            self.click_option(w_brgy)
                            
                            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                            question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                            
                        
                            if w_brgy == 'Yes':
                                input_field = question_items[-2].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(brgy_remark)
                                
                                input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(fv_remarks)
                            else:
                                input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(fv_remarks)
                    # self.click_element_by_CSS('button[data-automation-id="submitButton"]')

    def main(self):
        self.click_element_by_XPATH('//*[@id="form-main-content1"]/div/div[3]/button')

        for index, row in self.df.iterrows():

            status = str(row["STATUS"]).strip()

            if status =="DONE":
                continue
            
            ptp_date = None
            endo_date = None

            try:
                ptp_dar = datetime.strptime(row["PTP / PTVS Date"], "%Y-%m-%d")
                ptp_date = ptp_dar.strftime("%m/%d/%Y")
                
                endo_dar = datetime.strptime(row["Endorsement Date (FV)"], "%Y-%m-%d")
                endo_date = endo_dar.strftime("%m/%d/%Y")
            except:
                pass
            
            self.omtool()
            sleep(1)
            
            self.external()
            sleep(1)

            self.iat(
                row["Debt Collection Agency"], row["Email"], 
                str(int(float(row["CTL4 / Unit"]))) if float(row["CTL4 / Unit"]).is_integer() else str(row["CTL4 / Unit"]),  # Fixes TypeError
                row["Account Number - LAN/PAN"], row["Loan Account Name"],
                row["Transaction"], row["Report Submission"]
            )
            sleep(1)

       
            self.skip_trace(endo_date,row["Field Visit Attempt"],row["Address Visited"],row["Field Visit Result - Client"],row["Client Type"],ptp_date,row["Field Visit Result - Unit"],row["W/ Barangay Confirmation?"],row["Barangay Confirmation Remarks"],row["Field Visit Remarks"],row["Other Address"],row["Type of Address"])
            sleep(1)

            self.df.loc[index, ["STATUS"]] ="DONE"
            self.df.to_excel(f"{self.file_path}", index=False)
            print(f"{index + 1} / {self.total}")

            sleep(1)
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, f'//*[contains(text(),"Submit another response")]'))).click()

            sleep(3)
          
if __name__ == '__main__':
    Start().main()