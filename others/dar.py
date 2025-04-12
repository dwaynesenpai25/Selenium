# from selenium.webdriver.chrome.service import Service as ChromeService
# from selenium.webdriver.support.ui import Select
# from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.common.keys import Keys
# import math
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from datetime import datetime
from selenium import webdriver
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

        self.file_path = "DAR.xlsx"
        self.df = pd.read_excel(self.file_path,)


        self.df["STATUS"] = self.df["STATUS"].astype(str)
        self.df["Debt Collection Agency"] = self.df["Debt Collection Agency"].astype(str)
        self.df["Email"] = self.df["Email"].astype(str)
        self.df["CTL4 / Unit"] = self.df["CTL4 / Unit"].astype(str)
        self.df["Account Number/ LAN/PAN"] = self.df["Account Number/ LAN/PAN"].astype(str)
        self.df["Loan Account Name"] = self.df["Loan Account Name"].astype(str)
        self.df["Transaction"] = self.df["Transaction"].astype(str)
        self.df["Report Submission"] = self.df["Report Submission"].astype(str)

        self.df["Account Type"] = self.df["Account Type"].astype(str)
        self.df["Activity"] = self.df["Activity"].astype(str)

        self.df["Client Contact Number"] = self.df["Client Contact Number"].astype(str)
        self.df["Client Email Address"] = self.df["Client Email Address"].astype(str)
        self.df["Contact Status"] = self.df["Contact Status"].astype(str)

        self.df["With New Contact Information?"] = self.df["With New Contact Information?"].astype(str)
        self.df["New Contact Type"] = self.df["New Contact Type"].astype(str)
        self.df["New Email Address"] = self.df["New Email Address"].astype(str)
        self.df["New Contact Number"] = self.df["New Contact Number"].astype(str)

        self.df["Client Type"] = self.df["Client Type"].astype(str)
        self.df["Negotiation Remarks"] = self.df["Negotiation Remarks"].astype(str)
        self.df["Negotiation Status"] = self.df["Negotiation Status"].astype(str)
        self.df["PTP / PTVS Date (DAR)"] = self.df["PTP / PTVS Date (DAR)"].astype(str)
        self.df["Reason for Default (DAR)"] = self.df["Reason for Default (DAR)"].astype(str)


        # NON-NEEDED DATA
        self.df["Classification of Account (DAR)"] = self.df["Classification of Account (DAR)"].astype(str)
        self.df["Date Worked (DAR)"] = self.df["Date Worked (DAR)"].astype(str)
        self.df["Action (DAR)"] = self.df["Action (DAR)"].astype(str)
        self.df["Email Address (DAR) if CONTACTED VIA EMAIL"] = self.df["Email Address (DAR) if CONTACTED VIA EMAIL"].astype(str)
        self.df["Contact Number (DAR) if CONTACTED VIA CALL"] = self.df["Contact Number (DAR) if CONTACTED VIA CALL"].astype(str)
        self.df["REMARKS"] = self.df["REMARKS"].astype(str)

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

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    def dar(self, acc_type, act, c_num, c_email, c_status, with_contact, contact_type, new_email, new_num, client_type, nego_remarks, coa, dw, action, email_add, contact_num, nego_stat, ptp, reason_dar, remarks):  
        wait = WebDriverWait(self.driver, 5)  # Set an explicit wait time

        question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')

        for question_item in question_items:
            question_title = question_item.find_element(By.CSS_SELECTOR, 'div div span[data-automation-id="questionTitle"]').text.split('\n')[1]
            print(question_title)

            match question_title:
                case 'Account Type':
                    question_item.find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                    normalized_text = self.normalize_option_text(acc_type) 
                    self.click_option(normalized_text)

                    # Wait for new elements (Activity dropdown) to load dynamically
                    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                    question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')

                    # Select Activity
                    question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                    self.click_option(act)

                    # Wait for Client Contact Number or Email fields to appear
                    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                    question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')

                    # Identify Input Field
                    input_field = question_items[-2].find_element(By.CSS_SELECTOR, 'div div span input')
                    
                    input_field.clear()

                    # Determine whether to enter Contact Number or Email
                    input_value = c_num if act == "Call" else c_email
                    input_field.send_keys(input_value)


                    question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                    self.click_option(c_status)

                    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                    question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                    # question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()

                    if c_status == "Contacted":
                        self.click_option(client_type)

                        if client_type == "Insurance Company / HPG / LTO / Other Governmen":
                            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                            question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]') 

                            input_field = question_items[-2].find_element(By.CSS_SELECTOR, 'div div span textarea')
                            input_field.clear()
                            input_field.send_keys(nego_remarks)

                            self.click_element_by_CSS(f'span[data-automation-value="{with_contact}"]')

                            if with_contact == "Yes":
                                self.click_element_by_CSS(f'span[data-automation-value="{contact_type}"]')

                                question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                                
                                input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span input')
                                
                                input_field.clear()
                                input_value = new_email if contact_type == "Email Address" else new_num
                                input_field.send_keys(input_value)
                        else:
                            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                            question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]') 
                            question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                            normalized_text = self.normalize_option_text(reason_dar) 
                            self.click_option(normalized_text)

                            question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]') 
                            question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                            normalized_text = self.normalize_option_text(nego_stat) 
                            self.click_option(normalized_text)
                            
                            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                            question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]') 

                            if nego_stat in ["Promised-to-Pay", "Promised-to-Voluntary Surrender"]:
                                input = question_items[-3].find_element(By.CSS_SELECTOR, 'div div div div div div input')
                                input.clear()
                                input.send_keys(ptp)

                                input_field = question_items[-2].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(nego_remarks)

                                self.click_element_by_CSS(f'span[data-automation-value="{with_contact}"]')

                                if with_contact == "Yes":
                                    self.click_element_by_CSS(f'span[data-automation-value="{contact_type}"]')

                                    question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                                    
                                    input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span input')
                                    
                                    input_field.clear()
                                    input_value = new_email if contact_type == "Email Address" else new_num
                                    input_field.send_keys(input_value)
                                
                            else:
                                input_field = question_items[-2].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(nego_remarks)

                                self.click_element_by_CSS(f'span[data-automation-value="{with_contact}"]')

                                if with_contact == "Yes":
                                    self.click_element_by_CSS(f'span[data-automation-value="{contact_type}"]')

                                    question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                                    
                                    input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span input')
                                    
                                    input_field.clear()
                                    input_value = new_email if contact_type == "Email Address" else new_num
                                    input_field.send_keys(input_value)
                 
                    elif c_status == "Uncontacted - No answer / No reply":
                        self.click_element_by_CSS(f'span[data-automation-value="{with_contact}"]')

                        if with_contact == "Yes":
                            self.click_element_by_CSS(f'span[data-automation-value="{contact_type}"]')

                            question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                            
                            input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span input')
                            
                            input_field.clear()
                            input_value = new_email if contact_type == "Email Address" else new_num
                            input_field.send_keys(input_value)
                        else:
                            self.click_element_by_CSS(f'span[data-automation-value="{with_contact}"]')


                    self.click_element_by_CSS('button[data-automation-id="submitButton"]')

                # case 'Classification of Account (DAR)':
                #     question_item.find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                #     self.click_option(coa)

                # case 'Date Worked (DAR)':
                #     input = question_item.find_element(By.CSS_SELECTOR, 'div div div div div div input')
                #     input.clear()
                #     input.send_keys(dw)

                # case 'Action (DAR)':
                #     question_item.find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                #     self.click_option(action)

                #     question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')

                #     input = question_items[-2].find_element(By.CSS_SELECTOR, 'div div span input')
                #     input.clear()

                #     if action == "CONTACTED VIA EMAIL":
                #         input.send_keys(email_add)
                #     else:
                #         # contact_num if len(contact_num) >= 11 else f"0{contact_num}"
                #         input.send_keys(contact_num)

                #     question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                #     self.click_option(nego_stat)

                #     question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')

                #     if nego_stat in ["Promise-to-Pay", "Promise-to-Voluntary Surrender", "Uncontacted"] and ptp is not None:
                #         input = question_items[-3].find_element(By.CSS_SELECTOR, 'div div div div div div input')
                #         input.clear()
                #         input.send_keys(ptp)
                #         # pass

                #     question_items[-2].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                #     self.click_option(reason_dar)

                #     input = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span textarea')
                #     input.clear()
                #     input.send_keys(remarks)

                    
                   
    def main(self):
        self.click_element_by_XPATH('//*[@id="form-main-content1"]/div/div[3]/button/div')

        for index, row in self.df.iterrows():

            status = str(row["STATUS"]).strip()

            if status == "DONE":
                continue

            agency = str(row["Debt Collection Agency"]).strip()
            email = str(row["Email"]).strip()
            unit = str(row["CTL4 / Unit"]).strip()
            acc_num = str(row["Account Number/ LAN/PAN"]).strip()
            loan_acc_num = str(row["Loan Account Name"]).strip()
            transac = str(row["Transaction"]).strip()
            report_sub = str(row["Report Submission"]).strip()
            acc_type = str(row["Account Type"]).strip()
            act = str(row["Activity"]).strip()
            c_num = str(row["Client Contact Number"]).strip()
            c_email = str(row["Client Email Address"]).strip()
            c_status = str(row["Contact Status"]).strip()
            with_contact = str(row["With New Contact Information?"]).strip()
            contact_type = str(row["New Contact Type"]).strip()
            new_email = str(row["New Email Address"]).strip()
            new_num = str(row["New Contact Number"]).strip()

            client_type = str(row["Client Type"]).strip()
            nego_remarks = str(row["Negotiation Remarks"]).strip()


            coa = str(row["Classification of Account (DAR)"]).strip()

            date_worked_str = str(row["Date Worked (DAR)"]).strip()

            try:
                date_worked = datetime.strptime(date_worked_str, "%Y-%m-%d %H:%M:%S")
            except:
                date_worked = datetime.strptime(date_worked_str, "%Y-%m-%d")

            dw = date_worked.strftime("%m/%d/%Y")

            action = str(row["Action (DAR)"]).strip()
            email_add = str(row["Email Address (DAR) if CONTACTED VIA EMAIL"]).strip()
            contact_num = str(row["Contact Number (DAR) if CONTACTED VIA CALL"]).strip()
            nego_stat = str(row["Negotiation Status"]).strip()

            ptp_dar_str = str(row["PTP / PTVS Date (DAR)"]).strip()
            ptp = None

            try:
                ptp_dar = datetime.strptime(ptp_dar_str, "%Y-%m-%d")
                ptp = ptp_dar.strftime("%m/%d/%Y")
            except:
                pass
 
            reason_dar = str(row["Reason for Default (DAR)"]).strip()
            remarks = str(row["REMARKS"]).strip()

            self.omtool()
            sleep(1)

            self.external()
            sleep(1)

            self.iat(agency, email, unit, acc_num, loan_acc_num, transac, report_sub)
            sleep(1)

            self.dar(acc_type, act, c_num, c_email, c_status, with_contact, contact_type, new_email, new_num, client_type, nego_remarks, coa, dw, action, email_add, contact_num, nego_stat, ptp, reason_dar, remarks)            
            sleep(1)

            self.df.loc[index, ["STATUS"]] = "DONE"
            self.df.to_excel(f"{self.file_path}", index=False)
            print(f"{index + 1} / {self.total}")

            sleep(1)
            #WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, f'//*[contains(text(), "Submit another response")]'))).click()
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'span[data-automation-id="submitAnother"]'))).click()

            sleep(3)



if __name__ == '__main__':
    Start().main()
