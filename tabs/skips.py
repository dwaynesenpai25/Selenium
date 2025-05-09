from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from datetime import datetime
import re
from selenium.webdriver.chrome.options import Options
import streamlit as st
import html
class Skips():
    
    def __init__(self):
        options = Options()
        options.add_experimental_option("excludeSwitches", ["enable-logging"])

        print("Testing started")
        try:
            self.driver = webdriver.Chrome(options=options)
        except Exception as e:
            st.error(f"Failed to initialize WebDriver: {e}")
            st.stop()

        self.main_url = "https://forms.office.com/pages/responsepage.aspx?id=YITPJFFnikKrGajrBN_kRwfatORDXSVLm247ZiBPtmJUNVRVNFVSVEdMT0o1TDkwRUw1T0lLV0JLRi4u"
        self.driver.get(self.main_url)
        try:
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            print("Page loaded successfully")
        except Exception as e:
            st.error(f"Failed to load page: {e}")
            self.driver.quit()
            st.stop()

        self.driver.implicitly_wait(2)
        self.expander = st.expander("Process Details", expanded=False)  # Create expander here
        self.status_text = self.expander.text("Initializing...")
        self.progress_placeholder = self.expander.empty()

    def click_element_by_CSS(self, path):
        try:
            WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, path))).click()
        except Exception as e:
            self.expander.error(f"Failed to click element with CSS {path}: {e}")

    def click_element_by_XPATH(self, path):
        try:
            WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable((By.XPATH, path))).click()
        except Exception as e:
            self.expander.error(f"Failed to click element with XPATH {path}: {e}")

    def escape_xpath_text(self,text):
        """Escape apostrophes in a string for use in an XPath expression."""
        if "'" not in text:
            return f"'{text}'"
        parts = text.split("'")
        return "concat(" + ", \"'\", ".join(f"'{part}'" for part in parts) + ")"
    
    def normalize_option_text(self, text):
        # Replace curly quotes and long dashes
        text = text.replace("’", "'")
        text = re.sub(r"[-\u2013\u2014]", "-", text.strip())
        return " ".join(text.split())
    
    # def click_option(self, text):
    #     normalized_text = self.normalize_option_text(text)
    #     print(f"Attempting to click option: {normalized_text}")

    #     escaped_text = self.escape_xpath_text(normalized_text)
    #     # xpath = f"//div[@role='listbox']//span[contains(normalize-space(text()), {escaped_text})]"
    #     xpath = f"//div[@role='listbox']//span[contains(@aria-label, {escaped_text})]"
    #     print(f"XPath used: {xpath}")

    #     try:
    #         option = WebDriverWait(self.driver, 10).until(
    #             EC.element_to_be_clickable((By.XPATH, xpath))
    #         )
    #         option.click()
    #         print(f"Successfully selected option: {normalized_text}")
    #     except Exception as e:
    #         self.expander.error(f"Failed to select option {normalized_text}: {str(e)}")
    #         raise

    def click_option(self, text):
        normalized_text = self.normalize_option_text(text)
        print(f"Attempting to click option: {normalized_text}")

        escaped_text = self.escape_xpath_text(normalized_text)

        # 👉 Insert this small fix here:
        # Replace normal spaces with a pattern that matches any space (regular OR &nbsp;)
        xpath = f"""//div[@role='listbox']//span[contains(translate(normalize-space(text()), '\u00A0', ' '), {escaped_text})]"""

        print(f"XPath used: {xpath}")

        try:
            option = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            option.click()
            print(f"Successfully selected option: {normalized_text}")
        except Exception as e:
            self.expander.error(f"Failed to select option {normalized_text}: {str(e)}")
            raise


    def external(self, agency, email): 
        print("testing started")
        question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')

        for question_item in question_items:
            question_title = question_item.find_element(By.CSS_SELECTOR, 'div div span[data-automation-id="questionTitle"]').text.split('\n')[1]
            print(question_title)

            match question_title:
                case 'Debt Collection Agency':
                    question_item.find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                    self.click_option(agency)

                case 'DCA Email Address':
                    input = question_item.find_element(By.CSS_SELECTOR, 'div div span input')
                    input.clear()
                    input.send_keys(email)
                
                    self.click_element_by_CSS('span[data-automation-value^="I agree."]')

        self.click_element_by_CSS('button[data-automation-id="nextButton"]')

    def iat(self, acc_type, loan_acc_num, unit, acc_num, endo_date, transac, report_sub): 
        question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')

        for question_item in question_items:
            question_title = question_item.find_element(By.CSS_SELECTOR, 'div div span[data-automation-id="questionTitle"]').text.split('\n')[1]

            match question_title:
                case 'Account Type':
                    question_item.find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                    normalized_text = self.normalize_option_text(acc_type) 
                    self.click_option(normalized_text)

                case 'Account Name':
                    input = question_item.find_element(By.CSS_SELECTOR, 'div div span input')
                    input.clear()
                    input.send_keys(loan_acc_num)

                case 'CTL4':
                    question_item.find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                    self.click_option(unit)

                case 'Loan Account Number':
                    input = question_item.find_element(By.CSS_SELECTOR, 'div div span input')
                    input.clear()
                    input.send_keys(acc_num)

                case 'Endorsement Date':
                    input = question_item.find_element(By.CSS_SELECTOR, 'div div div div div div input')
                    input.clear()
                    input.send_keys(endo_date)
      
                case 'Transaction':
                    question_item.find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                    self.click_option(transac)

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
                case 'Field Visit Date':
                    input = question_item.find_element(By.CSS_SELECTOR, 'div div div div div div input')
                    input.clear()
                    input.send_keys(endo_date)

                case 'Field Visit Attempt':
                    question_item.find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                    normalized_text = self.normalize_option_text(tracing_attempt) 
                    self.click_option(normalized_text)

                case 'Address Visited':
                    question_item.find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                    normalized_text = self.normalize_option_text(add_visit) 
                    self.click_option(normalized_text)

                    if add_visit != "Lead Address (not same address provided in endorsement list)":
                        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                        question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                       
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
                            
                            if fv_unit != "Positive Unit (Seen in the address visited)":
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
                                wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                                question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
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

                            if fv_unit != "Positive Unit (Seen in the address visited)":
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
                                wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                                question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                                input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(fv_remarks)

                    else:
                        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                        question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                            
                        input = question_items[-2].find_element(By.CSS_SELECTOR, 'div div div div div div input')
                        input.clear()
                        input.send_keys(other_add)
                        
                        normalized_text = self.normalize_option_text(add_type)
                        self.click_element_by_CSS(f'span[data-automation-value^="{normalized_text}"]')

                        
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
                            
                            if fv_unit != "Positive Unit (Seen in the address visited)":
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
                                wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                                question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')
                                input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(fv_remarks)

                        else:
                            input = question_items[-2].find_element(By.CSS_SELECTOR, 'div div div div div div input')
                            input.clear()
                            input.send_keys(ptp_date)
                            
                            question_items[-1].find_element(By.CSS_SELECTOR, 'div div div div[aria-haspopup="listbox"]').click()
                            normalized_text = self.normalize_option_text(fv_unit)
                            self.click_option(normalized_text)
                            
                            if fv_unit not in  ['Positive Unit (Seen in the address visited)']:
                                print("fv_unit")
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
                                wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')))
                                question_items = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-automation-id="questionItem"]')

                                input_field = question_items[-1].find_element(By.CSS_SELECTOR, 'div div span textarea')
                                input_field.clear()
                                input_field.send_keys(fv_remarks)  

        self.click_element_by_CSS('button[data-automation-id="submitButton"]')
                    
    def main(self, df):
        print("Starting main process")
        try:
            self.click_element_by_XPATH('//*[@id="form-main-content1"]/div/div[3]/button/div')
            print("Initial button clicked")

            total = len(df)
            self.status_text.text("Starting process...")

            for index, row in df.iterrows():
                status = str(row["STATUS"]).strip()

                if status == "DONE":
                    self.status_text.text(f"Skipping row {index} - already done")
                    continue

                self.status_text.text(f"Processing row {index + 1} / {total}...")
                self.progress_placeholder.write(f"Current row: {index + 1} - {row['Debt Collection Agency']}")

                try:
                    self.external(row["Debt Collection Agency"], row["Email"])
                    sleep(2)

                    self.iat(
                        row["Account Type"], row["Loan Account Name"], row["CTL4 / Unit"],
                        row["Account Number/ LAN/PAN"], row["Endo Date"],
                        row["Transaction"], row["Report Submission"]
                    )

                    self.skip_trace(row["Field Visit Date"],row["Field Visit Attempt"],row["Address Visited"],row["Field Visit Result - Client"],row["Client Type"],row["PTP / PTVS Date"],row["Field Visit Result - Unit"],row["W/ Barangay Confirmation?"],row["Barangay Confirmation Remarks"],row["Field Visit Remarks"],row["Other Address"],row["Type of Address"])
                    sleep(1)

                    # self.dar(
                    #     row["Activity"], row["Client Contact Number"],
                    #     row["With New Contact Information?"], row["New Contact Type"],
                    #     row["New Email Address"], row["New Contact Number"],
                    #     row["Client Type"], row["Negotiation Remarks"],
                    #     row["Negotiation Status"], row["PTP / PTVS Date (DAR)"], 
                    #     row["PTP Amount"], row["Reason for Default (DAR)"]
                    # )

                    sleep(2)
                    submit_another = WebDriverWait(self.driver, 15).until(
                        EC.element_to_be_clickable((By.XPATH, f'//*[contains(text(), "Submit another response")]'))
                    )
                    submit_another.click()
                    sleep(3)

                    df.loc[index, "STATUS"] = "DONE"
                    df.to_excel("output.xlsx", index=False)
                    self.status_text.text(f"Processed {index + 1} / {total}")

                except Exception as e:
                    self.expander.error(f"Error at row {index}: {str(e)}")
                    self.driver.save_screenshot(f"error_row_{index}.png")
                    break

        except Exception as e:
            self.expander.error(f"Main process failed: {str(e)}")
            self.driver.save_screenshot("main_error.png")
        finally:
            self.driver.quit()
            self.status_text.text("Driver closed and process completed")
