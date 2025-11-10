import os
import pickle
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException

import openpyxl


class SalesforceAutomation:
    def __init__(self, cookie_file="salesforce_cookies.pkl"):
        self.url = "https://hr2day-2918.lightning.force.com/lightning/n/hr2d__Interaction_Center_Lightning"
        self.cookie_file = cookie_file
        self.driver = None
        self.selectors = {
            "new_declaratie": "/html/body/div[4]/div[1]/section/div[2]/div[2]/div[2]/div[1]/div/div/div/div/hr2d-lwc-e-i-c/div/div[1]/div[7]/div/div[2]/div[2]/c-lwc-employee-declarations/article/c-lwc-panel-header/div/header/div[3]/slot/button",
            "declaratie_iframe": "/html/body/div[4]/div[1]/section/div[2]/div[2]/div[2]/div[1]/div/div/div/div[2]/div/force-aloha-page/div/iframe",
            "declaratie": "/html/body/form/span/span/div/div/div[3]/div[1]/div/div[1]/span/span/div/div[2]/div/div/div/div/div/div[4]/div/span/div/ul/li",
            "uren": "/html/body/form/span/span/div/div/div[3]/div[1]/div/div[1]/span[2]/div/div[2]/div/div/div/div/div/div/div/div/div/div/a[1]",
            "uren_date": "/html/body/form/span/span/div/div/div[3]/div[1]/div/div[1]/span[2]/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div[1]/div[1]/div/span/input",
            "uren_type": "/html/body/form/span/span/div/div/div[3]/div[1]/div/div[1]/span[2]/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/select",
            "kilometers": '//*[@id="page:theForm:j_id8:j_id9:j_id131:j_id132:j_id135:j_id136:j_id139:j_id140:tabdetails:j_id166:PanelCards:j_id599:j_id612:j_id613:j_id620"]',
            "datefield": "/html/body/form/span/span/div/div/div[3]/div[1]/div/div[1]/span[2]/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div[1]/div[1]/div/span/input",
            "traveltype": "/html/body/form/span/span/div/div/div[3]/div[1]/div/div[1]/span[2]/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/select",
            "transporttype": "/html/body/form/span/span/div/div/div[3]/div[1]/div/div[1]/span[2]/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div[1]/div[4]/div/div/select",
            "from": "/html/body/form/span/span/div/div/div[3]/div[1]/div/div[1]/span[2]/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div[1]/div[6]/span/div[1]/div/div/div[2]/div/div/select",
            "to": "/html/body/form/span/span/div/div/div[3]/div[1]/div/div[1]/span[2]/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div[1]/div[7]/span/div[1]/div/div/div[2]/div/div/select",
            "retour": "/html/body/form/span/span/div/div/div[3]/div[1]/div/div[1]/span[2]/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div[1]/div[7]/span/div/div/div/div[1]/label",
            "bereken": "/html/body/form/span/span/div/div/div[3]/div[1]/div/div[1]/span[2]/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div[1]/div[9]/span/div/div/div/div[2]/div/div/a",
            "save": "#bijlagenPanel > div.slds-col.slds-size_1-of-1.slds-p-top_small > div > button.slds-button.slds-button_brand",
            "savenew": "#bijlagenPanel > div.slds-col.slds-size_1-of-1.slds-p-top_small > div > button:nth-child(2)",
        }

    def setup_driver(self):
        chrome_options = Options()
        # Remove automation flags to avoid detection
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option("useAutomationExtension", False)
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")

        service = Service()
        options = webdriver.ChromeOptions()
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.maximize_window()

    def save_cookies(self):
        cookies = self.driver.get_cookies()
        with open(self.cookie_file, "wb") as f:
            pickle.dump(cookies, f)
        print(f"Cookies saved to {self.cookie_file}")

    def load_cookies(self):
        if not os.path.exists(self.cookie_file):
            return False

        self.driver.execute_cdp_cmd("Network.enable", {})

        try:
            with open(self.cookie_file, "rb") as f:
                cookies = pickle.load(f)

            for cookie in cookies:
                if "domain" in cookie and cookie["domain"].startswith("."):
                    cookie["domain"] = cookie["domain"][1:]

                if "expiry" in cookie:
                    try:
                        if cookie["expiry"] < time.time():
                            del cookie["expiry"]
                    except:
                        del cookie["expiry"]

                try:
                    self.driver.execute_cdp_cmd("Network.setCookie", cookie)
                except Exception as e:
                    print(
                        f"Could not add cookie: {cookie.get('name', 'unknown')}, Error: {e}"
                    )

            print("Cookies loaded successfully into the browser.")

            self.driver.get(self.url)
            self.driver.execute_cdp_cmd("Network.disable", {})

            return True
        except Exception as e:
            print(f"Error loading cookies: {e}")
            return False

    def wait_for_manual_login(self, timeout=300):
        print("\n=== MANUAL LOGIN REQUIRED ===")
        print("Please complete the Microsoft login in the browser window.")
        print(f"You have {timeout} seconds to complete the login...")
        print("Script will continue automatically after successful login.\n")

        start_time = time.time()

        while time.time() - start_time < timeout:
            current_url = self.driver.current_url

            if (
                "force.com" in current_url
                and "login.microsoftonline.com" not in current_url
            ):
                print("Login detected! Continuing...")
                time.sleep(5)
                return True

            time.sleep(1)

        print("Login timeout reached!")
        return False

    def is_logged_in(self):
        """Check if we're successfully logged into Salesforce"""
        try:
            time.sleep(3)
            current_url = self.driver.current_url

            if "login.microsoftonline.com" in current_url:
                return False

            if (
                "force.com" in current_url
                and "interaction_center" in current_url.lower()
            ):
                return True

            return False
        except:
            return False

    def wait_for_xpath(self, xpath):
        result = None
        while result is None:
            try:
                result = self.driver.find_element(By.XPATH, xpath)
            except:
                result = None
        time.sleep(1)
        return result

    def wait_for_ec(self, xpath):
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )

    def load_xlsx_data(self):
        try:
            workbook = openpyxl.load_workbook("Data.xlsx")
            sheet = workbook.active
        except FileNotFoundError:
            print("Error: 'Data.xlsx' not found. Exiting...")
            exit()
        data = [list(row) for row in sheet.iter_rows(values_only=True)]
        data = data[1:]
        data = [sublist for sublist in data if sublist[0] == 0]
        return data

    def update_xlsx_data(self):
        pass

    def select_dropdown_value(self, selector_key, value, max_attempts=30):
        for attempt in range(max_attempts):
            try:
                element = self.wait_for_xpath(self.selectors[selector_key])
                time.sleep(1)
                select = Select(element)
                select.select_by_visible_text(value)
                return

            except StaleElementReferenceException:
                if attempt == max_attempts - 1:
                    raise
                time.sleep(0.5)

    def run_forms(self, data):
        travel = [item for item in data if item[6] == "None"]
        homework = [item for item in data if item[6] != "None"]
        button = self.wait_for_xpath(self.selectors["new_declaratie"])
        button.click()

        childIframe = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, self.selectors["declaratie_iframe"])
            )
        )
        self.driver.switch_to.frame(childIframe)

        button = self.wait_for_xpath(self.selectors["declaratie"])
        button.click()

        if len(travel) != 0:
            self.run_travel_form(travel)

        if len(homework) != 0:
            self.run_homework_form(homework)

    def run_travel_form(self, data):
        print("Open travel form")
        button = self.wait_for_element_with_text("a", "Kilometers")
        # button = self.wait_for_xpath(self.selectors["kilometers"])
        button.click()

        for i, date in enumerate(data):
            print("Processing travel data")
            self.wait_for_ec(self.selectors["datefield"])
            time.sleep(1)
            input = self.wait_for_xpath(self.selectors["datefield"])
            if isinstance(date[1], datetime):
                date_string = date[1].strftime("%d-%m-%Y")
            input.send_keys(date_string)

            self.select_dropdown_value("traveltype", date[2])
            self.select_dropdown_value("transporttype", date[3])
            self.select_dropdown_value("from", date[4])
            self.select_dropdown_value("to", date[5])

            self.wait_for_ec(self.selectors["bereken"])
            button = self.wait_for_xpath(self.selectors["bereken"])
            button.click()
            time.sleep(2)

            if i == len(data) - 1:
                print("Done, saving last one")
                time.sleep(2)
                self.driver.find_element(
                    By.CSS_SELECTOR, self.selectors["save"]
                ).click()
                time.sleep(2)
            else:
                print("Save and new")
                time.sleep(2)
                self.driver.find_element(
                    By.CSS_SELECTOR, self.selectors["savenew"]
                ).click()
                time.sleep(2)

            self.driver.switch_to.default_content()
            self.wait_for_xpath(self.selectors["declaratie_iframe"])
            childIframe = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, self.selectors["declaratie_iframe"])
                )
            )
            self.driver.switch_to.frame(childIframe)
        print("Done with travel form")

    def run_homework_form(self, data):
        print(data)
        print("Open homework form")
        button = self.wait_for_element_with_text("a", "Aantal/Uren")

        # button = self.wait_for_xpath(self.selectors["uren"])
        button.click()
        time.sleep(1)

        for i, date in enumerate(data):
            print("Processing homework data")
            input = self.wait_for_xpath(self.selectors["uren_date"])
            if isinstance(date[1], datetime):
                date_string = date[1].strftime("%d-%m-%Y")
            input.send_keys(date_string)

            self.select_dropdown_value("uren_type", date[6])

            if i == len(data) - 1:
                print("Done, saving last one")
                time.sleep(2)
                self.driver.find_element(
                    By.CSS_SELECTOR, self.selectors["save"]
                ).click()
                time.sleep(2)
            else:
                print("Save and new")
                time.sleep(2)
                self.driver.find_element(
                    By.CSS_SELECTOR, self.selectors["savenew"]
                ).click()
                time.sleep(2)

            self.driver.switch_to.default_content()
            self.wait_for_xpath(self.selectors["declaratie_iframe"])
            childIframe = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, self.selectors["declaratie_iframe"])
                )
            )
            self.driver.switch_to.frame(childIframe)

    def setup_web(self):
        try:
            self.setup_driver()

            if os.path.exists(self.cookie_file):
                print("Cookie file found. Attempting to load cookies...")

                if self.load_cookies():
                    if self.is_logged_in():
                        print("Successfully logged in using saved cookies!")
                    else:
                        print("Cookies didn't work. Manual login required.")
                        self.driver.get(self.url)
                        if self.wait_for_manual_login():
                            self.save_cookies()
                        else:
                            print("Login failed or timed out.")
                            return
                else:
                    print("Failed to load cookies. Manual login required.")
                    self.driver.get(self.url)
                    if self.wait_for_manual_login():
                        self.save_cookies()
                    else:
                        print("Login failed or timed out.")
                        return
            else:
                print("No cookie file found. First time login required.")
                self.driver.get(self.url)

                if self.wait_for_manual_login():
                    self.save_cookies()
                else:
                    print("Login failed or timed out.")
                    return

            print("\n=== Successfully connected to Salesforce! ===")

        except Exception as e:
            print(f"An error occurred: {e}")

    def wait_for_element_with_text(self, element_type, text_term, timeout=10):
        try:
            xpath = f"//{element_type}[contains(text(), '{text_term}')]"
            element = WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )

            return element

        except TimeoutException:
            print(
                f"Timeout: Element '{element_type}' with text '{text_term}' not found within {timeout} seconds"
            )
            return None
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            return None

    def run(self):
        data = self.load_xlsx_data()
        self.setup_web()
        self.run_forms(data)
        input("Press enter to exit....")
        self.driver.quit()


if __name__ == "__main__":
    automation = SalesforceAutomation()
    automation.run()
