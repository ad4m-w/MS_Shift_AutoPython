import os
import sys
import pandas as pd
import pyperclip
import pyautogui
import threading
import urllib
from pathlib import Path
from splinter import Browser
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep

# Paths
csv_file_path = 'mvw.csv'
output_csv_path = 'activation_data.csv'

current_directory = os.getcwd()
subdirectory_name = "Phone HWID Scraper"
dynamic_path = os.path.join(current_directory, subdirectory_name)

cancel_button_flag = True
counter = 0
# Function to handle Cancel button click
def handle_cancel_button(dynamic_path):
    while cancel_button_flag:  
        try:
            button_location = pyautogui.locateOnScreen(f'{dynamic_path}/cancel.png', confidence=0.8)
            if button_location:
                button_center = pyautogui.center(button_location)
                pyautogui.click(button_center)
                print("Cancel button clicked!")
                sleep(4)  
        except Exception as e:
            sleep(2)  

# Function to get site type
def get_site_type(url):
    parsed_url = urllib.parse.urlparse(url)
    domain = parsed_url.netloc.replace("www.", "")
    
    site_mapping = {
        "msshift-demo4.com": "DEMO4",
        "msshift-mi3.com": "MI3",
        "msshift-mvw-2.com": "MVW-2",
        "mvw.msshift.com": "MVW",
        "msshift-security-6.com": "SEC-6",
        "msshift-security-2.com": "SEC-2",
        "msshift-security-3.com": "SEC-3",
        "msshift-hw.com": "HW",
        "msshift-mi-es.com": "MI-ES",
        "fs.msshift.com": "FS"
    }

    for domain_key, site_type in site_mapping.items():
        if domain_key in domain:
            return site_type
    return "Unknown Site Type"

# Read credentials from the credentials.txt file and store them in a dictionary
def load_credentials(file_path):
    credentials = {}
    if os.path.exists(file_path):
        with open(file_path, 'r') as f:
            for line in f:
                parts = line.strip().split(',')
                if len(parts) == 3:
                    site_type, username, password = parts
                    credentials[site_type] = (username.strip(), password.strip())
    else:
        print(f"Error: {file_path} not found!")
        sys.exit()
    
    return credentials

# Load credentials
credentials = load_credentials('credentials.txt')

# Check if CSV file exists
if not os.path.exists(csv_file_path):
    print(f"Error: {csv_file_path} not found!")
    sys.exit()

df = pd.read_csv(csv_file_path, encoding="UTF-8")
pCount = len(df.index)

# Path to the Chrome driver
if getattr(sys, 'frozen', False):
    bundle_dir = Path(sys._MEIPASS)
else:
    bundle_dir = Path(os.path.abspath(os.path.dirname(__file__)))

chrome_binary = r'C:/Program Files/Google/Chrome/Application/chrome.exe'
chromedriver_path = bundle_dir / 'chromedriver.exe'

if not chromedriver_path.exists():
    print("Error: chromedriver.exe not found!")
    sys.exit()

options = Options()
options.binary_location = chrome_binary
options.add_experimental_option('excludeSwitches', ['enable-logging'])
service = Service(str(chromedriver_path))
browser = Browser('chrome', service=service, options=options)

data = []

try:
    for index, row in df.iterrows():
        hotel_name = row['Hotel Name']
        hotel_link = row['Hotel Link']

        print(f"Visiting: {hotel_name}")
        counter+= 1
        print(f"Property {counter}/1070")
        cancel_thread = threading.Thread(target=handle_cancel_button, args=(dynamic_path,))
        cancel_thread.daemon = True  
        cancel_thread.start()

        site_type = get_site_type(hotel_link)
        browser.visit(hotel_link)
        sleep(1)

        if site_type in credentials:
            username, password = credentials[site_type]
        else:
            print(f"Error: Credentials for {site_type} not found in credentials.txt.")
            continue  

        if site_type in ['MI3', 'MI-ES']:
            browser.find_by_name("MsSSOBtn").click()
            browser.fill('pf.username', username)
            browser.fill("pf.pass", password)
            browser.find_by_id('SMBT').click()
            url = 'https://msshift-mi3.com/d/Admin/Permission/iPhoneSetup' if site_type == 'MI3' else 'https://msshift-mi-es.com/d/Admin/Permission/iPhoneSetup'

        elif site_type in ['SEC-2', 'SEC-3', 'SEC-6']:
            browser.fill('lgnBox', username)
            browser.fill('username', password)
            browser.find_by_name('Submit').click()
            url_mapping = {
                'SEC-2': 'https://msshift-security-2.com/d/Admin/Permission/iPhoneSetup',
                'SEC-3': 'https://msshift-security-3.com/d/Admin/Permission/iPhoneSetup',
                'SEC-6': 'https://msshift-security-6.com/d/Admin/Permission/iPhoneSetup'
            }
            url = url_mapping.get(site_type)

        elif site_type in ['MVW', 'MVW-2', 'HW', 'FS']:
            browser.fill('lgnBox', username)
            browser.fill('username', password)
            browser.find_by_name('Submit').click()
            url_mapping = {
                'MVW': 'https://mvw.msshift.com/d/Admin/Permission/iPhoneSetup',
                'MVW-2': 'https://msshift-mvw-2.com/d/Admin/Permission/iPhoneSetup',
                'HW': 'https://msshift-hw.com/d/Admin/Permission/iPhoneSetup',
                'FS': 'https://fs.msshift.com/d/Admin/Permission/iPhoneSetup'
            }
            url = url_mapping.get(site_type)

        sleep(2)
        browser.visit(url)
        WebDriverWait(browser.driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr/td/select")))
        WebDriverWait(browser.driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr/td/select")))
        dropdown_xpaths = [
            "/html/body/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr/td/select", 
            "/html/body/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr/td/select"
        ]

        for xpath in dropdown_xpaths:
            try:
                dropdown = browser.driver.find_element(By.XPATH, xpath)
                options = dropdown.find_elements(By.TAG_NAME, 'option')

                for option in options:
                    value = option.get_attribute('value')
                    label = option.text 

                    parts = label.split(' == ')
                    if len(parts) == 3:
                        hwid, activation_date, serial_number = parts
                        data.append({
                            'Hotel Name': hotel_name,
                            'HWID': hwid,
                            'Activation Date': activation_date,
                            'Serial Number': serial_number
                        })

            except Exception as e:
                print(f"Error extracting data from dropdown at {xpath} for {hotel_name}: {e}")

except Exception as e:
    print(f"Error during script execution: {e}")
    import traceback
    traceback.print_exc()

finally:
    try:
        df_output = pd.DataFrame(data)
        df_output.to_csv(output_csv_path, index=False, encoding="UTF-8")
        print(f"Activation data saved to {output_csv_path}")
    except Exception as e:
        print(f"Error saving CSV: {e}")
    finally:
        cancel_button_flag = False
        print("Process complete. Press Enter to close the browser...")
        input()
        browser.quit()