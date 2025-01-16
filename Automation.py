# By Adam Waszczyszak
from os import system
system("title " + "Written By Adam Waszczyszak")

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
import os
import time
import sys
from pathlib import Path

def clear_screen():
    # For Windows
    if os.name == 'nt':
        _ = os.system('cls')
    # For macOS and Linux
    else:
        _ = os.system('clear')

clear_screen()
print("Choose a printer to test print from:")
print("1. DYMO 550")
print("2. DYMO 450")
print("3. LX-500")
print("4. Microsoft Print to PDF (for testing)")
print("5. ZD421")
print("6. Exit")
choice = int(input("Enter your choice: "))

if(choice == 6):
    print("Exiting...")
    exit()

options = Options()
options.add_experimental_option("detach", True)

# Check if the script is running as a packaged executable (PyInstaller)
if getattr(sys, 'frozen', False):
    # Running as a bundled executable
    bundle_dir = Path(sys._MEIPASS)  # This is where PyInstaller places the bundled files
else:
    # Running as a script
    bundle_dir = Path(os.path.abspath(os.path.dirname(__file__)))

# Use the correct path for chrome binary and chromedriver
chrome_binary = r'C:/Program Files/Google/Chrome/Application/chrome.exe'
chromedriver_path = bundle_dir / 'chromedriver.exe'

options.binary_location = chrome_binary
options.add_experimental_option('excludeSwitches', ['enable-logging'])

# Ensure chromedriver is in the same directory as the script (or bundled in with PyInstaller)
if not chromedriver_path.exists():
    print("Error: chromedriver.exe not found!")
    exit()

s = Service(str(chromedriver_path))
browser = webdriver.Chrome(service=s, options=options)

# Security2 URL for now, until Demo4 is fixed
url = 'REPLACED FOR SECURITY'

username = "REPLACED FOR SECURITY"
password = "REPLACED FOR SECURITY"

browser.get(url)
browser.find_element(By.ID, "lgnBox").send_keys(username)
browser.find_element(By.ID, "username").send_keys(password)
browser.find_element(By.NAME, "Submit").click()
print("\nLogged-in successfully...")

# The main logic continues based on the user's choice
if(choice == 1):
    url = 'REPLACED FOR SECURITY'
    browser.get(url)
    browser.find_element(By.LINK_TEXT, "Quick Visitor Search").click()
    url = 'REPLACED FOR SECURITY'
    browser.get(url)
    browser.find_element(By.XPATH, "//input[contains(@onclick, 'show_printer_settings2')]").click()
    browser.find_element(By.ID, "printersList2").click()
    browser.find_element(By.XPATH, "//option[text()='DYMO LabelWriter 550 Turbo']").click()
    browser.find_element(By.XPATH, "//input[contains(@onclick, 'save_printer2()')]").click()
    browser.implicitly_wait(2)
    browser.find_element(By.ID, "qpbtn").click()

elif(choice == 2):
    url = 'https://msshift-security-2.com/d/eventlog/VisitorsNew/VisitorMenu?Top=8'
    browser.get(url)
    browser.find_element(By.LINK_TEXT, "Quick Visitor Search").click()
    url = 'REPLACED FOR SECURITY'
    browser.get(url)
    browser.find_element(By.XPATH, "//input[contains(@onclick, 'show_printer_settings2')]").click()
    browser.find_element(By.ID, "printersList2").click()
    browser.find_element(By.XPATH, "//option[text()='DYMO LabelWriter 450']").click()
    browser.find_element(By.XPATH, "//input[contains(@onclick, 'save_printer2()')]").click()
    browser.implicitly_wait(2)
    browser.find_element(By.ID, "qpbtn").click()

elif(choice == 3):
    url = 'REPLACED FOR SECURITY'
    browser.get(url)
    browser.find_element(By.LINK_TEXT, "Quick Visitor Search").click()
    url = 'REPLACED FOR SECURITY'
    browser.get(url)
    browser.find_element(By.XPATH, "//input[contains(@onclick, 'show_printer_settings2')]").click()
    browser.find_element(By.ID, "printersList2").click()
    browser.find_element(By.XPATH, "//option[text()='Color Label 500']").click()
    browser.find_element(By.XPATH, "//input[contains(@onclick, 'save_printer2()')]").click()
    browser.implicitly_wait(2)
    browser.find_element(By.ID, "qpbtn").click()

elif(choice == 4):
    url = 'REPLACED FOR SECURITY'
    browser.get(url)
    browser.find_element(By.LINK_TEXT, "Quick Visitor Search").click()
    url = 'REPLACED FOR SECURITY'
    browser.get(url)
    browser.find_element(By.XPATH, "//input[contains(@onclick, 'show_printer_settings2')]").click()
    browser.find_element(By.ID, "printersList2").click()
    browser.find_element(By.XPATH, "//option[text()='Microsoft Print to PDF']").click()
    browser.find_element(By.XPATH, "//input[contains(@onclick, 'save_printer2()')]").click()
    browser.implicitly_wait(2)
    browser.find_element(By.ID, "qpbtn").click()

elif(choice == 5):
    url = 'REPLACED FOR SECURITY'
    browser.get(url)
    browser.find_element(By.LINK_TEXT, "Quick Found/Return Search").click()
    url = 'REPLACED FOR SECURITY'
    browser.get(url)
    browser.find_element(By.XPATH, "//input[contains(@onclick, 'show_printer_settings2')]").click()
    browser.find_element(By.ID, "printersList2").click()
    select_element = browser.find_element(By.ID, "printersList2")
    select = Select(select_element)
    select.select_by_visible_text("ZDesigner ZD421-203dpi ZPL")
    browser.find_element(By.XPATH, "//input[contains(@onclick, 'save_printer2()')]").click()
    browser.find_element(By.ID, "qpbtn").click()

# Comment this out if debugging
time.sleep(10)
browser.quit()
