# By Adam Waszczyszak

import os
import sys
import time
import psutil
import tkinter as tk
from tkinter import Button
from tkinter import filedialog, ttk
from tkinter import WORD
from tkinter import Text
from tkinter import END
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from pathlib import Path
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains

def clear_screen():
    if os.name == 'nt':
        _ = os.system('cls')
    else:
        _ = os.system('clear')

def kill_adobe():
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        try:
            if proc.info['name'] == 'AcroRd32.exe':
                proc.kill()  
                print(f"Killed process {proc.info['name']} with PID {proc.info['pid']}")
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

def select_printer(printer_name):
    options = Options()
    options.add_experimental_option("detach", True)
    options.add_argument("--headless")  # Runs Chrome headless, in background

    if getattr(sys, 'frozen', False):
        bundle_dir = Path(sys._MEIPASS)  
    else:
        bundle_dir = Path(os.path.abspath(os.path.dirname(__file__)))

    chrome_binary = r'C:/Program Files/Google/Chrome/Application/chrome.exe'
    chromedriver_path = bundle_dir / 'chromedriver.exe'

    options.binary_location = chrome_binary
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    if not chromedriver_path.exists():
        print("Error: chromedriver.exe not found!")
        return

    s = Service(str(chromedriver_path))
    browser = webdriver.Chrome(service=s, options=options)

    # URL and login credentials
    url = 'https://msshift-demo-4.com/default?sk=%2Bek%2BXTAHiLXY1CXbHuFEijToLeUSdd2q8eSzRLlkzTI%3D'
    username = "3"
    password = "3"
    
    # Navigate and perform actions based on selected printer
    browser.get(url)
    browser.find_element(By.ID, "lgnBox").send_keys(username)
    browser.find_element(By.ID, "username").send_keys(password)
    browser.find_element(By.NAME, "Submit").click()
    print(f"\nLogged-in successfully for {printer_name}...")

    # URL to navigate for selecting printer
    url = 'https://msshift-demo-4.com/d/EventLog/VisitorsNew/VisitorMenu?Top=8'
    browser.get(url)
    browser.find_element(By.LINK_TEXT, "Quick Visitor Search").click()
    url = 'https://msshift-demo-4.com/d/EventLog/VisitorsNew/AssignVisitor/Expand_Visitor_Note?features=scrollbars=yes,width=1015,height=875,top=5,left=450&IDNum=25-215&rf=rf&VisitorID=1206#TopPage'
    browser.get(url)
    browser.find_element(By.XPATH, "//input[contains(@onclick, 'show_printer_settings2')]").click()
    browser.find_element(By.ID, "printersList2").click()

    # Select printer based on the user choice
    if printer_name == "DYMO 550":
        browser.find_element(By.XPATH, "//option[text()='DYMO LabelWriter 550 Turbo']").click()
    elif printer_name == "DYMO 450":
        browser.find_element(By.XPATH, "//option[text()='DYMO LabelWriter 450']").click()
    elif printer_name == "LX-500":
        browser.find_element(By.XPATH, "//option[text()='Color Label 500']").click()
    elif printer_name == "Microsoft Print to PDF":
        browser.find_element(By.XPATH, "//option[text()='Microsoft Print to PDF']").click()
    elif printer_name == "ZD421":
        select_element = browser.find_element(By.ID, "printersList2")
        select = Select(select_element)
        select.select_by_visible_text("ZDesigner ZD421-203dpi ZPL")

    browser.find_element(By.XPATH, "//input[contains(@onclick, 'save_printer2()')]").click()
    browser.implicitly_wait(2)
    browser.find_element(By.ID, "qpbtn").click()
    time.sleep(5)
    kill_adobe()
    browser.quit()

def center_window(window):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - 275) // 2
    y = (screen_height - 300) // 2
    window.geometry(f"{275}x{300}+{x}+{y}")

master = tk.Tk()
master.geometry("250x300")
center_window(master)
master.title("adamwasz.com")

# GUI
button_1 = Button(master, text="DYMO 550", command=lambda: select_printer("DYMO 550"), width=30)
button_1.grid(row=0, column=0, pady=10, padx=20, sticky="nsew")

button_2 = Button(master, text="DYMO 450", command=lambda: select_printer("DYMO 450"), width=30)
button_2.grid(row=1, column=0, pady=10, padx=20, sticky="nsew")

button_3 = Button(master, text="LX-500", command=lambda: select_printer("LX-500"), width=30)
button_3.grid(row=2, column=0, pady=10, padx=20, sticky="nsew")

button_4 = Button(master, text="Microsoft Print to PDF", command=lambda: select_printer("Microsoft Print to PDF"), width=30)
button_4.grid(row=3, column=0, pady=10, padx=20, sticky="nsew")

button_5 = Button(master, text="ZD421", command=lambda: select_printer("ZD421"), width=30)
button_5.grid(row=4, column=0, pady=10, padx=20, sticky="nsew")

text_box = Text(master, height=2, width=30, wrap=WORD)
text_box.grid(row=5, column=0, columnspan=2, pady=10, padx=10)
text_box.insert(END, "Select a printer to test...\nGUI may freeze, be patient.")  

master.mainloop()
