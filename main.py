import selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options

from datetime import date, timedelta
from email.message import EmailMessage

import glob
import time
import shutil
import smtplib
import os
import base64
import email
import smtplib
import imaplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import email
import smtplib
import ssl
import csv
import openpyxl
import logging

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import (
    Fill,
    Font,
    PatternFill,
    Border,
    Side,
    Alignment,
    Protection,
    colors,
)
from copy import copy
from openpyxl.styles.colors import Color

# get username
import getpass

# Pandas & other shit!
import pandas as pd
import numpy as np
from pandas import ExcelWriter
import tkinter as tk
from tkinter import filedialog
import ntpath


def login_cdms(username, pass1, pass2):
    try:

        chrome_options =  Options()
        chrome_options.add_experimental_option("detach", True)
        chrome_options.page_load_strategy = 'normal'
        url = "https://cdms.exxonmobil.com"
        path_driver = os.path.join(os.path.dirname(__file__), "chromedriver.exe")
        browser = webdriver.Chrome(executable_path=path_driver, options=chrome_options)
        browser.get(url)
        user = username
        password = pass1
        password2 = pass2

        signing_button = '//*[@id="loginForm"]/div[7]/a'

        box1 = WebDriverWait(browser, 10).until(
            selenium.webdriver.support.expected_conditions.presence_of_element_located((By.NAME, "pf.username")))

        box1.clear()

        box1.send_keys(user)

        box2 = WebDriverWait(browser, 10).until(
            selenium.webdriver.support.expected_conditions.presence_of_element_located((By.NAME, "pf.pass"))
        )

        box2.clear()

        box2.send_keys(password)

        browser.find_element_by_xpath(signing_button).click()

        time.sleep(1)

        guid = browser.window_handles

        browser.switch_to.window(guid[-1])

        dd_data_source = browser.find_element_by_id("ddlDataSource_Input")

        dd_data_source.click()
        time.sleep(1)
        dd_options = browser.find_element_by_id(
            "ddlDataSource_DropDown"
        ).find_element_by_css_selector("ul")

        for nice_option in dd_options.find_elements_by_css_selector("li"):

            if nice_option.get_attribute("textContent") == "Production-SCL":
                nice_option.click()
                break
        time.sleep(2)

        # user send keys
        box4 = WebDriverWait(browser, 10).until(
            selenium.webdriver.support.expected_conditions.presence_of_element_located((By.NAME, "T1"))
        )

        box4.send_keys(user)

        # password send keys
        box5 = WebDriverWait(browser, 10).until(
            selenium.webdriver.support.expected_conditions.presence_of_element_located((By.NAME, "T2"))
        )

        box5.send_keys(password2)

        browser.find_element_by_id("Submit2").click()
        print("Login Successful")
        time.sleep(2)
    except Exception as e:
        print(e)
        browser.quit()
        login_cdms(username="SWPabourg", pass1="31Forestavenue*", pass2="02forestavenue")


def email_auto_script(to_email, subject, body, file_path):
    print('Sending Email!!!')

    gmail_user = "graham.scripting@gmail.com"
    gmail_password = "directorchris"

    try:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = gmail_user
        msg["To"] = to_email
        msg.set_content(body)

        for file in file_path:
            print(file)
            with open(file, "rb") as f:
                file_data = f.read()
                file_name = os.path.split(file)
                file_name = file_name[-1]

                msg.add_attachment(
                    file_data,
                    maintype="application",
                    subtype="vnd.openxmlfo  mats-officedocument.spreadsheetml.sheet",
                    filename=file_name,
                )

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(gmail_user, gmail_password)
            smtp.send_message(msg)
            print("Email sent!")
    except:
        print("Something went wrong...Email not Sent.")

def labour_ovr(first_name, last_name, rt, ot, dot, ovr_reason, comment, handler):
    print("in LABOUR OVR")


if __name__ == '__main__':
    login_cdms(username="SWPabourg", pass1="31Forestavenue*", pass2="02forestavenue")
    # file_path = []
    # email_to = 'momin.ejaz@graham.ca'
    # subject = 'Test 1'
    # body = "Hey Momin!\nThis is a test to see if your function works!!\n\nRegards\nGraham Scripting"
    # file_path = [r'C:\Users\momine\Documents\Python\CDMS\Charm New Project\Area List CDMS.xlsx']
    #
    #email_auto_script(email_to, subject, body, file_path)


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
