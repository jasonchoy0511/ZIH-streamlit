import streamlit as st
import pandas as pd
import numpy as np
import time
import datetime
from glob import glob
from pathlib import Path
import xlwings as xw
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert
from webdriver_manager.chrome import ChromeDriverManager


def Main(today, Downloads):
    def GetEgFile():
        User = chrome.find_element_by_id("usrname")
        User.send_keys("zen")
        time.sleep(0.5)
        Password = chrome.find_element_by_id("usrpass")
        Password.send_keys("zen123")
        time.sleep(0.5)
        LoginButton = chrome.find_element_by_id("btn_login")
        LoginButton.click()
        time.sleep(10)
        url = 'http://www.e-gain.com.hk:8088/ittms/ittms.php/ctrl_lot_enquiry'
        chrome.get(url)
        time.sleep(10)
        Search = chrome.find_element_by_id("f_lot_enquiry_btn_search")
        Search.click()
        time.sleep(10)
        chrome.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        time.sleep(5)
        Export = chrome.find_element_by_xpath(
            '//*[@id="table_lot_enquiry_browse_pager"]/div/span[1]/span/span[2]')
        Export.click()
        time.sleep(5)
        chrome.quit()

    options = Options()
    options.add_argument("--disable-notifications")
    chrome = webdriver.Chrome(
        ChromeDriverManager().install(), chrome_options=options)
    url = 'http://www.e-gain.com.hk:8088/ittms/ittms.php'
    chrome.get(url)
    time.sleep(1)
    GetEgFile()
    time.sleep(2)

    filepaths = glob(f'{Downloads}\*.xls')
    filepath = filepaths[0]
    wb = xw.Book(filepath)
    wb.save(f'{Downloads}\EG.xlsx')
    wb.close()
    file = Path(filepath)
    file.unlink()

    ExcelPaths = glob(f'{Downloads}\EG.xlsx')
    ExcelPath = ExcelPaths[0]
    df = pd.read_excel(ExcelPath)
    df = df.iloc[3:, :].T.reset_index()
    df = df.iloc[:, 1:].T
    UseCol = [0, 1, 3, 23, 6]
    _df = df[UseCol]
    _df = _df.rename(columns={0: 'Item No.', 1: 'Receive Date', 3: 'Lot No.',
                              23: 'Description', 6: 'Available Quantity'})
    _df2 = df[12]
    _df2 = _df2.str.split(' ', expand=True)
    _df2 = _df2.rename(columns={0: 'Expiry Date'})
    _df2 = _df2['Expiry Date']
    df = pd.concat([_df, _df2], axis=1)
    df = df[['Item No.', 'Receive Date', 'Lot No.',
             'Description', 'Expiry Date', 'Available Quantity']]
    df = df.sort_values(['Receive Date', 'Lot No.'])
    df.to_excel(f"{Downloads}\Inventory_{today}.xlsx", index=False)
    EG = Path(ExcelPath)
    EG.unlink()
    return df


today = datetime.date.today()
date = today + datetime.timedelta(days=1)
date = str(date)
home = str(Path.home())
Downloads = f'{home}\\Downloads'


st.title("ZEN-NOH EGG TEAM")
st.sidebar.write("Download Inventory")
button = st.sidebar.button("Download")
if button:
    Main(today, Downloads)

if st.sidebar.checkbox("Show EG Inventory"):
    df = pd.read_excel(f"{Downloads}\Inventory_{today}.xlsx")
    st.dataframe(df, width=3000, height=500)
