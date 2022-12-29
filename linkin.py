from tkinter import *
from tkinter import ttk
import pandas as pd
import numpy as np
import json
import time
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait as W
from urllib.request import urlopen
from selenium import webdriver
import xlrd
from openpyxl import load_workbook
from xlutils.copy import copy

s = Service("C:\\Users\\ASUS\\Downloads\\chromedriver_win32\\chromedriver.exe")
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress","127.0.0.1:9111")
driver = webdriver.Chrome(service=s)

df1 = pd.DataFrame()
df2 = pd.DataFrame()
df3 = pd.DataFrame()

def clicked():
    driver.get('https://www.linkedin.com/')
    time.sleep(10)

def start():
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(3)
    print(driver.current_url)
    #print(driver.window_handles)
    UAE = driver.find_elements(By.TAG_NAME, 'a')
    print(UAE)
    for a in UAE:
        if a.text == 'Hospitality 1 UAE':
            a.click()
            break

def direct_html():
    # driver.get('file:///C:/Users/ASUS/PycharmProjects/pythonProject/test_table.html')
    n1 = driver.find_elements(By.TAG_NAME, "table")
    #print("...........table starts here..............")
    for table in n1:
        headings =[]
        data_head = table.find_elements(By.TAG_NAME, 'thead')
        for dh in data_head:
            d2 = dh.find_elements(By.TAG_NAME, 'tr')
            for d3 in d2:
                print("......HEADINGS ...............")
                d4 = d3.find_elements(By.TAG_NAME, 'th')
                for dt in d4:
                    headings.append(dt.text)
        print("headings",headings)
        values = []
        data_body = table.find_elements(By.TAG_NAME, 'tbody')
        for dh in data_body:
            d2 = dh.find_elements(By.TAG_NAME, 'tr')
            for d3 in d2:
                d4 = d3.find_elements(By.TAG_NAME, 'td')
                index_no = 0
                row = {}
                for dt in d4:
                    head_name = headings[index_no]
                    if head_name == "Name":
                        C_name = dt.text.split("\n")
                        # print('c_name')
                        print(C_name)
                        try:
                            Name = C_name[0]
                        except:
                            Name = C_name
                        # links = dt.find_elements(By.TAG_NAME, 'a')
                        # print('links elements')
                        # print(links)
                        # for lin in links:
                        #     link = lin.get_attribute('href')
                        #     print("Printing Links")
                        #     print(link)
                    row[head_name] = dt.text
                    index_no += 1
                    # values.append(row)
                    #Getting the clients links
                #for dt in d4:
                link = 'Null'
                try:
                    links = d4[0].find_elements(By.TAG_NAME, 'a')[0]
                    print(links)
                    link = links.get_attribute('href')
                    print("Printing Links")
                    print(link)
                except:
                    link = 'Null'
                row['Profile_link'] = link

                values.append(row)

                    # for lin in links:
                    #     link = lin.get_attribute('href')
                    #     print("Printing Links")
                    #     print(link)
                    #     break
        print('values')
        print(values)
        # for dl in dt:
        #     links = dl.find_elements(By.TAG_NAME, 'a')
        #     print(links)
        #     for lin in links:
        #         link = lin.get_attribute('href')
        #         print("Printing Links")
        #         print(link)
    df0 = pd.json_normalize(values)
    return df0


def Leads_data():

    count = 12
    final_df_list = []
    for i in range(count):
        print("Page No ")
        print(i)
        df1 = direct_html()
        final_df_list.append(df1)
        time.sleep(5)
        abc = driver.find_element(By.XPATH, "(//span[normalize-space()='Next'])[1]")   #Next Button path
        print(abc)
        abc.location_once_scrolled_into_view
        try:
            abc.click()
        except:
            break
        time.sleep(20)
    final_df = pd.concat(final_df_list,ignore_index=True)
    print(final_df)
    Excel = 'LinkedIn.xlsx'
    final_df.to_excel(Excel)




window = Tk()
window.title(" LogIn for LinkedIn using Web Scrapping ")
window.geometry('600x100')
window.configure(background="white")
label_file_explorer = Label(window, text=" Signin to LinkedIn Account ",
                            width=100, height=3,
                            fg="blue")
button_exit = Button(window,
                     text=" Login ",
                     command=clicked)
Leads = Button(window,
                     text=" start Leads",
                     command=start)
page1 = Button(window,
                     text=" Extract Leads Data ",
                     command=Leads_data)
label_file_explorer.grid(column=1, row=1)
button_exit.grid(column=1, row=2)
Leads.grid(column=1,row=3)
page1.grid(column=1, row=4)
window.mainloop()