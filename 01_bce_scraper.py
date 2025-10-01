                                # BCE Max. Interest Rates Scraper
# Objective: 1. Get the maximum interest rates for each credit type from the BCE website 
# Author: Nicolas Chuquimarca-Arguello (IDB)
# version 0.1: 2025-10-01: First version

# PENDING WORK: Make the code work for 2015 and before. Generalize the exceptions.


# 0. Packages
import os, time, pyttsx3, requests, io, cv2, math, pandas as pd, numpy as np
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException # Use the try & except to handle the selenium exceptions
from selenium.common.exceptions import TimeoutException       # Use the try & except to handle the selenium exceptions
from datetime import date                                     # Get the current date
from datetime import datetime                                 # Get the current date and time                              

# 0 Set working directory
wd_path = "C:\\Users\\nicoc\\Dropbox\\IDB\\InterestRatesBCE"
os.chdir(wd_path)

# Functions
# Fn01: element_test = Check whether an element exists in the page or not
def element_test(driver,xpath):
    try:
        element = driver.find_element(By.XPATH,xpath)
        element_test = True
    except NoSuchElementException:
        element_test = False
    return element_test

# Fn02: get_table_df = Get the table from the xpath and return it into a dataframe
def get_table_df(driver,xpath):
    table_element = driver.find_element(By.XPATH, xpath)
    table_html = table_element.get_attribute('outerHTML')
    soup = BeautifulSoup(table_html, 'html.parser')    # Parse the HTML content with BeautifulSoup
    table = [[cell.text for cell in row.find_all('td')] for row in soup.find_all('tr')] # Extract the table rows and columns with BeautifulSoup
    df = pd.DataFrame(table) # Convert the extracted data into a pandas DataFrame
    return df

# Fn03: delete_empty_rows = Delete empty rows from a dataframe
def delete_empty_rows(df):
    df = df[df['TasaMaxima'].notna()]        
    df = df[df['TasaMaxima'] != '% anual']
    df = df[df['TasaMaxima'] != '%\n  anual']
    df = df[df['TasaMaxima'] != '\n% anual\n']
    df = df[df['TasaMaxima'] != '']
    df = df[df['TasaMaxima'] != 'Tasas Máximas']
    df = df[df['TasaMaxima'] != 'para el segmento:']
    df = df[df['TasaMaxima'] != ' ']
    df = df[df['TasaMaxima'] != '\n \n'] 
    df = df[df['Segmento'] != ' ']
    return df

# Fn04: format_dates = Format the dates in the dataframe
def format_dates(df):
    # Convert the interest rate to numeric and add Year and Month columns
    df['TasaMaxima'] = df['TasaMaxima'].str.replace(',', '.').astype(float)
    df['Year'] = int(year)
    df['Month'] = int(month)
    df = df.reset_index(drop=True) # Reset the index
    df['YearMonth'] = df['Year'].astype(str) + '-' + df['Month'].astype(str).str.zfill(2) 
    return df

# Fn05: format_table01_df = Format the table into a clean dataframe (Valid from 2023-07 to 2025-09)
def format_table_01_df(df, year, month):
    # Keep only the relevant columns and remove NaN values
    if year == '2022' and month in ['01','02','03','04','05','06', '07']:
        df.columns = ['Segmento2','TasaReferencial','Segmento','TasaMaxima']
    elif year == '2021' or year == '2020' or year == '2019' or year == '2018' or year == '2017' or (year == '2016' and month not in ['01','02','06']):
        df.columns = ['Segmento2','TasaReferencial','Segmento','TasaMaxima']
    elif year == '2016' and month in ['01','02','06']:
        df.columns = ['Segmento2','TasaReferencial','Segmento','TasaMaxima','DummyCol']
    elif year == '2015':
        df.columns = ['Segmento2','TasaReferencial','Segmento','TasaMaxima']
    else:
        df.columns = ['Segmento','TasaMaxima','Col3','Col4']

    # Select relevant columns and clean the dataframe
    df = df[['Segmento','TasaMaxima']]
    df = delete_empty_rows(df) # Delete empty rows        
    df = format_dates(df)      # Format the dates in the dataframe
    
    # Return the clean dataframe
    return df

# Fn04: bce_interest_rates_scraper = Main function to scrap the interest rates from the BCE website
def bce_interest_rates_scraper(wd_path, year, month, wait_seconds):
    # Step 1. Call the selenium driver
    url = 'https://contenido.bce.fin.ec/documentos/Estadisticas/SectorMonFin/TasasInteres/TasasVigentes' + month + year + '.htm'
    driver = webdriver.Chrome()
    driver.get(url)
    wait = WebDriverWait(driver, wait_seconds)  # 10 is the timeout in seconds

    # Step 2. Check the table is in the format I neeed
    table_xpath = '/html/body/div/table[1]/tbody'
    table_test  = element_test(driver,table_xpath) # Check if the summary table exists
    if table_test == False:
        table_xpath = '/html/body/div/div/table/tbody'
    table_test  = element_test(driver,table_xpath) # Check if the summary table exists
    table_df    = get_table_df(driver,table_xpath)
    table_df    = format_table_01_df(table_df,year, month) # Format the table into a clean dataframe

    # Step 3. Save the table into a csv file
    file_path = wd_path + '\\data\\raw\\BCE_Max_Interest_Rates_' + year + month + '.xlsx'
    table_df.to_excel(file_path, index=False)
    print(f"The interest rates for the year {year} and month {month} have been saved to an excel file")

    return table_df

# General parameters
wait_seconds = 2  # Seconds to wait for the page to load
year = '2015'
for month in range(1, 13):
    month = str(month).zfill(2)
    print(f"Scraping data for {year}-{month}")
    ir_df = bce_interest_rates_scraper(wd_path, year, month, wait_seconds)

ir_df.shape[1]