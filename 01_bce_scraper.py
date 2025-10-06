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
    df = df[df['TasaMaxima'] != '\n% \n      anual']
    df = df[df['TasaMaxima'] != '\n%\n  anual\n']
    df = df[df['TasaMaxima'] != '\n\xa0']
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
    elif year == '2015' and month not in ['08','09']:
        df.columns = ['Segmento2','TasaReferencial','Segmento','TasaMaxima']
    elif year == '2014' or year == '2013' or year == '2012' or year == '2011' or year == '2010':
        df.columns = ['Segmento2','TasaReferencial','Segmento','TasaMaxima']
    elif year == '2009' and month not in ['03','05','06','07','08','09','10','11','12']:
        df.columns = ['Segmento2','TasaReferencial','Segmento','TasaMaxima','Col5','Col6']
    elif year == '2009' and month in ['03']:
        df.columns = ['DummyCol','Segmento2','TasaReferencial','Segmento','TasaMaxima','Col6','Col7']
    elif year == '2009' and month in ['05','06','07','08','10','11','12']:
        df.columns = ['Segmento2','TasaReferencial','Segmento','TasaMaxima']
    elif year == '2009' and month in ['09']:
        df.columns = ['Segmento2','TasaReferencial','Segmento','TasaMaxima','Col5']
    elif year == '2008' and month in ['01']:
        df.columns = ['Segmento2','TasaReferencial','Segmento','TasaMaxima','Col5','Col6','Col7']
    elif year == '2008' and month in ['08','11']:
        df.columns = ['Segmento2','TasaReferencial','Segmento','TasaMaxima']
    elif year == '2008' and month not in ['01','08','11']:
        df.columns = ['Segmento2','TasaReferencial','Segmento','TasaMaxima','Col5','Col6']
    elif year == '2007':
        df.columns = ['Segmento','TasaMaxima']
    else:
        df.columns = ['Segmento','TasaMaxima','Col3','Col4']

    # Select relevant columns and clean the dataframe
    df = df[['Segmento','TasaMaxima']]
    df = delete_empty_rows(df) # Delete empty rows        
    df = format_dates(df)      # Format the dates in the dataframe
    
    # Return the clean dataframe
    return df

# Fn06: bce_interest_rates_scraper = Main function to scrap the interest rates from the BCE website
def bce_interest_rates_scraper(wd_path, year, month, wait_seconds):
    # Step 1. Call the selenium driver
    url = 'https://contenido.bce.fin.ec/documentos/Estadisticas/SectorMonFin/TasasInteres/TasasVigentes' + month + year + '.htm'
    driver = webdriver.Chrome()
    driver.get(url)
    wait = WebDriverWait(driver, wait_seconds)  # 10 is the timeout in seconds

    # Step 2. Check the table is in the format I neeed
    if year not in ['2023','2024','2025']:
        table_xpath = '/html/body/div/table[1]/tbody'
    elif year in ['2023'] and month in ['07']:
        table_xpath = '/html/body/div[4]/table/tbody'
    elif year in ['2023'] and month in ['08','09','10','11','12']:
        table_xpath = '/html/body/div/table[2]/tbody'
    elif year in ['2024']:
        table_xpath = '/html/body/div/table[2]/tbody'
    elif year in ['2025'] and month in ['01','02']:
        table_xpath = '/html/body/div/table[2]/tbody'
    elif year in ['2025'] and month not in ['01','02']:
        table_xpath = '/html/body/div/table[1]/tbody'
    
    table_test  = element_test(driver,table_xpath) # Check if the summary table exists
    if table_test == False:
        print("Unable to find the table in the original xpath provided, trying an alternative xpath")
        table_xpath = '/html/body/div/div/table/tbody'
    table_test  = element_test(driver,table_xpath) # Check if the summary table exists
    table_df    = get_table_df(driver,table_xpath)
    table_df    = format_table_01_df(table_df,year, month) # Format the table into a clean dataframe

    # Step 3. Save the table into a csv file
    file_path = wd_path + '\\data\\raw\\BCE_Max_Interest_Rates_' + year + month + '.xlsx'
    table_df.to_excel(file_path, index=False)
    print(f"The interest rates for the year {year} and month {month} have been saved to an excel file")

    return table_df

# Fn07 = append_excel_files = Append all the csv files in the raw data folder into a single dataframe
def append_excel_files(wd_path):
    folder_path = wd_path + "\\data\\raw"
    files_vector = os.listdir(folder_path) # Get all the files in the folder
    # 2. Initialize an empty list to store DataFrames
    dfs = []
    # 3. Iterate over each file in the directory
    for file in files_vector:
       if file.endswith('.xlsx'):
          file_path = os.path.join(folder_path, file) # Get the file path
          date_string = file[-11:]                    # Extract the date string from the file name
          df = pd.read_excel(file_path)
          df['FileDate'] = date_string                # Add a column with the file date
          dfs.append(df)
    # 4. Concatenate all DataFrames in the list into a single DataFrame
    final_df = pd.concat(dfs, ignore_index=True)
    return final_df



# 1. Call the scraper 
wait_seconds = 2  
year = '2025'
for month in range(1, 3):
    month = str(month).zfill(2)
    print(f"Scraping data for {year}-{month}")
    ir_df = bce_interest_rates_scraper(wd_path, year, month, wait_seconds)

# 2. Data cleansing
df = append_excel_files(wd_path)
seg_fn_df = pd.read_excel(wd_path + '\\data\\SegmentosFinalNames.xlsx')
df = df.merge(seg_fn_df, on='Segmento', how='left')
df = df[df['Keep'] == 'Yes' ] 
df = df.rename(columns={'Segmento': 'SegmentoRawName', 'SegmentoFinalName': 'SegmentoCleanName'})
new_order = ['SegmentoRawName', 'SegmentoCleanName', 'TasaMaxima' ,'Year', 'Month', 'YearMonth', 'Keep']
df = df[new_order]
df.to_excel(wd_path + '\\data\\clean\\BCE_Max_Interest_Rates.xlsx', index=False)