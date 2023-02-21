from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import pandas as pd
import openpyxl
import sqlite3
import glob
import os.path


##DOWNLOAD FILE##

# Set the options for the Chrome webdriver
options = webdriver.ChromeOptions()
prefs = {'download.default_directory': '/Users/charm/Downloads/'}
options.add_experimental_option('prefs', prefs)

# Create a new instance of the Chrome webdriver
driver = webdriver.Chrome(options=options)

# Navigate to the page with the CSV download button
url = "https://jobs.homesteadstudio.co/data-engineer/assessment/download"
driver.get(url)

# Find the download button and click it
button = driver.find_element(By.LINK_TEXT, 'Download')
button.click()

# Wait for the download to complete
time.sleep(30)

# Close the webdriver
driver.quit()

# Get the latest downloaded file
file_path = '/Users/charm/Downloads/*.xlsx'
files = sorted(glob.iglob(file_path), key=os.path.getctime, reverse=True)
print(files[0])

#P#IVOT TABLE CREATION##

# Load the Excel file into a Pandas DataFrame object
data = pd.read_excel(files[0], sheet_name='data')

# Create a pivot table with the sum of the values in each column
pivot_table = pd.pivot_table(data, values=['Spend', 'Attributed Rev (1d)', 'Imprs', 'Visits', 'New Visits', 'Transactions (1d)', 'Email Signups (1d)'], index='Platform (Northbeam)', aggfunc='sum')
pivot_table_sorted = pivot_table.sort_values(by=['Attributed Rev (1d)'], ascending=False)


##SQLite DB Creation##

conn = sqlite3.connect('output_file.db')

# Store the pivot table in a new table in the database
pivot_table_sorted.to_sql('pivot_table', conn)

# Close the database connection
conn.close()