from tkinter import filedialog
import os
import pandas as pd
from bs4 import BeautifulSoup

# Load HTML file# Load HTML file
 #open an explorer window to select the file
file = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("html files","*.html"),("all files","*.*")))
#read as utf-8
with open(file, 'r', encoding='utf-8') as f:
    html = f.read()
    
# Parse HTML using BeautifulSoup
soup = BeautifulSoup(html, 'html.parser')

# Find all table elements
table_elements = soup.find_all('table')
#remove empty tables that say No data available in table
table_elements = [table for table in table_elements if table.find('td').text != 'No data available in table']
#if the table is empty  remove it
table_elements = [table for table in table_elements if table.find('td').text != '']

# Extract table names and dataframes
tables = {}
for table_element in table_elements:
    # Extract table name from table element
    table_name_element = table_element.find_previous('strong')
    table_name = table_name_element.text.strip() if table_name_element else 'Unknown'
    print(table_name)
    #remove duplicate table names
    # Convert table to DataFrame
    df = pd.read_html(str(table_element))[0]

    # Add table to dictionary
    tables[table_name] = df
#output the tables to csv file with a sheet for each table
writer = pd.ExcelWriter(file + '.xlsx', engine='xlsxwriter')
for table_name, df in tables.items():
	df.to_excel(writer, sheet_name=table_name, index=False)

#make each sheet in to a table in the excel file
workbook = writer.book
for sheet in writer.sheets.values():
    worksheet = sheet
    worksheet.autofilter(0, 0, len(df), len(df.columns))
    worksheet.freeze_panes(1, 1)
    worksheet.set_column(0, len(df.columns), 20)
    
#save the report to the same folder as the html file
writer._save()
#open the excel file
os.startfile(file+ '.xlsx')
