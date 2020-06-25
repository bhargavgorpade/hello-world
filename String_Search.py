# -*- coding: utf-8 -*-
"""
Created on Wed Jun 24 21:48:37 2020

@author: bgorpade
"""
#---------------------------------------STRING - SEARCH-----------------------------------------
#CHECKLIST BEFORE RUNNING THE SCRIPT 
#   1. Python 3.6 or above must be installed.
#   2. ODBC Driver 17 for SQL Server must be installed.
#   3. Make sure the required libraries are already installed, If not -
#        installation command is 'pip install <library_name>'
#   4. Create a new excel file at the desired output path.
#-----------------------------------------------------------------------------------------------
#importing required libraries...
import pyodbc 
import itertools
import pandas as pd
from openpyxl import load_workbook

#Input data - Strings to be modified...
server = 'servsqlqa.database.windows.net' #The Server in which you want to search
database = 'CIRRUS'                       #The DB in which you want to search
column = 'DEPT'                           #Enter Column Name
search_str = 'InvestorServices'           #Enter string to be searched for 

#CREATE A NEW EXCEL FILE AT DESIRED PATH AND PUT THE PATH BELOW to op_path variable
op_path = r"C:\Users\bgorpade\Documents\Python Scripts\OUTPUT.xlsx"

#Establishing ODBC Connection and retrieving all table names into 'final_result' list...
conn_str = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';Trusted_Connection=yes;'
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()
cursor.execute('SELECT TABLE_SCHEMA,TABLE_NAME FROM cirrus.INFORMATION_SCHEMA.TABLES;')
result = cursor.fetchall()
final_result = [list(i) for i in result]

#Parsing through each table in 'final_result'...
for combination in final_result:
    #Getting All column names for a particular table...
    cursor.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '"+combination[0]+"' AND TABLE_NAME = '"+combination[1]+"';")
    result = cursor.fetchall()
    column_result = [list(i) for i in result]
    column_result_flat = list(itertools.chain.from_iterable(column_result))
    #Checking if input column name is present in the particular table...
    if column in column_result_flat:
        db_string = combination[0]+"."+combination[1]
        qry  = "SELECT * FROM "+db_string+" WHERE "+column+" = '"+search_str+"';"
        df = pd.read_sql_query(qry, conn)
        #Loading query results to already created Excel 
        book = load_workbook(op_path)
        writer = pd.ExcelWriter(op_path, engine = 'openpyxl')
        writer.book = book
        df.to_excel(writer, sheet_name = db_string)
        writer.save()
        writer.close()
        
    
    

