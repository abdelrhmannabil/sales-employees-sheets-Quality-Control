# -*- coding: utf-8 -*-
"""
Created on Thu Jan 18 12:36:44 2024
We will enter google sheets of sales employees to check functions
@author: abdelrahman.nabil
"""
import pip                                                            #module for piping and installing other modules
import time                                                           #module for overcome quota of requests of google to pause for seconds between every request
import gspread                                                        #module for using google API 
import pandas as pd                                                   #module for working with data frames and data manipulation
from oauth2client.service_account import ServiceAccountCredentials    #module for authentication to access google sheets

#M:/تحليل البيانات/statistical analysis with r/sales/employee sales report/sales form functions/banded-torus-337718-6ba11f3822b5.json
# define the scope


'M:/تحليل البيانات/statistical analysis with r/sales/employee sales report/sales form functions/exp.xlsx' #paht for sheet containing employees sheets links on google drive
wb_data = pd.ExcelFile('M:/تحليل البيانات/statistical analysis with r/sales/employee sales report/sales form functions/Employees Links.xlsx') #importing sheet of employees sheets links on google drive
df = pd.read_excel(wb_data, index_col=None, na_values=["NA"]) #converting the file to dataframe
for link in df['Link']:                  #for loop for automation of reading all employees google sheets
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

    # add credentials to the account
    creds = ServiceAccountCredentials.from_json_keyfile_name('M:/تحليل البيانات/statistical analysis with r/sales/employee sales report/sales form functions/banded-torus-337718-6ba11f3822b5.json', scope)

    # authorize the clientsheet 
    client = gspread.authorize(creds)
    wks =  client.open_by_url(link)  #reading main workbook
    emp_name_form = wks.worksheet("Closed Deals") #importing first sheet for extraction of employee name for modifying some function to search for the employee name
    emp_name = emp_name_form.acell('C1').value  #getting value of employee name from the sheet

    dashboard_salesDB = wks.worksheet("Dashboard Sales DB")   #importing another sheet of interest
    salesDB = wks.worksheet("Sales DB")                       
    for_powerBI = wks.worksheet("For power bi")               #accessing the sheet we need to modify values of some sales that contains some function
    target = wks.worksheet("Target")                          #accessing Target sheet to modify a cell containing query to import sales data and manipulate it to present for employee
    

    for i in range(2,4):
        for_powerBI.update('AT{}'.format(i), '=IF(\'Sales Dashbaord\'!$J$2 = "" , TRUE , IF(\'Sales Dashbaord\'!$J$2 = D{},TRUE,""))'.format(i))      #modifying values of cells in a column with sequences of function according to it's location in the column
        for_powerBI.update('AU{}'.format(i) , '=IF(\'Sales Dashbaord\'!$J$3 = "" , TRUE , IF(\'Sales Dashbaord\'!$J$3 = Z{},TRUE,""))'.format(i))     #modifying values of cells in a column with sequences of function according to it's location in the column
        for_powerBI.update('AV{}'.format(i) , '=IF(\'Sales Dashbaord\'!$J$4 = "" , TRUE , IF(\'Sales Dashbaord\'!$J$4 = H{},TRUE,""))'.format(i))     #modifying values of cells in a column with sequences of function according to it's location in the column
        for_powerBI.update('AW{}'.format(i) , '=IF(\'Sales Dashbaord\'!$J$5 = "" , TRUE , IF(\'Sales Dashbaord\'!$J$5 = J{},TRUE,""))'.format(i))     #modifying values of cells in a column with sequences of function according to it's location in the column 
        
    for i in range(3,4):
        for_powerBI.update('BE{}'.format(i) , '=IFERROR(BB{}/BD{},)'.format(i,i))
        for_powerBI.update('BF{}'.format(i) , '=IF(BE{}*0.2 >=0.2,0.2,BE{}*0.2)'.format(i,i))

        for_powerBI.update('BI{}'.format(i) , '=IF(BD{} <>"", (BH{} +BB{})/BD{},"")'.format(i,i,i,i))
        for_powerBI.update('BG{}'.format(i) , '=IF(BD$1 = "" , TRUE , IF(BD$1 = BJ{},TRUE,""))'.format(i))
        for_powerBI.update('BH{}'.format(i) , '=IFERROR(VLOOKUP(BA{},Sheet25!$A$127:$F$178,6))'.format(i))
        
        
    for i in range(2,6):
        for_powerBI.update('AY{}'.format(i) , '=\'Sales Dashbaord\'!J{}'.format(i))
       

    for_powerBI.update('AZ2'.format(i) , '=query(B:AM,"SELECT C,AM, SUM (AA) WHERE C is not null And AM >= date """& text(AZ1,"yyyy-mm-dd")&""" GROUP BY C , AM ")')


    
    for_powerBI.update('BC3'.format(i) , '=TRANSPOSE(QUERY(IMPORTRANGE("1GUmTSjlh4niWZqW6mq0ODYhF_kAzY_WnQFbUBOvkU3g","Employees!B:BZ")," SELECT Col20,Col21 , Col22 ,Col23, Col24, Col25,Col26,Col27,Col28 ,Col29 ,Col30, Col31,Col32 ,Col33 , Col34,Col35 ,Col36 ,Col37 , Col38 ,Col39 , Col40 ,Col41,Col42,Col43,Col44,Col45,Col46,Col47,Col48,Col49,Col50,Col51,Col52,Col53,Col54 WHERE Col1 = \'{}\' ",1))'.format(emp_name))


        

    for_powerBI.update('BD2' , 'Target')      #renaming of columns header
    for_powerBI.update('BE2' , 'Achieved')    #renaming of columns header
    for_powerBI.update('BF2' , 'Percentage')  #renaming of columns header
    

    for_powerBI.update('AT1' , 'Country')     #renaming of columns header
    for_powerBI.update('AU1' , 'Source')      #renaming of columns header
    for_powerBI.update('AV1' , 'Service')     #renaming of columns header
    for_powerBI.update('AW1' , 'Sector')      #renaming of columns header
    


    for_powerBI.update('AX2' , 'Country')     #renaming of columns cells for data manipulation
    for_powerBI.update('AX3' , 'Source')      #renaming of columns cells for data manipulation
    for_powerBI.update('AX4' , 'Service')     #renaming of columns cells for data manipulation
    for_powerBI.update('AX5' , 'Sector')      #renaming of columns cells for data manipulation
    



    target.update('B9' , '=QUERY(\'For power bi\'!BA:BI , "SELECT BA , BD , BB, BE , BI WHERE BA IS NOT NULL AND BA >= date """& text(D2,"yyyy-mm-dd" )&""" AND BA <= date """& text(D5,"yyyy-mm-dd" )&""" LABEL BI\'Acheieved if made 20% Discount \' ,BB \'Total Paid\' ")')
    time.sleep(10)         # to overcome the problem of requests Quota
    



