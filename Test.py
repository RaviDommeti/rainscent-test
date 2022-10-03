#Contactenates two excel files or sheets and create a new sheet
import string
import pandas as pds
import numpy as np
import xlwings as xw
import openpyxl as op
# from openpyxl import Workbook
# from openpyxl.styles import Color, PatternFill, Font, Border
# from openpyxl.styles import colors
# from openpyxl.cell import Cell

file1 =('Vendors Consolidated.xlsx')
file2 =('Accounts.xlsx')
debug_flag = 0;#Custom int variable. For testing, to print statements with PRINT ,set it to 1

#Concatenate from two sheets / files
vendors_file = pds.read_excel(file1)
accounts_file = pds.read_excel(file2)
joinedData = pds.concat([vendors_file, accounts_file])

print("\n Size of Vendors File: ",vendors_file.shape," with Rows: ",vendors_file.shape[0]," Columns: ",vendors_file.shape[1])
# Find rows and columns with null values
result_error = np.where(pds.isnull(vendors_file))

if(debug_flag == 1):
    # result_error[0] has rows with null values while result_error[1] has columns with null value starting from 0
    print("\n Error in Rows: ",result_error[0],"\nError in Columns:  ",result_error[1])


rows_error = result_error[0] # Storing Error rows
cols_error = result_error[1] # Storing Error cols
len_rows_errors = len(result_error[0])
print("\nSize of Rows Errors:  ",len_rows_errors) # Number of records with errors
#Retrieve locations of error values
if(debug_flag == 1):
    print("\nError Value Locations ")
    # for(i = 0; i < len_rows_errors; i++)
    for i in range(len_rows_errors):
        print("\n[",rows_error[i],"] [",cols_error[i],"]")



#Storing column names of Vendors consolidated
cols_list = list(vendors_file.columns)
if(debug_flag == 1):
    print("\n Vendors Consolidated Columns: ",cols_list,"\n")


#Finding number of columsn in Vendors consolidated starts from 1
total_cols_vendors = len(cols_list)
if(debug_flag == 1):
    print("\n Number of Columns in Vendors Consolidated",total_cols_vendors,"\n")

# Duplicating vendors file
dup_vendors_file = vendors_file

error_string = []
# for(i = 0; i < len_rows_errors; i++)
for i in  range(vendors_file.shape[0]):
    error_string.append("")
    
if(debug_flag == 1):
    print("\nError String: ",error_string)

#print("\nError String: ",error_string[i]))
for i in range(len_rows_errors):
    rows_num = rows_error[i]
    error_cols_name = cols_list[cols_error[i]]
    error_string[rows_num] = error_string[rows_num]+" Missing "+error_cols_name

if(debug_flag==1):
    print("\nError String: ",error_string)
    
for i in  range(vendors_file.shape[0]):
    dup_vendors_file['Errors'] = error_string

if(debug_flag==1):
    print("\nVendors Error Col is ",dup_vendors_file['Errors'])


missing_value_file = dup_vendors_file
missing_value_file.drop(missing_value_file[missing_value_file['Errors'] == ""].index, inplace = True)
if(debug_flag == 1):
    print("\n",missing_value_file)

missing_value_file.to_excel("Vendor Missing values.xlsx")

#Please note vendors_file now contains only rows which have values in Error column. vendors_file cannot be used for further processing
#Create a new reference to the Vendors Consolidated file to process the original file
vendors_file = pds.read_excel("Vendors Consolidated.xlsx")
#Removing rows with null values in Vendors Consolidated
vendors_file = vendors_file.dropna(how='any')
#Converting DC NO to String from int. Failure to do this is resulting in deletion of values
# when used with str.strip()
vendors_file = vendors_file.astype({"DC NO": str})

# Removing spaces before and after values
vendors_file['DC NO'] = vendors_file['DC NO'].str.strip()
vendors_file['TRUCK NO'] = vendors_file['TRUCK NO'].str.strip()
# Removing spaces in middle of values
vendors_file['DC NO'] = vendors_file['DC NO'].str.replace(" ","")
vendors_file['TRUCK NO'] = vendors_file['TRUCK NO'].str.replace(" ","")


print("\n","Vendors File Cleaned ",vendors_file.shape,"\n")
vendors_file.to_excel("Vendors Cleaned.xlsx")

if(debug_flag == 0):
    print("\nAccounts File ",accounts_file,"\n")

#Converting DC NO to String from int. Failure to do this is resulting in deletion of values
# when used with str.strip()
accounts_file = accounts_file.astype({"DC NO": str})
# Removing spaces before and after values
accounts_file['DC NO'] = accounts_file['DC NO'].str.strip()
accounts_file['TRUCK NO'] = accounts_file['TRUCK NO'].str.strip()
# Removing spaces in middle of values
accounts_file['DC NO'] = accounts_file['DC NO'].str.replace(" ","")
accounts_file['TRUCK NO'] = accounts_file['TRUCK NO'].str.replace(" ","")
if(debug_flag == 1):
    print("\nAccounts File ",accounts_file,"\n")
    print("\nData types of Columns ",accounts_file.dtypes)
    print("\n Data Info of Accounts File\n",accounts_file.info())

#Merge Tables
joinedData = vendors_file.merge(accounts_file, how="inner", on=['DC NO','TRUCK NO','DATE'],suffixes=('_Vendors', '_Accounts'))
if(debug_flag == 1):
    print("\n Joined Data",joinedData)
#joinedData.to_excel("Accounts Summary.xlsx")



joinedData['REC Difference'] = np.where((joinedData['Received Qty_Vendors'] != joinedData['Received Qty_Accounts']), joinedData['Received Qty_Vendors'] - joinedData['Received Qty_Accounts'], "All OK")
joinedData['ACC Difference'] = np.where((joinedData['Accepted Qty_Vendors'] != joinedData['Accepted Qty_Accounts']), joinedData['Accepted Qty_Vendors'] - joinedData['Accepted Qty_Accounts'], "All OK")
joinedData = joinedData[['DATE','DC NO','TRUCK NO','Accepted Qty_Vendors','Accepted Qty_Accounts','ACC Difference','Received Qty_Vendors','Received Qty_Accounts','REC Difference']]
#joinedData = joinedData.loc[(joinedData["ACC Difference"] == "All OK" and joinedData["ACC Difference"] == "All OK") ]
#joinedData.drop(joinedData[joinedData['ACC Difference'] == "All OK" and joinedData['REC Difference'] == "All OK"].index)
joinedData.to_excel("Account Errors.xlsx")

finalData = vendors_file.merge(accounts_file, how="inner", on=['DC NO','TRUCK NO','DATE','Received Qty','Accepted Qty'])
finalData = finalData.replace(np.nan,0)
print("\nAccounts Summary  ",finalData,"\n")

finalData['AMOUNT'] = finalData['Accepted Qty'] * finalData['PRICE']
finalData['TAX'] = (3/100) * finalData['AMOUNT']
finalData['GRAND TOTAL'] = finalData['AMOUNT'] + finalData['TAX']
finalData['BALANCE'] = finalData['GRAND TOTAL'] - (finalData['T PAID'] + finalData['FRT'])


finalData.to_excel("Account Summary.xlsx")
if(debug_flag == 1):
    print("\nAccounts Summary  ",finalData,"\n")
    print("\nAccounts Summary Info\n ",finalData.info(),"\n")

