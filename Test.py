#Contactenates two excel files or sheets and create a new sheet
import string
from tempfile import tempdir
import pandas as pds
import numpy as np
import xlwings as xw
import openpyxl as op
# from openpyxl import Workbook
# from openpyxl.styles import Color, PatternFill, Font, Border
# from openpyxl.styles import colors
# from openpyxl.cell import 
from tkinter import *
from tkinter import messagebox
from tkinter.ttk import *
from tkinter import ttk
import time
import webbrowser
import datetime

# https://likegeeks.com/python-gui-examples-tkinter-tutorial/
feedback_str = []
version_no = "T 1.0"

#Defintion of Log function
def writeToLog(message):
    f = open("log.txt", "a")
    # ct stores current time
    ct = datetime.datetime.now()
    print("at ", ct,file=f)
    print(message, file=f)
    f.close()

#Defintion of  def displayMessage function 
def displayMessage(title,message):
    messagebox.showinfo(title,message)


#*************************************Beginning of Function updateProgress()******************************************
def updateProgress(message_update,bar_value):
    
    feedback_str.append(message_update)
    bar['value'] = bar_value
    temp_str =""
    for i in range(len(feedback_str)):
        temp_str = temp_str + feedback_str[i]
    run_feedback.configure(text=temp_str)

#*************************************Beginning of Function runTool()******************************************
def runTool():
    # get the start time of program
    temp_str = "-------------------Version "+version_no+" (TRIAL VERSION)-------------------------"
    writeToLog(temp_str)
    temp_str = "-------------------RUN TOOL clicked-------------------------"
    writeToLog(temp_str)
    start_time = time.time()
    file1 =('Vendors Consolidated.xlsx')
    file2 =('Accounts.xlsx')
    debug_flag = 0;#Custom int variable. For testing, to print statements with PRINT ,set it to 1

    try:
        #Concatenate from two sheets / files
        vendors_file = pds.read_excel(file1)
        accounts_file = pds.read_excel(file2)
        joinedData = pds.concat([vendors_file, accounts_file])
        temp_str = "Input files read successfully"
        writeToLog(temp_str)
        updateProgress("\nReading both input files completed",10)
    except Exception as err_msg:
        print("\nError reading files: "+str(err_msg))
        displayMessage("Error",str(err_msg))
        temp_str = "Error "+str(err_msg)
        writeToLog(temp_str)

    if(debug_flag == 1):
        print("\n Size of Vendors File: ",vendors_file.shape," with Rows: ",vendors_file.shape[0]," Columns: ",vendors_file.shape[1])
    
    # Find rows and columns with null values
    result_error = np.where(pds.isnull(vendors_file))
    #feedback_str.append("\nRows with empty or blank values identified")
    temp_str = "\n Error in Rows: "+ str(result_error[0])+"\nError in Columns:  "+str(result_error[1])
    writeToLog(temp_str)

    if(debug_flag == 1):
        # result_error[0] has rows with null values while result_error[1] has columns with null value starting from 0
        print("\n Error in Rows: ",result_error[0],"\nError in Columns:  ",result_error[1])


    rows_error = result_error[0] # Storing Error rows
    cols_error = result_error[1] # Storing Error cols
    len_rows_errors = len(result_error[0])
    temp_str = "Size of Rows Errors:  "+str(len_rows_errors)
    writeToLog(temp_str)
    #Retrieve locations of error values
    if(debug_flag == 1):
        print("\nSize of Rows Errors:  ",len_rows_errors) # Number of records with errors
        print("\nError Value Locations ")
        # for(i = 0; i < len_rows_errors; i++)
        for i in range(len_rows_errors):
            print("\n[",rows_error[i],"] [",cols_error[i],"]")
            temp_str = "["+str(rows_error[i])+"] ["+str(cols_error[i])+"]"
            writeToLog(temp_str)



    #Storing column names of Vendors consolidated
    cols_list = list(vendors_file.columns)
    if(debug_flag == 1):
        print("\n Vendors Consolidated Columns: ",cols_list,"\n")


    #Finding number of columns in Vendors consolidated starts from 1
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

    missing_value_file.to_excel("Vendor Missing Values.xlsx")
    temp_str = "Vendor Missing Values.xlsx created"
    writeToLog(temp_str)
    updateProgress("\nFile with name Vendor Missing Values.xlsx created",30)

    try:
        #Please note vendors_file now contains only rows which have values in Error column. vendors_file cannot be used for further processing
        #Create a new reference to the Vendors Consolidated file to process the original file
        vendors_file = pds.read_excel("Vendors Consolidated.xlsx")
        
    except Exception as err_msg:
        print("\nError reading files: "+str(err_msg))
        displayMessage("Error",str(err_msg))
        temp_str = "Exception raised: "+str(err_msg)
        writeToLog(temp_str)
    
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

    if(debug_flag == 1):
        print("\n","Vendors File Cleaned ",vendors_file.shape,"\n")
    
    temp_str = "Vendor files cleaned"
    writeToLog(temp_str)
    updateProgress("\nSpaces removed in DC NO, TRUCK NO",50)
    #vendors_file.to_excel("Vendors Cleaned.xlsx")

    if(debug_flag == 1):
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
    updateProgress("\nMerging data in files...",50)



    joinedData['REC Difference'] = np.where((joinedData['Received Qty_Vendors'] != joinedData['Received Qty_Accounts']), joinedData['Received Qty_Vendors'] - joinedData['Received Qty_Accounts'], "All OK")
    joinedData['ACC Difference'] = np.where((joinedData['Accepted Qty_Vendors'] != joinedData['Accepted Qty_Accounts']), joinedData['Accepted Qty_Vendors'] - joinedData['Accepted Qty_Accounts'], "All OK")
    joinedData = joinedData[['DATE','DC NO','TRUCK NO','Accepted Qty_Vendors','Accepted Qty_Accounts','ACC Difference','Received Qty_Vendors','Received Qty_Accounts','REC Difference']]
    #joinedData = joinedData.loc[(joinedData["ACC Difference"] == "All OK" and joinedData["ACC Difference"] == "All OK") ]
    #joinedData.drop(joinedData[joinedData['ACC Difference'] == "All OK" and joinedData['REC Difference'] == "All OK"].index)
    joinedData.to_excel("Account Errors.xlsx")
    temp_str = "Account Errors.xlsx created"
    writeToLog(temp_str)
    updateProgress("\nFile created with name Account Errors.xlsx",70)

    try:
        finalData = vendors_file.merge(accounts_file, how="inner", on=['DC NO','TRUCK NO','DATE','Received Qty','Accepted Qty'])
        finalData = finalData.replace(np.nan,0)
        temp_str = "Accounts vs Vendors merge successful"
        writeToLog(temp_str)
        updateProgress("\nReading both input files completed",10)
    except Exception as err_msg:
        print("\nError reading files: "+str(err_msg))
        displayMessage("Error merging: ",str(err_msg))
        temp_str = "Error "+str(err_msg)
        writeToLog(temp_str)

    if(debug_flag == 1):
        print("\nAccounts Summary  ",finalData,"\n")

    updateProgress("\nCalculating values....",90)

    finalData['AMOUNT'] = finalData['Accepted Qty'] * finalData['PRICE']
    finalData['TAX'] = (3/100) * finalData['AMOUNT']
    finalData['GRAND TOTAL'] = finalData['AMOUNT'] + finalData['TAX']
    finalData['BALANCE'] = finalData['GRAND TOTAL'] - (finalData['T PAID'] + finalData['FRT'])

    finalData.to_excel("Account Summary.xlsx")
    temp_str = "Account Summary.xlsx created"
    writeToLog(temp_str) 
    updateProgress("\nFile with name Account Summary.xlsx is created",100)
    
    if(debug_flag == 1):
        print("\nAccounts Summary  ",finalData,"\n")
        print("\nAccounts Summary Info\n ",finalData.info(),"\n")
    #End time of program
    end_time = time.time()
    time_elapsed = (end_time - start_time)/60
    time_elapsed = round(time_elapsed,2)

    updateProgress("\nProgram run successful",100)
    temp_str = "-------------------Program completed-------------------------"
    writeToLog(temp_str)


    



#*************************************End of Function runTool()******************************************


#*************************************GUI Code******************************************

window = Tk()
window.title("ExcelBuddy")
window.geometry('500x500') # Width X Height

lbl = Label(window, text="Welcome to ExcelBuddy", font=("Arial Bold", 25,))
lbl.grid(column=0, row=0)

text_str = "Version: "+version_no+"\n"
lbl_version = Label(window, text=text_str, font=("Arial", 12))
lbl_version.grid(column=0, row=1)

btn = Button(window, text="RUN TOOL", command=runTool)
btn.grid(column=0, row=2)

# style = ttk.Style()
# style.theme_use('default')
# style.configure("black.Horizontal.TProgressbar", background='black')

# bar = Progressbar(window, length=200, style='black.Horizontal.TProgressbar')

bar = Progressbar(window, length=200)

bar.grid(column=0, row=3)

run_feedback = Label(window, text=" ", font=("Arial Bold", 15,))
run_feedback.grid(column=0, row=4,padx=15)

lbl_company_name = Label(window, text="Tool developed by Rainscent Tech Pvt. Ltd. ", font=("Arial Bold", 15,))
lbl_company_name.grid(column=0, row=5)

#Defintion of Website function
def toWebsite(url):
    webbrowser.open_new_tab(url)

lbl_website = Label(window, text="https://www.rain-scent.com/", font=("Arial Bold", 15), foreground= "blue", cursor="hand2")
url= "https://www.rain-scent.com/"
lbl_website.bind("<Button-1>", lambda e:toWebsite(url))
lbl_website.grid(column=0, row=6)

#displayMessage("Note","This version is only for the purpose of testing only")

window.mainloop()

#*************************************End of GUI Code******************************************

# main = Tk()
# ourMessage ='This is our Message'
# messageVar = Message(main, text = ourMessage)
# messageVar.config(bg='lightgreen')
# messageVar.pack( )
# main.mainloop( )

