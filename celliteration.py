#Contactenates two excel files or sheets and create a new sheet
import pandas as pds
import numpy as np
import openpyxl

# To open the workbook
# workbook object is created
wb_obj = openpyxl.load_workbook("Accounts.xlsx")
 
# Get workbook active sheet object
# from the active attribute
sheet_obj = wb_obj.active
 
# Cell objects also have a row, column,
# and coordinate attributes that provide
# location information for the cell.
 
# Note: The first row or
# column integer is 1, not 0.
 
# Cell object is created by using
# sheet object's cell() method.
cell_obj = sheet_obj.cell(row = 1, column = 1)
 
# Print value of cell object
# using the value attribute
print(cell_obj.value)

# for i in range(accounts_file.shape[0]): #iterate over rows
#     for j in range(accounts_file.shape[1]): #iterate over columns
#         value = accounts_file.at[i, j] #get cell value
#         print(value, end="\t")
#     print()

# dictionary of lists
dict = {'name':["aparna", "pankaj", "sudhir", "Geeku"],
        'degree': ["MBA", "BCA", "M.Tech", "MBA"],
        'score':[90, 40, 80, 98]}
 
# creating a dataframe from a dictionary
df = pds.DataFrame(dict)
#print("\nData Frame","\n",df,"\n")
 
# # iterating over rows using iterrows() function
# for i, j in df.iterrows():
#     print("\n","i=",i,"j=",j,"\n")
#     print(i, j)
#     print
#     print()

# df = pds.DataFrame(
# 	[[1, 2, 3],
# 	[4, 5, 6],
# 	[7, 8, 9],
# 	[10, 11, 12]])

# for i in range(df.shape[0]): #iterate over rows
#     for j in range(df.shape[1]): #iterate over columns
#         value = df.at[i, j] #get cell value
#         print(value, end="\t")
#     print()

