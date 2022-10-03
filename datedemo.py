import pandas as pd

# https://www.educative.io/answers/what-is-the-pandastodatetime-method

# arg: an integer, float, string, list, or dict object to convert into a DateTime object.
# dayfirst: set it to true if the input contains the day first.
# yearfirst: set it to true if the input contains the year first.
# utc: returns the UTC DatetimeIndex if True.
# format: specifies the position of the day, month, and y
# input in mm.dd.yyyy format
date = ['01.02.2019']
# output in yyyy-mm-dd format
print("OUTPUT BEGINS HERE\n")
print("\nOutput in yyyy-mm-dd format\n")
print(pd.to_datetime(date))
print("\n")

# date (mm.dd.yyyy) and time (H:MM:SS)
date = ['01.02.2019 1:30:00 PM']
# output in yyyy-mm-dd HH:MM:SS
print(pd.to_datetime(date))
print("\n")

date = '2019-07-31 12:00:00-UTC'
print(pd.to_datetime(date, format = '%Y-%m-%d %H:%M:%S-%Z'))
print("\n") 

# Using pandas.to_datetime() 
# with dates in dd-mm-yyyy and yy-mm-dd format
# pandas interprets this date to be in m-d-yyyy format
print(pd.to_datetime('8-2-2019'))
print("\n")
# if the specified date contains the day first, then 
# that has to be specified.
# output in yyyy-mm-dd format.
print(pd.to_datetime('8-2-2019', dayfirst = True))
print("\n")
# if the specified date contains the year first, then 
# that has to be specified.
# output in yyyy-mm-dd format.
print(pd.to_datetime('10-2-8', yearfirst = True))
print("\n")

# Using pandas.to_datetime() to obtain a timezone-aware timestamp
date = '2019-01-01T15:00:00+0100'
print(pd.to_datetime(date, utc = True))