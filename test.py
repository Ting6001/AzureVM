import pyodbc 
import pandas as pd
from datetime import datetime
import sys
import traceback
import time
# Some other example server values are
# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = 'sqlserver-mia.database.windows.net' 
database = 'DB-mia' 
username = 'admia' 
password = 'Mia01@wistron' 
driver = '{ODBC Driver 17 for SQL Server}'
# driver = '{SQL Server Native Client 11.0}'
drivers = [item for item in pyodbc.drivers()] # ['SQL Server', 'SQL Server Native Client 11.0', 'ODBC Driver 11 for SQL Server']
driver = drivers[-1]
print('drivers:', drivers)
conn = pyodbc.connect('DRIVER='+driver+';SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = conn.cursor()


lst_1 = []
lst_2 = []
lst_diff = []
for i in range(10):
    start = time.process_time()
    query = 'select * from UtilizationRateInfo'
    df = pd.read_sql(query, conn)
    time_taken_1 = round(time.process_time() - start, 3)

    print(i, '【1】Totally Take:', time_taken_1, 's')
    # print(df)

    start_2 = time.process_time()
    query = 'select * from UtilizationRateInfo2'
    df = pd.read_sql(query, conn)
    time_taken_2 = round(time.process_time() - start_2, 3)

    time_diff= round(time_taken_1 - time_taken_2, 3)
    print(i, '【2】Totally Take:', time_taken_2, 's')
    print(i, 'Time Difference：', time_diff, 's')
    lst_1.append(time_taken_1)
    lst_2.append(time_taken_2)
    lst_diff.append(time_diff)
print('================')
print('average 1:', round(sum(lst_1)/len(lst_1), 3), 's', lst_1)
print('average 2:', round(sum(lst_2)/len(lst_2), 3), 's', lst_2)
print('average diff:', round(sum(lst_diff)/len(lst_diff), 3), 's', lst_diff)
# average 1: 0.909 s [1.062, 0.922, 0.672, 0.984, 0.688, 0.875, 0.75, 1.109, 1.281, 0.75]
# average 2: 0.459 s [0.531, 0.578, 0.375, 0.359, 0.422, 0.406, 0.5, 0.391, 0.641, 0.391]
# average diff: 0.45 s [0.531, 0.344, 0.297, 0.625, 0.266, 0.469, 0.25, 0.718, 0.64, 0.359]
# print(df)

cursor.close()
conn.close()
