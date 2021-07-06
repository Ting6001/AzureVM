import pyodbc 
import pandas as pd
from datetime import datetime
import sys
import traceback
# Some other example server values are
# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = 'sqlserver-mia.database.windows.net' 
database = 'DB-mia' 
username = 'admia' 
password = 'Mia01@wistron' 
# driver = '{ODBC Driver 17 for SQL Server}'
# driver = '{SQL Server Native Client 11.0}'
drivers = [item for item in pyodbc.drivers()] # ['SQL Server', 'SQL Server Native Client 11.0', 'ODBC Driver 11 for SQL Server']
driver = drivers[-1]
print('drivers:', drivers)
cnxn = pyodbc.connect('DRIVER='+driver+';SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

#Sample select query
# cursor.execute("SELECT @@version;") 
# row = cursor.fetchone() 
# while row:
#     print('@@') 
#     print(row[0])
#     row = cursor.fetchone()

# try:
#     a = 1 / 0
# except Exception as e:
#     print(type(e), e)
#     error_msg = str(e)
#     save_time = datetime.now()
#     user_id = '001'
#     user_input = '[{div:"230000"}]'

#     query = 'INSERT INTO LogInfo (user_input, error_msg, save_time, user_id) VALUES (?,?,?,?)'

#     err_class = e.__class__.__name__ #取得錯誤類型
#     detail = e.args[0] #取得詳細內容
#     cl, exc, tb = sys.exc_info() #取得Call Stack
#     lastCallStack = traceback.extract_tb(tb)[-1] #取得Call Stack的最後一筆資料

#     fileName = lastCallStack[0] #取得發生的檔案名稱
#     lineNum = lastCallStack[1] #取得發生的行號
#     funcName = lastCallStack[2] #取得發生的函數名稱
#     err_msg = "File \"{}\", line {}, in {}: [{}] {}".format(fileName, lineNum, funcName, err_class, detail)
#     print(err_msg)
#     data = (user_input, err_msg, save_time, user_id)
#     cursor.execute(query, data)
#     cnxn.commit()

# row = cursor.fetchone()
# cursor.execute(query,'000')
query = 'select * from LogInfo'
df = pd.read_sql(query, cnxn)
print(df)


