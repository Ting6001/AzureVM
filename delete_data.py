import os
def install_function():
    os.system('sudo curl https://packages.microsoft.com/keys/microsoft.asc | apt-key add -')
    os.system('sudo curl https://packages.microsoft.com/config/ubuntu/18.04/prod.list > /etc/apt/sources.list.d/mssql-release.list')
    os.system('sudo apt-get update')
    os.system('sudo ACCEPT_EULA=Y apt-get install -y msodbcsql17')

install_function()

try:
    import pyodbc
    print('Find module: pyodbc')
except ImportError:
    from pip._internal import main as pip
    print('Start to install module: pyodbc')
    pip(['install','pyodbc'])
    import pyodbc
 

drivers = [item for item in pyodbc.drivers()]
print(drivers)
driver = drivers[-1]
print("driver:{}".format(driver))

server = 'sqlserver-mia.database.windows.net' 
database = 'DB-mia' 
username = 'admia' 
password = 'Mia01@wistron' 
# driver = '{SQL Server Native Client 11.0}'
# driver = '{ODBC Driver 17 for SQL Server}'
cnxn = pyodbc.connect('DRIVER='+driver+';SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

table_name = 'ProjInfo_test_pk3'
query = 'delete from {table_name} where user_id =?'.format(table_name=table_name)
try:
    user_id = '000'
    cursor.execute(query, user_id)
    cnxn.commit()
    print('Finish delete')
except Exception as e:
    print(e)