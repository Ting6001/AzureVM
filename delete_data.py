try:
    import pyodbc
    print('Find module: pyodbc')
except ImportError:
    from pip._internal import main as pip
    print('Start to install module: pyodbc')
    pip(['install','pyodbc'])
    import pyodbc
 

server = 'sqlserver-mia.database.windows.net' 
database = 'DB-mia' 
username = 'admia' 
password = 'Mia01@wistron' 
driver = '{SQL Server Native Client 11.0}'
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