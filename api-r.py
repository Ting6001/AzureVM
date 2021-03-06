from flask import Flask, request 
from flask_restful import Resource, Api,reqparse
import random
import pandas as pd
import json
import time
from datetime import datetime
import sys, traceback
import pyodbc 
import socket

# R
import rpy2.robjects as robjects
from rpy2.robjects import pandas2ri
from rpy2.robjects.conversion import localconverter



app = Flask(__name__)
app.config['DEBUG'] = True
api = Api(app)
root = './' # local
path_r = root + 'powerApp_func_multi_v7.R'                

class WorkRate (Resource):
    parser = reqparse.RequestParser()
    parser = parser.add_argument('division', type=str, required=True,  help='div is required')
    parser = parser.add_argument('user', type=str, required=True,  help='user is required')
    parser = parser.add_argument('data_Prj', type=dict, location='json', action="append", help='data Project info is required')
    parser = parser.add_argument('data_HC', type=dict, location='json', action="append", help='data Hc info is required')

    def post(self): # create
        try:
            start = time.process_time()
            arg = self.parser.parse_args()
            # print(arg)     
            set_dep = set()
            set_sub = set()
            df_prj = pd.DataFrame()
            df_HC = pd.DataFrame()

            div = arg['division']
            if arg['data_Prj']:
                df_prj = pd.DataFrame(arg['data_Prj'])
                # pandas DataFrame去除na，有na值的話，直接轉換會error
                df_prj.fillna("", inplace=True)
                df_prj = df_prj.astype({'project_code':str, 'project_code_old':str})
            if arg['data_HC']:   
                df_HC = pd.DataFrame(arg['data_HC'])
                # pandas DataFrame去除na，有na值的話，直接轉換會error
                df_HC = df_HC.drop(['save_time', 'user_id'], axis=1)
                df_HC = df_HC.astype({'div':str, 'deptid':str, 'project_code':str})
            try:
                # print('div:', div)
                #################  先把code放進來 debug 用 ####################################################
                robjects.r.source(path_r)
                function_r = robjects.globalenv['hr_cal_multi']

                # === 將 Pandas.df 轉換成 R df ===
                # print('==== Transfer Pandas.df to R.df ====')
                with localconverter(robjects.default_converter + pandas2ri.converter):
                    df_prj_r = robjects.conversion.py2rpy(df_prj)
                    df_HC_r = robjects.conversion.py2rpy(df_HC)
        
                # === 呼叫 R Function， return R df ===
                # print('==== Before R Functioin ====')
                time_R_start = time.process_time()

                df_result_r = function_r(df_prj_r, df_HC_r, div)

                time_taken_R = round(time.process_time() - time_R_start,3)
                # print('==== Transfer R.df to Pandas.df ====')
                # === 將 R df 轉換成  Pandas.df ===
                with localconverter(robjects.default_converter + pandas2ri.converter):
                    df_result = robjects.conversion.rpy2py(df_result_r)       
                # print('======== Function return ==========')
                dic_result = {}
                if isinstance(df_result, pd.DataFrame):
                    df_result.fillna(0.0, inplace=True)
                    dic_result = df_result.to_dict('records')
                    # print(df_result)
                time_taken = round(time.process_time() - start,3)
                # print('Totally Take:', time_taken, 's,  R fuction:', time_taken_R, 's')
                # print('====================================================================================')
                return dic_result
            except Exception as e:
                self.save_log(e, arg)
        except Exception as e:
            self.save_log(e, None)


    def save_log(self, e, arg):
        query = 'INSERT INTO LogInfo (user_input, error_msg, save_time, user_id) VALUES (?,?,?,?)'
        user_id, user_input = '', ''
        if arg != None:
            user_id = arg['user']
            user_input = str(arg)
        save_time = datetime.now()
        err_class = e.__class__.__name__ #取得錯誤類型
        detail = e.args[0] #取得詳細內容
        cl, exc, tb = sys.exc_info() #取得Call Stack
        lst_call_stack = traceback.extract_tb(tb)
        err_msg = '【{}】 {}'.format(err_class, detail)
        for CallStack in lst_call_stack:
            fileName = CallStack[0] #取得發生的檔案名稱
            lineNum = CallStack[1] #取得發生的行號
            funcName = CallStack[2] #取得發生的函數名稱
            err_msg += "\nFile \"{}\", line {}, in {}".format(fileName, lineNum, funcName)

        # print(err_msg)
        data = (user_input, err_msg, save_time, user_id)

        conn, cursor = self.connect_db()
        cursor.execute(query, data)
        conn.commit()
        self.close_db(conn, cursor)

        ##############################################################################################
        # 假資料區
    def connect_db(self):
        server = 'sqlserver-mia.database.windows.net' 
        database = 'DB-mia' 
        username = 'admia' 
        password = 'Mia01@wistron' 
        # driver = '{ODBC Driver 17 for SQL Server}'
        # driver = '{SQL Server Native Client 11.0}'
        drivers = [item for item in pyodbc.drivers()] # ['SQL Server', 'SQL Server Native Client 11.0', 'ODBC Driver 11 for SQL Server']
        # print('drivers:', drivers)
        driver = drivers[-1]
        conn = pyodbc.connect('DRIVER='+driver+';SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
        cursor = conn.cursor()
        # print('Connect to DB')
        return conn, cursor
    def close_db(self, conn, cursor):
        # print('Close DB Connection')
        cursor.close()
        conn.close()
        
class HelloWorld(Resource):
    def get(self, name="world"):
        return {'hello': name}

api.add_resource(HelloWorld, '/','/hello/<string:name>')

# api.add_resource(User, '/user/<string:name>')
api.add_resource(WorkRate, '/workrate/')

if __name__ == '__main__':
# 檢查5000 port是否使用中，未使用才run
    cur_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    port = 5000
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    result = sock.connect_ex(('127.0.0.1',port)) # 52.163.121.219
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)

    msg = '【'+ cur_time + '】'
    # global root
    # global path_r
    path_log = root + '/api_local.log'
    if local_ip == '10.0.0.6': # VM
        root = '/home/mia06/AzureVM/'
        path_r = root + "powerApp_func_multi_v7.R"
        path_log = '/home/mia06/api.log'

    if result == 0:
        msg += (local_ip + ", Port {} is already open\n".format(port))
        with open(path_log, 'a') as f:
            f.write(msg)
    else:
        msg += (local_ip + ", Port {} is not open\n".format(port))
        with open(path_log, 'a') as f:
            f.write(msg)
        app.run(host='0.0.0.0', debug=True, use_reloader=False) 

    print(msg)
    sock.close()
    
    
    
