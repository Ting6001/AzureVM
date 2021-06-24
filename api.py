from flask import Flask, request 
from flask_restful import Resource, Api,reqparse
# from werkzeug.contrib.fixers import ProxyFix # new
import random
import pandas as pd
import json
# R
import rpy2.robjects as robjects
from rpy2.robjects import pandas2ri
from rpy2.robjects.conversion import localconverter

app = Flask(__name__)
app.config['DEBUG'] = True
api = Api(app)

class WorkRate (Resource):
    parser = reqparse.RequestParser()
    # parser.add_argument('email', required=True, help='Email is required')
    # parser = parser.add_argument('lst', type=str, location='json', action="append")
    parser = parser.add_argument('div', type=str, required=True,  help='div is required')
    parser = parser.add_argument('data_Prj', type=dict, location='json', action="append", help='data Project info is required')
    parser = parser.add_argument('data_HC', type=dict, location='json', action="append", help='data Hc info is required')

    def post(self): # create
        arg = self.parser.parse_args()
        print(arg)
        # print(arg)     
        set_dep = set()
        set_sub = set()
        df_prj = pd.DataFrame()
        df_HC = pd.DataFrame()
        div = arg['div']
        if arg['data_Prj']:
            df_prj = pd.DataFrame(arg['data_Prj'])
            for item in arg['data_Prj']:
                print(item)
        if arg['data_HC']:
            df_HC = pd.DataFrame(arg['data_HC'])
            for item in arg['data_HC']:
                if item['HC'] > 0:
                    set_dep.add(item['deptid'])
                    set_sub.add(item['sub_job_family'])
        
        print('div:', div)
        print('dep:', set_dep)
        print('sub:', set_sub)
        # 假資料區
        lst_return = []    
        dic_info = {'DEP':set_dep,
                    'SUB':set_sub
                    }
        div_data = {'230000':{'DEP':["230110","230120","230320","230350","230I10","230I30","230R30",],
                        'SUB':["RH", "ME"]},
                    '23D000':{'DEP':["23D000","23D100","23D200","23D300","23D500","23D600"],
                        'SUB':["AR", "HW", "OP", "RF", "RH"]},
                    '23M000':{'DEP':["23M000","23M100","23M200","23M300","23M500","23M600"],
                        'SUB':["ME"]},
                    '23R000':{'DEP':["23R000","23R200","23R300"],
                        'SUB':["HW", "ME", "RH"]}
                    }
                    
        for key, lst in div_data[div].items():
            for item in lst:
                print(div, key, item)
                dic_tmp = {}
                dic_tmp['div'] = div
                dic_tmp['type'] = key
                dic_tmp['title'] = item
                dic_tmp['HC'] = 0
                for i in range(1,4): # b_1, b_2, b_3, a_1, a_2, a_3
                    dic_tmp['b_'+str(i)] = round(random.uniform(0.8, 1.3),1)
                    dic_tmp['a_'+str(i)] = round(random.uniform(0.8, 1.3),1)  
                lst_return.append(dic_tmp)  

        return {'data':lst_return}

# def getWorkRate(div, df_prj, df_HC):
#     import pandas as pd
#     import rpy2.robjects as robjects
#     from rpy2.robjects import pandas2ri
#     from rpy2.robjects.conversion import localconverter

#     # Defining the R script and loading the instance in Python
#     robjects.r.source("./powerApp_func_v3.r")
#     # robjects.r.SayHi("John")

#     # Loading the function we have defined in R.
#     function_r = robjects.globalenv['hr_dept']

#     # Reading and processing data
#     # 讀取成Python的df
#     # df_prj = pd.read_csv("./data/df_prj.csv")
#     # df_HC = pd.read_csv("./data/df_HC.csv")
#     # div = "230000"

#     # pandas DataFrame去除na，有na值的話，直接轉換會error
#     df_prj.fillna("", inplace=True)
#     df_HC = df_HC.iloc[:,:-2]

#     # ================================
#     # === 將 Pandas.df 轉換成 R df ===
#     # ================================
#     #converting it into r object for passing into r function
#     with localconverter(robjects.default_converter + pandas2ri.converter):
#         df_prj_r = robjects.conversion.py2rpy(df_prj)
#         df_HC_r = robjects.conversion.py2rpy(df_HC)
#     # # df_prj_r = pandas2ri.py2ri(df_prj) # 舊寫法不能用了

#     # =====================================
#     # === 呼叫 R Function， return R df ===
#     # =====================================
#     #Invoking the R function and getting the result
#     df_result_r = function_r(df_prj_r, df_HC_r, div)

#     # ================================
#     # === 將 R df 轉換成  Pandas.df ===
#     # ================================
#     # Converting it back to a pandas dataframe.
#     with localconverter(robjects.default_converter + pandas2ri.converter):
#         df_result = robjects.conversion.rpy2py(df_result_r)

#     # df_result = pandas2ri.py2ri(df_result_r) # 舊寫法不能用了
#     return df_result



# api.add_resource(User, '/user/<string:name>')
api.add_resource(WorkRate, '/workrate/')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0') 
