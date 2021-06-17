from flask import Flask, request 
from flask_restful import Resource, Api,reqparse
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
        # print(arg)     
        set_dep = set()
        set_sub = set()
        df_prj = pd.DataFrame()
        df_HC = pd.DataFrame()

        div = arg['div']
        if arg['data_Prj']:
            df_prj = pd.DataFrame(arg['data_Prj'])
        if arg['data_HC']:
            df_HC = pd.DataFrame(arg['data_HC'])
        
        print('div:', div)
        # df_result = getWorkRate(div, df_prj, df_HC)
        # print(df_result)
        #################  先把code放進來 debug 用 ####################################################
        robjects.r.source("./powerApp_func.r")
        function_r = robjects.globalenv['hr_dept']

        # pandas DataFrame去除na，有na值的話，直接轉換會error
        # df = pd.read_excel("./data/powerapp_df.xlsx")
        df_prj.fillna("", inplace=True)
        df_HC = df_HC.drop(['savetime', 'UserId'], axis=1)
        print(df_prj)
        print('@@@@@@@@')
        print(df_HC)

        # === 將 Pandas.df 轉換成 R df ===
        with localconverter(robjects.default_converter + pandas2ri.converter):
            df_prj_r = robjects.conversion.py2rpy(df_prj)
            df_HC_r = robjects.conversion.py2rpy(df_HC)
            # df_r = robjects.conversion.py2rpy(df)
        # === 呼叫 R Function， return R df ===
        df_result_r = function_r(df_prj_r, df_HC_r, div)

        # === 將 R df 轉換成  Pandas.df ===
        with localconverter(robjects.default_converter + pandas2ri.converter):
            df_result = robjects.conversion.rpy2py(df_result_r)        
        
        dic_result = df_result.to_dict('records')
        return dic_result

        ##############################################################################################
        # 假資料區

def getWorkRate(div, df_prj, df_HC):
    import pandas as pd
    import rpy2.robjects as robjects
    from rpy2.robjects import pandas2ri
    from rpy2.robjects.conversion import localconverter

    # Defining the R script and loading the instance in Python
    robjects.r.source("./powerApp_func.r")
    # robjects.r.SayHi("John")

    # Loading the function we have defined in R.
    function_r = robjects.globalenv['hr_dept']

    # Reading and processing data
    # 讀取成Python的df
    # df_prj = pd.read_csv("./data/df_prj.csv")
    # df_HC = pd.read_csv("./data/df_HC.csv")
    # div = "230000"

    # pandas DataFrame去除na，有na值的話，直接轉換會error
    df_prj.fillna("", inplace=True)
    df_HC = df_HC.iloc[:,:-2]

    # ================================
    # === 將 Pandas.df 轉換成 R df ===
    # ================================
    #converting it into r object for passing into r function
    with localconverter(robjects.default_converter + pandas2ri.converter):
        df_prj_r = robjects.conversion.py2rpy(df_prj)
        df_HC_r = robjects.conversion.py2rpy(df_HC)
    # # df_prj_r = pandas2ri.py2ri(df_prj) # 舊寫法不能用了

    # =====================================
    # === 呼叫 R Function， return R df ===
    # =====================================
    #Invoking the R function and getting the result
    df_result_r = function_r(df_prj_r, df_HC_r, div)

    # ================================
    # === 將 R df 轉換成  Pandas.df ===
    # ================================
    # Converting it back to a pandas dataframe.
    with localconverter(robjects.default_converter + pandas2ri.converter):
        df_result = robjects.conversion.rpy2py(df_result_r)

    # df_result = pandas2ri.py2ri(df_result_r) # 舊寫法不能用了
    return df_result



# api.add_resource(User, '/user/<string:name>')
api.add_resource(WorkRate, '/workrate/')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0') 
