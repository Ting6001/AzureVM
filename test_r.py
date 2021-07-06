import pandas as pd
import rpy2.robjects as robjects
from rpy2.robjects import pandas2ri
from rpy2.robjects.conversion import localconverter
import time
start = time.process_time()
# Defining the R script and loading the instance in Python
robjects.r.source("./powerApp_func_v11.R")
# robjects.r.SayHi("John")

# Loading the function we have defined in R.
function_r = robjects.globalenv['hr_cal']

# Reading and processing data
# 讀取成Python的df
df_prj = pd.read_csv("./data/df_prj.csv")
df_HC = pd.read_csv("./data/df_HC.csv")
div = "23R000"
# print(df_prj)
# print(df_HC)
# print(df_prj.dtypes)
# print(df_HC.dtypes)
# print('===============================================')
# pandas DataFrame去除na，有na值的話，直接轉換會error
df_prj.fillna("", inplace=True)
df_HC = df_HC.iloc[:,:-2]
# print(df_prj)
# print(df_prj.isna())
# print(df_HC)
# print(df_HC.isna())
df_prj = df_prj.astype({'project_code':str, 'project_code_old':str})
df_HC = df_HC.astype({'div':str, 'deptid':str, 'project_code':str})
print(df_prj)
print(df_HC)
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
print('==== Before call R function ====')
df_result_r = function_r(df_prj_r, df_HC_r, div)
print('==== After call R function ====')
# ================================
# === 將 R df 轉換成  Pandas.df ===
# ================================
# Converting it back to a pandas dataframe.
print('==== Transfer R.df to Pandas.df ====')
with localconverter(robjects.default_converter + pandas2ri.converter):
  df_result = robjects.conversion.rpy2py(df_result_r)
        
print('======== Function return ==========')
dic_result = {}
if isinstance(df_result, pd.DataFrame):
    df_result.fillna(0.0, inplace=True)
    dic_result = df_result.to_dict('records')
    print(dic_result)
# df_result = pandas2ri.py2ri(df_result_r) # 舊寫法不能用了
time_taken = round(time.process_time() - start,2)
print('Take:', time_taken, 's')