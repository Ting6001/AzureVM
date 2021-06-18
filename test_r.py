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
df_prj = pd.read_csv("./data/df_prj.csv")
df_HC = pd.read_csv("./data/df_HC.csv")
div = "230000"
print(df_prj)
print(df_HC)
print(df_prj.dtypes)
print(df_HC.dtypes)
print('===============================================')
# pandas DataFrame去除na，有na值的話，直接轉換會error
df_prj.fillna("", inplace=True)
df_HC = df_HC.iloc[:,:-2]
print(df_prj)
print(df_HC)
print(df_prj.dtypes)
print(df_HC.dtypes)
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
print(df_result)