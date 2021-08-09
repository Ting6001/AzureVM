lst_time = []
# for i in range(1):
import pandas as pd
import rpy2.robjects as robjects
from rpy2.robjects import pandas2ri
from rpy2.robjects.conversion import localconverter
import time
import numpy as np
from tqdm import tqdm
import sys, traceback
from datetime import datetime
start_begin = time.process_time()

# Defining the R script and loading the instance in Python
# robjects.r.source("./powerApp_func_v11.R")
# robjects.r.source("./powerApp_func_multi_v3.R")
robjects.r.source("./powerApp_func_multi_v7_test.R")

# robjects.r.SayHi("John")

# Loading the function we have defined in R.
# function_r = robjects.globalenv['hr_cal']
function_r = robjects.globalenv['hr_cal_multi']

# Reading and processing data
# 讀取成Python的df
# df_prj = pd.read_csv("./data/df_prj.csv")
# df_HC = pd.read_csv("./data/df_HC.csv")
# df_util = pd.read_csv("./data/UtilizationRateInfo_3.csv")
div = "23R000"
user = "001"

path_file = './data/Utilization_Add1proj.xlsx'
# df_allProj = pd.read_excel(path_file, engine='openpyxl')

df_allProj = pd.DataFrame(columns=['No', 'project_code', 'Div', 'Type', 'Title', 'b_1', 'b_2', 'b_3', 'hc', 'a_1', 'a_2', 'a_3'])
lst_div = ['23D000', '23M000', '23R000']

with open ('./data/project_code.txt', 'r') as f:
  lst_project = f.read().split('\n')
count_project = len(lst_project)


# '1PD062550001,QA618 JAWS,21498.20,3.17'
for i, project_code_name in enumerate(tqdm(['1PD062550001,QA618 JAWS,37529.36,3.87'])): # 23D000、1PD05R550001,  23R000、1PD062550001,1RDZ12280001
  project_code = project_code_name.split(',')[0]
  project_name = project_code_name.split(',')[1]
  execute_hour = float(project_code_name.split(',')[2]) * 1
  execute_month = float(project_code_name.split(',')[3]) * 1
  # if i+1 <= 25:
  #   continue
  start = time.process_time()
  lst_prj = [
      # {
      #   "project_code": "0",
      #   "project_name": "new project A",
      #   "project_code_old": project_code,
      #   "project_name_old": project_name,
      #   "execute_hour": execute_hour,
      #   "execute_month": execute_month
      # },
      {
        "project_code": "0",
        "project_name": "new project B",
        "project_code_old": project_code,
        "project_name_old": project_name,
        "execute_hour": execute_hour,
        "execute_month": execute_month,
         "start_date": "2021/9/9"
      }
    ]
  lst_HC = [
    # {
    #   "HC": 1,
    #   "div": div,
    #   "deptid": "23R000",
    #   "project_code": "0",
    #   "project_name": "new project A",
    #   "sub_job_family": "HW",
    #   "save_time": "2021/7/6 下午05:15",
    #   "user_id": "001"
    # },
    # {
    #   "HC": 1,
    #   "div": div,
    #   "deptid": "23R200",
    #   "project_code": "",
    #   "project_name": "",
    #   "sub_job_family": "HW",
    #   "save_time": "2021/7/6 下午05:15",
    #   "user_id": "001"
    # }  
  ]
  print({"user": "001", 'division':div, 'df_prj':lst_prj, 'df_HC':lst_HC})
  df_prj = pd.DataFrame(lst_prj)
  df_HC = pd.DataFrame(lst_HC)
  # print('===============================================')
  # pandas DataFrame去除na，有na值的話，直接轉換會error
  # print(df_prj.isna())
  # df_prj.fillna("", inplace=True)
  if not df_prj.empty:
    df_prj = df_prj.astype({'project_code':str, 'project_code_old':str})
  if not df_HC.empty:
    df_HC = df_HC.iloc[:,:-2]
    df_HC = df_HC.astype({'div':str, 'deptid':str, 'project_code':str})

  # lst_str = ["emplid",
  # "div",
  # "deptid",
  # "sub_job_family",
  # "project_code",
  # "project_start_stage",
  # "stage"
  # ]
  # dic_type = {'str':lst_str}
  # dic_str = { i : str for i in lst_str}
  
  # ================================
  # === 將 Pandas.df 轉換成 R df ===
  # ================================
  with localconverter(robjects.default_converter + pandas2ri.converter):
      df_prj_r = robjects.conversion.py2rpy(df_prj)
      df_HC_r = robjects.conversion.py2rpy(df_HC)

  # =====================================
  # === 呼叫 R Function， return R df ===
  # =====================================
  print('==== Before call R function ====')
  try:
    df_result_r = function_r(df_prj_r, df_HC_r, div)
    print('==== After call R function ====')
    # ================================
    # === 將 R df 轉換成  Pandas.df ===
    # ================================
    # print('==== Transfer R.df to Pandas.df ====')
    with localconverter(robjects.default_converter + pandas2ri.converter):
      df_result = robjects.conversion.rpy2py(df_result_r)
            
    # print('======== Function return ==========')
    dic_result = {}
    # print(type(df_result))
    print(df_result)
    # print(df_result.columns)
    # df_result['date'] = df_result['date'].astype('datetime64[ns]')
    f_name = './data/test/__0909__df_proj_hour_{}_proj_1_Ratio_00_HC_0.xlsx'.format(project_code)
    
    lst_sort = ['sub_job_family_2','date']
    print('ori', type(df_result.at['1', 'date']), df_result.at['1', 'date'])
    # df_proj_hour date float
    if 'df_proj_hour' in f_name:
      df_result['date'] = df_result['date'].apply(lambda x: pd.to_datetime(x,unit='D', origin='1970-1-1'))
      df_result['start_date'] = df_result['start_date'].apply(lambda x: pd.to_datetime(x,unit='D', origin='1970-1-1'))
      df_result['add_hour'] = df_result['add_hour'].apply(lambda x: round(x,1))  
      print('pd.to_datetime(x)', type(df_result.at['1', 'date']), df_result.at['1', 'date'])

    df_result['date'] = df_result['date'].apply(lambda x: x.date())
    print('x.date()', type(df_result.at['1', 'date']), df_result.at['1', 'date'])
    base_date = datetime.strptime("2021-08-01", "%Y-%m-%d").date()
    # df_result = df_result[df_result['date'] > base_date]
    
    if 'rate_tmp' in f_name:
      df_result['start_date'] = df_result['start_date'].apply(lambda x: pd.to_datetime(x,unit='D', origin='1970-1-1'))
      df_result['add_hour'] = df_result['add_hour'].apply(lambda x: round(x,1))  
      df_result['uti_rate'] = df_result['uti_rate'].apply(lambda x: round(x,1))
      df_result['add_hour_pct'] = df_result['add_hour_pct'].apply(lambda x: round(x,1))
      df_result['total_hour_dept_func'] = df_result['total_hour_dept_func'].apply(lambda x: round(x,1))

    if 'df_dept_rate_future' in f_name:
      df_result['total_hour_by_dep_func_cal'] = df_result['total_hour_by_dep_func_cal'].apply(lambda x: round(x,1))
      df_result['add_hour_pct'] = df_result['add_hour_pct'].apply(lambda x: round(x,1))
      df_result['total_hour_dept_func'] = df_result['total_hour_dept_func'].apply(lambda x: round(x,1))
      lst_sort = ['date']
    
    df_result.sort_values(lst_sort, inplace=True)
    
    # print(df_result)
    df_result.to_excel(f_name, index=False)

    # if isinstance(df_result, pd.DataFrame):
    #     df_result.fillna(0.0, inplace=True)
    #     dic_result = df_result.to_dict('records')
    #     # print(dic_result)
    #     df_result['No'] = i+1
    #     df_result['project_code'] = ''
    #     # df_result['project_name'] = ''
    #     for j in range(1,4):
    #       df_result['a_'+str(j)] = df_result['a_'+str(j)].apply(lambda x: x*100)
    #       df_result['b_'+str(j)] = df_result['b_'+str(j)].apply(lambda x: x*100)
    #     df_allProj = df_allProj.append(df_result)

    # time_taken = round(time.process_time() - start,2)
    # print('Take:', time_taken, 's', i+1, project_code)
    # lst_time.append(time_taken)
  except Exception as e:
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
        print(err_msg)
    break

# time_total = round(time.process_time() - start_begin, 2)
# print('Total Time take:', time_total, 's')
# print(lst_time)
# print(df_allProj)
df_allProj.to_excel(path_file, index=False)

# print(lst_time)
# print('average:', round(sum(lst_time)/len(lst_time), 3), 's')
# Table_1：[10.83, 2.41, 2.45, 2.38, 1.61, 1.55, 1.58, 1.59, 1.66, 1.59]
# Table_1：average: 2.765 s
# Table_2：[9.08, 1.53, 1.5, 1.38, 1.34, 1.23, 1.41, 1.91, 1.81, 2.02]
# Table_2：average: 2.321 s
