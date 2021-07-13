lst_time = []
for i in range(1):
  import pandas as pd
  import rpy2.robjects as robjects
  from rpy2.robjects import pandas2ri
  from rpy2.robjects.conversion import localconverter
  import time
  import numpy as np

  start = time.process_time()
  # Defining the R script and loading the instance in Python
  # robjects.r.source("./powerApp_func_v11.R")
  robjects.r.source("./powerApp_func_multi_v2.r")

  # robjects.r.SayHi("John")

  # Loading the function we have defined in R.
  # function_r = robjects.globalenv['hr_cal']
  function_r = robjects.globalenv['hr_cal_multi']

  # Reading and processing data
  # 讀取成Python的df
  # df_prj = pd.read_csv("./data/df_prj.csv")
  # df_HC = pd.read_csv("./data/df_HC.csv")
  # df_util = pd.read_csv("./data/UtilizationRateInfo_3.csv")
  div = "23D000"
  user = "001"

  lst_prj = [
      {
        "project_code": "1",
        "project_name": "New Project A",
        "project_code_old": "91.10P55.001",
        "project_name_old": "AD158",
        "execute_hour": np.random.randint(3000, 8000),
        "execute_month": 16
      }
    ]

  lst_HC = [
      {
        "HC": 2,
        "div": "23D000",
        "deptid": "23D000",
        "project_code": "1",
        "project_name": "New Project A",
        "sub_job_family": "HW",
        "save_time": "2021/7/12 下午06:02",
        "user_id": "001"
      },
      {
        "HC": 2,
        "div": "23D000",
        "deptid": "23D100",
        "project_code": "1",
        "project_name": "New Project A",
        "sub_job_family": "HW",
        "save_time": "2021/7/12 下午06:02",
        "user_id": "001"
      }
    ]


  df_prj = pd.DataFrame(lst_prj)
  df_HC = pd.DataFrame(lst_HC)
  # print('===============================================')
  # pandas DataFrame去除na，有na值的話，直接轉換會error
  # print(df_prj.isna())
  # df_prj.fillna("", inplace=True)
  df_HC = df_HC.iloc[:,:-2]

  lst_str = ["emplid",
  "div",
  "deptid",
  "sub_job_family",
  "project_code",
  "project_start_stage",
  "stage"
  ]
  dic_type = {'str':lst_str}
  dic_str = { i : str for i in lst_str}
  df_prj = df_prj.astype({'project_code':str, 'project_code_old':str})
  df_HC = df_HC.astype({'div':str, 'deptid':str, 'project_code':str})
  # print(df_prj)
  # print(df_HC)
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
  df_result_r = function_r(df_prj_r, df_HC_r, div)
  print('==== After call R function ====')
  # ================================
  # === 將 R df 轉換成  Pandas.df ===
  # ================================
  print('==== Transfer R.df to Pandas.df ====')
  with localconverter(robjects.default_converter + pandas2ri.converter):
    df_result = robjects.conversion.rpy2py(df_result_r)
          
  print('======== Function return ==========')
  dic_result = {}
  if isinstance(df_result, pd.DataFrame):
      df_result.fillna(0.0, inplace=True)
      dic_result = df_result.to_dict('records')
      # print(dic_result)
  time_taken = round(time.process_time() - start,2)
  print(i, 'Take:', time_taken, 's')
  lst_time.append(time_taken)

print(lst_time)
print('average:', round(sum(lst_time)/len(lst_time), 3), 's')
# Table_1：[10.83, 2.41, 2.45, 2.38, 1.61, 1.55, 1.58, 1.59, 1.66, 1.59]
# Table_1：average: 2.765 s
# Table_2：[9.08, 1.53, 1.5, 1.38, 1.34, 1.23, 1.41, 1.91, 1.81, 2.02]
# Table_2：average: 2.321 s
