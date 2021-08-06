import pyodbc 
import pandas as pd
from datetime import datetime
import sys
import traceback
from tqdm import tqdm
from datetime import date
# Some other example server values are
# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = 'sqlserver-mia.database.windows.net' 
database = 'DB-mia' 
username = 'admia' 
password = 'Mia01@wistron' 
driver = '{ODBC Driver 17 for SQL Server}'
driver = '{SQL Server Native Client 11.0}'
conn = pyodbc.connect('DRIVER='+driver+';SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = conn.cursor()

with open ('./data/project_code_183.txt', 'r') as f:
  lst_project = f.read().split('\n')

# lst_project = ['1PD006270001,', '1PD02W27A001,']
lst_month = list(map(str, range(1, 51)))
lst_info = ['project_code','sub_job_family_2', 'start_stage','pmcs_bu_start_dt', 'date_fst','date_lst']
lst_col = lst_info + lst_month
df_all = pd.DataFrame(columns=lst_col)

for i, project in enumerate(tqdm(lst_project)):
    try:
        project_code = project.split(',')[0]
        print(i+1, project_code)
        query = '''
            select ROW_NUMBER() OVER(ORDER BY (SELECT 1)) AS No,
            project_code, sub_job_family_2, project_start_stage start_stage,stage, convert(date,[date]) [date], 
            avg(C0_day) C0_day, avg(C1_day) C1_day, avg(C2_day) C2_day, 
            avg(C3_day) C3_day, avg(C4_day) C4_day, avg(C5_day) C5_day, avg(C6_day) C6_day, 
            count(*) cnt, sum(total_hour) sum_hour, round(sum(total_hour)/count(*),1) hour_per_cnt, convert(date,pmcs_bu_start_dt) pmcs_bu_start_dt
            from UtilizationRateInfo
            where project_code = '{}'
            and sub_job_family_2 = 'RH'
            and [date] < convert(date, '2021-07-01')
            group by project_code, sub_job_family_2, date, project_start_stage, stage, pmcs_bu_start_dt
            order by sub_job_family_2, date
        '''.format(project_code)

        df = pd.read_sql(query, conn)
        lst_sub = ['RH']
        for sub in lst_sub:
            df_sub = df[df['sub_job_family_2']==sub].reset_index()
            # print(sub, df_sub.shape)
            # df.set_index('No', inplace=True)
            if len(df_sub) > 0:
                dic = {}
                dic['sub_job_family_2'] = sub
                dic['project_code'] = project_code
                dic['start_stage'] = df_sub.loc[0, 'start_stage']
                dic['pmcs_bu_start_dt'] = df_sub.loc[0, 'pmcs_bu_start_dt']
                dic['date_fst'] = df_sub.loc[0, 'date']
                dic['date_lst'] = df_sub.loc[df_sub.index[-1], 'date']

                for i in df_sub.index: # index 從0開始
                    dic[str(i+1)] = df_sub.loc[i, 'hour_per_cnt']
                df_tmp = pd.DataFrame(dic, index=[0])
                df_all = df_all.append(df_tmp)
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

p_file = './data/hour_per_cnt_bySub_183.xlsx'
df_all.to_excel(p_file, index=False)
# df_all = pd.read_excel('./data/hour_per_cnt_bySub_70.xlsx')
print(df_all.shape)
lst_sub = ['RH']
for sub in lst_sub:
    df_sub = df_all[df_all['sub_job_family_2']==sub]
    print(sub, df_sub.shape)
    num_empty = len(lst_info) 

    lst_max = [''] * num_empty + df_sub.max(axis=0).to_list()[num_empty:]
    lst_min = [''] * num_empty + df_sub.min(axis=0).to_list()[num_empty:]
    lst_mean = [''] * num_empty + df_sub.fillna(0).iloc[:,num_empty:].mean(axis=0).round(1).to_list()
    lst_50 = []
    lst_100 = []
    lst_200 = []
    lst_300 = []
    for month in lst_month:
        df_col = df_sub[[month]]

        lst_50.append(len(df_col[ df_col[month] < 50]))
        lst_100.append(len(df_col[ (df_col[month] > 50) & (df_col[month] < 100) ]))
        lst_200.append(len(df_col[ (df_col[month] > 100) & (df_col[month] < 200) ]))
        lst_300.append(len(df_col[ (df_col[month] > 200) & (df_col[month] < 300) ]))


    df_max = pd.DataFrame([lst_max], columns=lst_col, index=['max'])
    df_min = pd.DataFrame([lst_min], columns=lst_col, index=['min'])
    df_mean = pd.DataFrame([lst_mean], columns=lst_col, index=['mean'])

    df_50 = pd.DataFrame([[''] * num_empty + lst_50], columns=lst_col, index=['__50'])
    df_100 = pd.DataFrame([[''] * num_empty + lst_100], columns=lst_col, index=['__100'])
    df_200 = pd.DataFrame([[''] * num_empty + lst_200], columns=lst_col, index=['__200'])
    df_300 = pd.DataFrame([[''] * num_empty + lst_300], columns=lst_col, index=['__300'])
    # print(df_all.isna())
    df_new = pd.concat([df_sub, df_max, df_min, df_mean, df_50, df_100, df_200, df_300])
    print(df_new.shape)

    df_new.to_excel('./data/hour_per_cnt_bySub_183{}.xlsx'.format(sub))
cursor.close()
conn.close()



