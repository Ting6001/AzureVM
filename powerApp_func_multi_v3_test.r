hr_cal_multi <- function(df_prj,
                         df_hc,
                         division = NA){
  ## Package names
  packages <- c('data.table', 
                'tidyverse', 
                'dplyr', 
                'readxl',
                'lubridate', 
                'reshape2',
                'odbc',
                'DBI')
  
  ## Install packages not yet installed
  installed_packages <- packages %in% rownames(installed.packages())
  if (any(installed_packages == FALSE)) {
    install.packages(packages[!installed_packages])
  }
  
  ## Packages loading
  invisible(lapply(packages, library, character.only = TRUE))
  
  ## Do not show warnings & Suppress summarise info
  options(warn = -1,
          dplyr.summarise.inform = FALSE)
  
  ###----- Load Data ------------------------------
  odbc_driver = odbcListDrivers() %>%
    filter(str_detect(name, '^ODBC')) %>%
    distinct(name) %>%
    slice(1)
  conn <- dbConnect(odbc(),
                    Driver = odbc_driver$name,
                    Server = "sqlserver-mia.database.windows.net",
                    UID = "admia",
                    PWD = "Mia01@wistron",
                    database = "DB-mia")

  df_all = dbGetQuery(conn, "select * from [dbo].[UtilizationRateInfo]")
  dbDisconnect(conn)
  
  ## When sub job family is ('AR', 'HW', 'OP', 'RF', 'RX') replace by 'RH'
  df_all <- df_all %>%
    mutate(sub_job_family = ifelse(sub_job_family %in% c('AR', 'HW', 'OP', 'RF', 'RX'), 'RH', sub_job_family))
  
  if (is.na(division) == F){
    df <- df_all %>%
      filter(div == division)
  }
  
  ## Project stage days (all Divisions)
  df_proj <- df_all %>%
    filter(pmcs_bu_start_dt >= as.character(min(df_all$date))) %>%  ## 篩選 專案起始日期(pmcs_bu_start_dt) > min(該筆資料年月日date)
    group_by(project_code) %>%
    summarise(C0 = mean(C0_day),
              C1 = mean(C1_day),
              C2 = mean(C2_day),
              C3 = mean(C3_day),
              C4 = mean(C4_day),
              C5 = mean(C5_day),
              C6 = mean(C6_day),
              execute_hour = mean(execute_hour),   ## 取平均：C0~C6、執行時數、執行月數
              execute_month = mean(execute_month), ## 取最大：執行日數
              execute_day = max(execute_day),
              project_start_stage = unique(project_start_stage)) %>%
    ungroup()
  
  ## settings
  if (sum(dim(df_hc)) != 0){
    head_count <- df_hc %>%
      mutate(HC = as.numeric(HC),
             sub_job_family = ifelse(sub_job_family %in% c('AR', 'HW', 'OP', 'RF', 'RX'), 'RH', sub_job_family)) %>%
      group_by(div, deptid, project_code, project_name, sub_job_family) %>%
      summarise(hc = sum(HC)) %>%
      ungroup() %>%
      group_by(project_code, sub_job_family) %>%
      mutate(hc_pct = hc / sum(hc)) %>%
      ungroup()
  } else{
    head_count <- data.frame(div = NA, deptid = NA, project_code = NA, project_name = NA, sub_job_family = NA, hc = 0, hc_pct = 0)
  }
  
  
  ###----- Calculate ------------------------------
  if (length(setdiff(df_prj$project_code_old, df_proj$project_code)) != 0){ # 確認PowerApps 所選專案有在篩出來的df_proj裡面 專案起始日期(pmcs_bu_start_dt) > min(該筆資料年月日date)
    print('Invalid project code!!')
  } else {
    if (sum(dim(df_prj)) == 0){
      #------------#
      #- Function -#
      #------------#
      ## utilization rate
      df_func_rate <- df %>%                                            ## df：以所選div，篩選出的資料
        group_by(div, sub_job_family, date) %>%                         ## group by (處級、次職類、該筆資料年月日)
        summarise(uti_rate = mean(utilization_rate_by_div_func),        ## 工時率 = Avg(當月處底下"次職類"工時使用率utilization_rate_by_div_func)
                  total_hour_func = sum(total_hour),                    ## 加總  ：個人每月PTS專案總工時(total_hour)、應到班工時
                  attendance_func = sum(attendance),
                  emp_cnt = n_distinct(emplid) - sum(termination_n),    ## 每個月的員工數：該月員工ID數 - 該月預計離職人數
                  attendance_emp = mean(attendance)) %>%                ## 平均值：該月應到班工時(attendance)
        arrange(date) %>%
        ungroup()
      
      ## output table
      out_func <- df_func_rate %>%
        mutate(type = 'SUB') %>%
        select(div, type, sub_job_family) %>%
        distinct() %>%                                                   ## 取出 計算過去三個月工時率 的資料(目前7月的話, 取456)
        left_join(df_func_rate %>%                                       ## max(date)到未來第三個月(目前7月的話,未來是789)
                    filter(date >= max(ymd(df_all$date)) %m-% months(5) & date <= max(ymd(df_all$date)) %m-% months(3)) %>%
                    select(div, sub_job_family, date, uti_rate) %>%
                    mutate(uti_rate = round(uti_rate, 2)) %>%
                    spread(date, uti_rate),
                  by = c('div', 'sub_job_family')) %>%
        left_join(head_count %>%
                    select(-c(deptid, hc_pct)),
                  by = c("div", "sub_job_family")) %>%
        left_join(df_func_rate %>%                                      ## 取出 計算未來三個月工時率 的資料(目前7月的話, 取789)
                    filter(date >= max(ymd(df_all$date)) %m-% months(2) & date <= max(ymd(df_all$date))) %>%
                    select(div, sub_job_family, date, uti_rate) %>%
                    mutate(uti_rate = round(uti_rate, 2)) %>%
                    spread(date, uti_rate),
                  by = c('div', 'sub_job_family'))
      out_func[is.na(out_func)] <- 0
      names(out_func) <- c('Div', 'Type', 'Title', 'b_1', 'b_2', 'b_3', 'hc', 'a_1', 'a_2', 'a_3')
      
      if (nrow(out_func) != 0){
        out_func <- out_func
      } else{
        out_func <- df_func_rate %>%
          mutate(type = 'SUB') %>%
          select(div, type, sub_job_family) %>%
          distinct() %>%
          rename(Div = div,
                 Type = type,
                 Title = sub_job_family) %>%
          mutate(b_1 = 0, 
                 b_2 = 0, 
                 b_3 = 0, 
                 hc = 0, 
                 a_1 = 0, 
                 a_2 = 0, 
                 a_3 = 0)
      }
      
      
      #--------------#
      #- Department -#
      #--------------#
      ## utilization rate
      df_dept_rate <- df %>%
        group_by(div, deptid, sub_job_family, date) %>%
        summarise(uti_rate = mean(utilization_rate_by_dep_func),    ## 工時率 = Avg(每月部門次職類專案工時使用率 utilization_rate_by_dep_func)
                  total_hour_dept_func = sum(total_hour),
                  emp_cnt = n_distinct(emplid) - sum(termination_n),
                  attendance_emp = mean(attendance)) %>%
        arrange(date) %>%
        ungroup()
      
      ## output table
      out_dept <- df_dept_rate %>%
        mutate(type = 'DEP') %>%
        select(div, type, deptid) %>%
        distinct() %>%                                                   ## 取出 計算過去三個月工時率 的資料(目前7月的話, 取456)
        left_join(df_dept_rate %>%
                    filter(date >= max(ymd(df_all$date)) %m-% months(5) & date <= max(ymd(df_all$date)) %m-% months(3)) %>%
                    group_by(div, deptid, date) %>%
                    summarise(uti_rate = round(mean(uti_rate), 2)) %>%
                    ungroup() %>%
                    spread(date, uti_rate),
                  by = c('div', 'deptid')) %>%
        left_join(head_count %>%
                    group_by(div, deptid) %>%
                    summarise(hc = sum(hc)),
                  by = c("div", "deptid")) %>%
        left_join(df_dept_rate %>%                                      ## 取出 計算未來三個月工時率 的資料(目前7月的話, 取789)
                    filter(date >= max(ymd(df_all$date)) %m-% months(2) & date <= max(ymd(df_all$date))) %>%
                    group_by(div, deptid, date) %>%
                    summarise(uti_rate = round(mean(uti_rate), 2)) %>%
                    ungroup() %>%
                    spread(date, uti_rate),
                  by = c('div', 'deptid'))
      out_dept[is.na(out_dept)] <- 0
      names(out_dept) <- c('Div', 'Type', 'Title', 'b_1', 'b_2', 'b_3', 'hc', 'a_1', 'a_2', 'a_3')
      
      if (nrow(out_dept) != 0){
        out_dept <- out_dept
      } else{
        out_dept <- df_dept_rate %>%
          mutate(type = 'DEP') %>%
          select(div, type, deptid) %>%
          distinct() %>%
          rename(Div = div,
                 Type = type,
                 Title = deptid) %>%
          mutate(b_1 = 0, 
                 b_2 = 0, 
                 b_3 = 0, 
                 hc = 0, 
                 a_1 = 0, 
                 a_2 = 0, 
                 a_3 = 0)
      }
      
      
      #-----------------------------------#
      #- Combine Department and Function -#
      #-----------------------------------#
      out <- bind_rows(out_dept,
                       out_func)
    } else{  ## ============================ 【有新增專案 的工時率計算】 ============================ 
      ## Only calculate have both new and old project
      df_prj <- df_prj %>%
        filter(!(is.na(project_code) | project_code == ''),
               !(is.na(project_code_old) | project_code_old == ''))
      
      
      #------------#
      #- Function -#
      #------------#
      ## Previous utilization rate
      df_func_rate <- df %>%                                               ## df：以所選div，篩選出的資料
        group_by(div, sub_job_family, date) %>%                            ## group by (處級、次職類、該筆資料年月日)
        summarise(uti_rate = mean(utilization_rate_by_div_func),           ## 工時率 = Avg(當月處底下"次職類"工時使用率utilization_rate_by_div_func)
                  total_hour_func = sum(total_hour),                       ## 加總  ：執行時數、應到班工時
                  attendance_func = sum(attendance),                        
                  emp_cnt = n_distinct(emplid) - sum(termination_n),       ## 每個月的員工數：該月員工ID數 - 該月預計離職人數
                  attendance_emp = mean(attendance)) %>%                   ## 平均值：該月應到班工時(attendance)
        arrange(date) %>%
        ungroup()
      
      ## New and Old project proportion
      proj_prop <- data.frame()
      for (newname in unique(df_prj$project_name)){
        df_tmp <- df_prj %>%
          filter(project_name == newname)
        
        proj_tmp <- data.frame()
        for (pcode in unique(df_tmp$project_code_old)){
          df_tmp1 <- df_tmp %>%
            filter(project_code_old == pcode)
          
          tmp <- df_proj %>%                                              ## df_proj：篩選 專案起始日期(pmcs_bu_start_dt) > min(該筆資料年月日date)
            filter(project_code == pcode) %>%                             ## 執行月數、執行時數=0 比例：1，其它情況：比例=新專案/舊專案
            select(project_code, execute_hour, execute_month, project_start_stage) %>%
            mutate(hr_ratio = ifelse(unique(df_tmp1$execute_hour) %in% c(0, '') | is.na(unique(df_tmp1$execute_hour)), 1, unique(df_tmp1$execute_hour) / execute_hour),
                   mon_ratio = ifelse(unique(df_tmp1$execute_month) %in% c(0, '') | is.na(unique(df_tmp1$execute_month)), 1, unique(df_tmp1$execute_month) / execute_month))
          
          proj_tmp <- bind_rows(proj_tmp,
                                tmp)
        }
        
        proj_prop <- bind_rows(proj_prop,
                               proj_tmp)
      }
      
      ## Calculate next three month utilization rate
      df_func_rate_future <- data.frame() 
      for (newname in unique(df_prj$project_name)){
        pcode = df_prj$project_code_old[df_prj$project_name == newname]
        
        hc_tmp <- head_count %>%
          filter(project_name == newname) %>%                           ## 取出該新專案 所新增的HC數
          select(-deptid)
        
        if (nrow(hc_tmp) == 0){
          hc_tmp <- data.frame(div = NA, project_code = NA, project_name = NA, sub_job_family = NA, hc = NA, hc_pct = NA)
        }
        
        
        proj_prop_tmp <- proj_prop %>%
          filter(project_code == pcode)
        
        rate_tmp <- data.frame()
        for (i in 1:nrow(proj_prop_tmp)){
          df_proj_hour <- df_all %>%
            filter(project_code == proj_prop_tmp$project_code[i]) %>%
            group_by(project_code, sub_job_family, date) %>%
            summarise(total_hour_pro_func = sum(total_hour)) %>%
            arrange(date) %>%
            ungroup() %>%                                        ## 增加的時數 = DB裡每個月總時數 * 新舊專案工時比例 / 新舊專案月數比例？
            mutate(add_hour = (total_hour_pro_func * proj_prop_tmp$hr_ratio[i]) / (proj_prop_tmp$mon_ratio[i])) %>%
            group_by(project_code, sub_job_family) %>%
            mutate(n = row_number()) %>%
            ungroup() %>%
            filter(n <= 3) %>%
            mutate(date = case_when(n == 1 ~ max(ymd(df_all$date)) %m-% months(2),
                                    n == 2 ~ max(ymd(df_all$date)) %m-% months(1),
                                    n == 3 ~ max(ymd(df_all$date)),
                                    TRUE ~ ymd(date))) %>%
            select(-n)
          
          rate_tmp1 <- df_func_rate %>%
            left_join(hc_tmp,
                      by = c("div", "sub_job_family")) %>%
            replace_na(list(hc = 0,
                            hc_pct = 1)) %>%
            left_join(df_proj_hour %>%
                        select(-project_code),
                      by = c('sub_job_family', 'date')) %>%
            group_by(div, date, sub_job_family) %>%
            mutate(add_hour_pct = add_hour * hc_pct,                     ## 增加的時數pct = 增加的時數 * hc_pct？??????
                   add_attendance_emp = (attendance_emp * hc)) %>%       ## 新增的應到工時 = 平均該月應到班工時(attendance_emp) * HC數
            ungroup() %>%
            distinct()
          
          rate_tmp <- bind_rows(rate_tmp,
                                rate_tmp1)
        }
        
        
        df_func_rate_future <- bind_rows(df_func_rate_future,
                                         rate_tmp) %>%
          group_by(div, date, sub_job_family) %>%
          summarise(total_hour_func = unique(total_hour_func),
                    attendance_func = unique(attendance_func),
                    add_hour_pct = sum(add_hour_pct, na.rm = T),
                    add_attendance_emp = sum(add_attendance_emp, na.rm = T),
                    hc = sum(hc, na.rm = T), 
                    
                    total_hour_by_func_cal = (total_hour_func + add_hour_pct),  ## 總工時 = DB裡每月的總工時 + 新增工時pct ???????????????
                    uti_rate_cal = round(total_hour_by_func_cal / (attendance_func + add_attendance_emp), 2)) %>%
          ungroup()
      }
      
      ## Combine previous and future rate
      out_func <- df_func_rate %>%
        mutate(type = 'SUB') %>%
        select(div, type, sub_job_family) %>%
        distinct() %>%
        left_join(df_func_rate %>%
                    filter(date >= max(ymd(df_all$date)) %m-% months(5) & date <= max(ymd(df_all$date)) %m-% months(3)) %>%
                    select(div, sub_job_family, date, uti_rate) %>%
                    mutate(uti_rate = round(uti_rate, 2)) %>%
                    spread(date, uti_rate),
                  by = c('div', 'sub_job_family')) %>%
        left_join(df_func_rate_future %>%
                    filter(date >= max(ymd(df_all$date)) %m-% months(2) & date <= max(ymd(df_all$date))) %>%
                    select(div, sub_job_family, hc, date, uti_rate_cal) %>%
                    mutate(uti_rate_cal = round(uti_rate_cal, 2)) %>%
                    spread(date, uti_rate_cal),
                  by = c('div', 'sub_job_family'))
      out_func[is.na(out_func)] <- 0
      names(out_func) <- c('Div', 'Type', 'Title', 'b_1', 'b_2', 'b_3', 'hc', 'a_1', 'a_2', 'a_3')
      
      if (nrow(out_func) != 0){
        out_func <- out_func
      } else{
        out_func <- df_func_rate %>%
          mutate(type = 'SUB') %>%
          select(div, type, sub_job_family) %>%
          distinct() %>%
          rename(Div = div,
                 Type = type,
                 Title = sub_job_family) %>%
          mutate(b_1 = 0, 
                 b_2 = 0, 
                 b_3 = 0, 
                 hc = 0, 
                 a_1 = 0, 
                 a_2 = 0, 
                 a_3 = 0)
      }
      
      
      #--------------#
      #- Department -#
      #--------------#
      ## Previous utilization rate
      df_dept_rate <- df %>%
        group_by(div, deptid, sub_job_family, date) %>%
        summarise(uti_rate = mean(utilization_rate_by_dep_func),
                  total_hour_dept_func = sum(total_hour),
                  emp_cnt = n_distinct(emplid) - sum(termination_n),
                  attendance_emp = mean(attendance)) %>%
        arrange(date) %>%
        ungroup()
      
      ## Calculate next three month utilization rate
      df_dept_rate_future <- data.frame()
      for (newname in unique(df_prj$project_name)){
        pcode = df_prj$project_code_old[df_prj$project_name == newname]
        
        hc_tmp <- head_count %>%
          filter(project_name == newname)
        if (nrow(hc_tmp) == 0){
          hc_tmp <- data.frame(div = NA, deptid = NA, project_code = NA, project_name = NA, sub_job_family = NA, hc = NA, hc_pct = NA)
        }
        
        
        proj_prop_tmp <- proj_prop %>%
          filter(project_code == pcode)
        
        rate_tmp <- data.frame()
        df_test <- data.frame()
        for (i in 1:nrow(proj_prop_tmp)){
          # df_test <- df_all %>%
          #   filter(project_code == proj_prop_tmp$project_code[i]) %>%
          #   group_by(project_code, sub_job_family, date) %>%
          #   summarise(total_hour_pro_func = sum(total_hour)) %>%
          #   arrange(date) %>%
          #   ungroup() %>%
          #   mutate(add_hour = (total_hour_pro_func * proj_prop_tmp$hr_ratio[i]) / (proj_prop_tmp$mon_ratio[i]))

          df_proj_hour <- df_all %>%
            filter(project_code == proj_prop_tmp$project_code[i]) %>%
            group_by(project_code, sub_job_family, date) %>%
            summarise(total_hour_pro_func = sum(total_hour)) %>%
            arrange(date) %>%
            ungroup() %>%
            mutate(add_hour = (total_hour_pro_func * proj_prop_tmp$hr_ratio[i]) / (proj_prop_tmp$mon_ratio[i])) %>%
            group_by(project_code, sub_job_family) %>%
            mutate(n = row_number()) %>%
            ungroup() %>%
            filter(n <= 3) %>%
            mutate(date = case_when(n == 1 ~ max(ymd(df_all$date)) %m-% months(2),
                                    n == 2 ~ max(ymd(df_all$date)) %m-% months(1),
                                    n == 3 ~ max(ymd(df_all$date)),
                                    TRUE ~ ymd(date))) %>%
            select(-n)
          print('@@@@@@@@@@@@@@@@@@@@@ 111111111111111')
          # print(df_test)
          # return(df_proj_hour)
          rate_tmp1 <- df_dept_rate %>%
            left_join(hc_tmp,
                      by = c("div", "deptid", "sub_job_family")) %>%
            replace_na(list(hc = 0,
                            hc_pct = 1)) %>%
            left_join(df_proj_hour %>%
                        select(-project_code),
                      by = c('sub_job_family', 'date')) %>%
            group_by(div, date, deptid, sub_job_family) %>%
            mutate(add_hour_pct = add_hour * hc_pct,
                   add_attendance_emp = (attendance_emp * hc)) %>%
            ungroup() %>%
            distinct()
          
          rate_tmp <- bind_rows(rate_tmp,
                                rate_tmp1)
        }
        print('@@@@@@@@@@@@@@@@@@@@@ 111111111111111')
        return(rate_tmp)
        df_dept_rate_future <- bind_rows(df_dept_rate_future,
                                         rate_tmp) %>%
          group_by(div, date, deptid, sub_job_family) %>%
          summarise(total_hour_dept_func = unique(total_hour_dept_func),
                    attendance_emp = unique(attendance_emp),
                    emp_cnt = unique(emp_cnt),
                    add_hour_pct = sum(add_hour_pct, na.rm = T),
                    add_attendance_emp = sum(add_attendance_emp, na.rm = T),
                    hc = sum(hc, na.rm = T),
                    
                    total_hour_by_dep_func_cal = (total_hour_dept_func + add_hour_pct),
                    uti_rate_cal_dept = round(total_hour_by_dep_func_cal / (attendance_emp * emp_cnt + add_attendance_emp), 2)) %>%
          ungroup()
      }
      print('@@@@@@@@@@@@@@@@@@@@@ 2222222222222222')
      print(df_dept_rate_future)
      return(df_dept_rate_future)
      ## Combine previous and future rate
      out_dept <- df_dept_rate %>%
        mutate(type = 'DEP') %>%
        select(div, type, deptid) %>%
        distinct() %>%
        left_join(df_dept_rate %>%
                    filter(date >= max(ymd(df_all$date)) %m-% months(5) & date <= max(ymd(df_all$date)) %m-% months(3)) %>%
                    group_by(div, deptid, date) %>%
                    summarise(uti_rate = round(mean(uti_rate), 2)) %>%
                    ungroup() %>%
                    spread(date, uti_rate),
                  by = c('div', 'deptid')) %>%
        left_join(df_dept_rate_future %>%
                    filter(date >= max(ymd(df_all$date)) %m-% months(2) & date <= max(ymd(df_all$date))) %>%
                    group_by(div, deptid, date) %>%
                    summarise(hc = mean(hc),
                              uti_rate_cal_dept = round(mean(uti_rate_cal_dept), 2)) %>%
                    ungroup() %>%
                    spread(date, uti_rate_cal_dept),
                  by = c('div', 'deptid'))
      out_dept[is.na(out_dept)] <- 0
      names(out_dept) <- c('Div', 'Type', 'Title', 'b_1', 'b_2', 'b_3', 'hc', 'a_1', 'a_2', 'a_3')
      print('@@@@@@@@@@@@@@@@@@@@@ 333333333333')
      print(out_dept)

      if (nrow(out_dept) != 0){
        out_dept <- out_dept
      } else{
        out_dept <- df_dept_rate %>%
          mutate(type = 'DEP') %>%
          select(div, type, deptid) %>%
          distinct() %>%
          rename(Div = div,
                 Type = type,
                 Title = deptid) %>%
          mutate(b_1 = 0, 
                 b_2 = 0, 
                 b_3 = 0, 
                 hc = 0, 
                 a_1 = 0, 
                 a_2 = 0, 
                 a_3 = 0)
      }
      
      
      #-----------------------------------#
      #- Combine Department and Function -#
      #-----------------------------------#
      out <- bind_rows(out_dept,
                       out_func)
    }
    return(out)
  }
}
