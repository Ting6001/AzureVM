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

  df_all = dbGetQuery(conn, "select * from [dbo].[UtilizationRateInfo2]")
  dbDisconnect(conn)
  
  if (is.na(division) == F){
    df <- df_all %>%
      filter(div == division)
  }
  
  ## Project stage days (all Divisions)
  df_proj <- df_all %>%
    filter(pmcs_bu_start_dt >= as.character(min(df_all$date))) %>%
    group_by(project_code) %>%
    summarise(C0 = mean(C0_day),
              C1 = mean(C1_day),
              C2 = mean(C2_day),
              C3 = mean(C3_day),
              C4 = mean(C4_day),
              C5 = mean(C5_day),
              C6 = mean(C6_day),
              execute_hour = mean(execute_hour),
              execute_month = mean(execute_month),
              execute_day = max(execute_day),
              project_start_stage = unique(project_start_stage)) %>%
    ungroup()
  
  ## settings
  if (sum(dim(df_hc)) != 0){
    head_count <- df_hc %>%
      mutate(HC = as.numeric(HC),
             ## When sub job family is ('AR', 'HW', 'OP', 'RF', 'RX') replace by 'RH'
             sub_job_family_2 = ifelse(sub_job_family %in% c('AR', 'HW', 'OP', 'RF', 'RX'), 'RH', sub_job_family)) %>%
      group_by(div, deptid, project_code, project_name, sub_job_family_2) %>%
      summarise(hc = sum(HC)) %>%
      ungroup()
  } else{
    head_count <- data.frame(div = NA, deptid = NA, project_code = NA, project_name = NA, sub_job_family_2 = NA, hc = 0)
  }
  
  ## utilization rate by function
  df_func_rate <- df %>%
    group_by(div, sub_job_family_2, date) %>%
    summarise(uti_rate = mean(utilization_rate_by_div_func),
              total_hour_func = sum(total_hour),
              attendance_func = sum(attendance),
              emp_cnt = n_distinct(emplid) - sum(termination_n),
              attendance_emp = mean(attendance)) %>%
    arrange(date) %>%
    ungroup()
  
  ## utilization rate by department
  df_dept_rate <- df %>%
    group_by(div, deptid, sub_job_family_2, date) %>%
    summarise(uti_rate = mean(utilization_rate_by_dep),
              total_hour_dept_func = sum(total_hour),
              emp_cnt = n_distinct(emplid) - sum(termination_n),
              attendance_emp = mean(attendance)) %>%
    arrange(date) %>%
    ungroup()
  
  

  ###----- Calculate ------------------------------
  if (length(setdiff(df_prj$project_code_old, df_proj$project_code)) != 0){
    print('@@@@@@@@@@@@@@ Invalid project code!!')
  } else {
    if (sum(dim(df_prj)) == 0){
      #------------#
      #- Function -#
      #------------#
      ## output table
      out_func <- df_func_rate %>%
        mutate(type = 'SUB') %>%
        select(div, type, sub_job_family_2) %>%
        distinct() %>%
        left_join(df_func_rate %>%
                    filter(date >= max(ymd(df_all$date)) %m-% months(5) & date <= max(ymd(df_all$date)) %m-% months(3)) %>%
                    select(div, sub_job_family_2, date, uti_rate) %>%
                    mutate(uti_rate = round(uti_rate, 2)) %>%
                    spread(date, uti_rate),
                  by = c('div', 'sub_job_family_2')) %>%
        left_join(head_count %>%
                    select(-c(deptid, project_code, project_name)),
                  by = c("div", "sub_job_family_2")) %>%
        left_join(df_func_rate %>%
                    filter(date >= max(ymd(df_all$date)) %m-% months(2) & date <= max(ymd(df_all$date))) %>%
                    left_join(head_count %>%
                                select(-c(deptid, project_code, project_name)),
                              by = c("div", "sub_job_family_2")) %>%
                    mutate(uti_rate_cal = total_hour_func / (attendance_func + attendance_emp * hc)) %>%
                    select(div, sub_job_family_2, date, uti_rate_cal) %>%
                    mutate(uti_rate_cal = round(uti_rate_cal, 2)) %>%
                    spread(date, uti_rate_cal),
                  by = c('div', 'sub_job_family_2'))
      out_func[is.na(out_func)] <- 0
      names(out_func) <- c('Div', 'Type', 'Title', 'b_1', 'b_2', 'b_3', 'hc', 'a_1', 'a_2', 'a_3')
      
      if (nrow(out_func) != 0){
        out_func <- out_func
      } else{
        out_func <- df_func_rate %>%
          mutate(type = 'SUB') %>%
          select(div, type, sub_job_family_2) %>%
          distinct() %>%
          rename(Div = div,
                 Type = type,
                 Title = sub_job_family_2) %>%
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
      ## output table
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
        left_join(head_count %>%
                    group_by(div, deptid) %>%
                    summarise(hc = sum(hc)),
                  by = c("div", "deptid")) %>%
        left_join(df_dept_rate %>%
                    filter(date >= max(ymd(df_all$date)) %m-% months(2) & date <= max(ymd(df_all$date))) %>%
                    left_join(head_count %>%
                                group_by(div, deptid) %>%
                                summarise(hc = sum(hc)),
                              by = c("div", "deptid")) %>%
                    replace_na(list(hc = 0)) %>%
                    mutate(uti_rate_cal = total_hour_dept_func / ((emp_cnt + hc) * attendance_emp)) %>%
                    group_by(div, deptid, date) %>%
                    summarise(uti_rate_cal = round(mean(uti_rate_cal), 2)) %>%
                    ungroup() %>%
                    spread(date, uti_rate_cal),
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
    } else{
      ## Only calculate have both new and old project
      df_prj <- df_prj %>%
        filter(!(is.na(project_code) | project_code == ''),
               !(is.na(project_code_old) | project_code_old == ''))
      
      
      #------------#
      #- Function -#
      #------------#
      ## New and Old project proportion
      proj_prop <- data.frame()
      for (newname in unique(df_prj$project_name)){
        df_tmp <- df_prj %>%
          filter(project_name == newname)
        
        proj_tmp <- data.frame()
        for (pcode in unique(df_tmp$project_code_old)){
          df_tmp1 <- df_tmp %>%
            filter(project_code_old == pcode)
          
          tmp <- df_proj %>%
            filter(project_code == pcode) %>%
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
          filter(project_name == newname) %>%
          select(-deptid)
        
        if (nrow(hc_tmp) == 0){
          hc_tmp <- data.frame(div = NA, project_code = NA, project_name = NA, sub_job_family_2 = NA, hc = NA)
        }
        
        
        proj_prop_tmp <- proj_prop %>%
          filter(project_code == pcode) %>%
          distinct()
        
        rate_tmp <- data.frame()
        for (i in 1:nrow(proj_prop_tmp)){
          df_proj_hour <- df_all %>%
            filter(project_code == proj_prop_tmp$project_code[i]) %>%
            group_by(project_code, sub_job_family_2, date) %>%
            summarise(total_hour_pro_func = sum(total_hour)) %>%
            arrange(date) %>%
            ungroup() %>%
            mutate(add_hour = (total_hour_pro_func * proj_prop_tmp$hr_ratio[i]) / (proj_prop_tmp$mon_ratio[i])) %>%
            group_by(project_code, sub_job_family_2) %>%
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
                      by = c("div", "sub_job_family_2")) %>%
            replace_na(list(hc = 0)) %>%
            left_join(df_proj_hour %>%
                        select(-project_code),
                      by = c('sub_job_family_2', 'date')) %>%
            mutate(emp = emp_cnt + hc) %>%
            distinct() %>%
            group_by(div, date, sub_job_family_2) %>%
            mutate(emp_pct = emp / sum(emp),
                   add_hour_pct = add_hour * emp_pct,
                   add_attendance_emp = (attendance_emp * hc)) %>%
            ungroup()
          
          rate_tmp <- bind_rows(rate_tmp,
                                rate_tmp1) 
        }
        
        
        df_func_rate_future <- bind_rows(df_func_rate_future,
                                         rate_tmp) %>%
          group_by(div, date, sub_job_family_2) %>%
          summarise(total_hour_func = unique(total_hour_func),
                    attendance_func = unique(attendance_func),
                    add_hour_pct = sum(add_hour_pct, na.rm = T),
                    add_attendance_emp = sum(add_attendance_emp, na.rm = T),
                    hc = sum(hc, na.rm = T), 
                    
                    total_hour_by_func_cal = (total_hour_func + add_hour_pct),
                    uti_rate_cal = round(total_hour_by_func_cal / (attendance_func + add_attendance_emp), 2)) %>%
          ungroup()
      }
      
      ## Combine previous and future rate
      out_func <- df_func_rate %>%
        mutate(type = 'SUB') %>%
        select(div, type, sub_job_family_2) %>%
        distinct() %>%
        left_join(df_func_rate %>%
                    filter(date >= max(ymd(df_all$date)) %m-% months(5) & date <= max(ymd(df_all$date)) %m-% months(3)) %>%
                    select(div, sub_job_family_2, date, uti_rate) %>%
                    mutate(uti_rate = round(uti_rate, 2)) %>%
                    spread(date, uti_rate),
                  by = c('div', 'sub_job_family_2')) %>%
        left_join(df_func_rate_future %>%
                    filter(date >= max(ymd(df_all$date)) %m-% months(2) & date <= max(ymd(df_all$date))) %>%
                    select(div, sub_job_family_2, hc, date, uti_rate_cal) %>%
                    mutate(uti_rate_cal = round(uti_rate_cal, 2)) %>%
                    spread(date, uti_rate_cal),
                  by = c('div', 'sub_job_family_2'))
      out_func[is.na(out_func)] <- 0
      names(out_func) <- c('Div', 'Type', 'Title', 'b_1', 'b_2', 'b_3', 'hc', 'a_1', 'a_2', 'a_3')
      
      if (nrow(out_func) != 0){
        out_func <- out_func
      } else{
        out_func <- df_func_rate %>%
          mutate(type = 'SUB') %>%
          select(div, type, sub_job_family_2) %>%
          distinct() %>%
          rename(Div = div,
                 Type = type,
                 Title = sub_job_family_2) %>%
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
      ## Calculate next three month utilization rate
      df_dept_rate_future <- data.frame()
      for (newname in unique(df_prj$project_name)){
        pcode = df_prj$project_code_old[df_prj$project_name == newname]
        
        hc_tmp <- head_count %>%
          filter(project_name == newname)
        if (nrow(hc_tmp) == 0){
          hc_tmp <- data.frame(div = NA, deptid = NA, project_code = NA, project_name = NA, sub_job_family_2 = NA, hc = NA, hc_pct = NA)
        }
        
        
        proj_prop_tmp <- proj_prop %>%
          filter(project_code == pcode) %>%
          distinct()
        
        rate_tmp <- data.frame()
        for (i in 1:nrow(proj_prop_tmp)){
          df_proj_hour <- df_all %>%
            filter(project_code == proj_prop_tmp$project_code[i]) %>%
            group_by(project_code, sub_job_family_2, date) %>%
            summarise(total_hour_pro_func = sum(total_hour)) %>%
            arrange(date) %>%
            ungroup() %>%
            mutate(add_hour = (total_hour_pro_func * proj_prop_tmp$hr_ratio[i]) / (proj_prop_tmp$mon_ratio[i])) %>%
            group_by(project_code, sub_job_family_2) %>%
            mutate(n = row_number()) %>%
            ungroup() %>%
            filter(n <= 3) %>%
            mutate(date = case_when(n == 1 ~ max(ymd(df_all$date)) %m-% months(2),
                                    n == 2 ~ max(ymd(df_all$date)) %m-% months(1),
                                    n == 3 ~ max(ymd(df_all$date)),
                                    TRUE ~ ymd(date))) %>%
            select(-n)
          print('@@@@@@@@@@@@@@@@@@@@@ 111111111111111')
          # print(df_proj_hour)
          # return(df_proj_hour)
          rate_tmp1 <- df_dept_rate %>%
            left_join(hc_tmp,
                      by = c("div", "deptid", "sub_job_family_2")) %>%
            replace_na(list(hc = 0)) %>%
            left_join(df_proj_hour %>%
                        select(-project_code),
                      by = c('sub_job_family_2', 'date')) %>%
            mutate(emp = emp_cnt + hc) %>%
            distinct() %>%
            group_by(div, date, sub_job_family_2) %>%
            mutate(emp_pct = emp / sum(emp)) %>%
            ungroup() %>%
            group_by(div, date, deptid, sub_job_family_2) %>%
            mutate(add_hour_pct = add_hour * emp_pct,
                   add_attendance_emp = (attendance_emp * hc)) %>%
            ungroup() %>%
            distinct()
          
          rate_tmp <- bind_rows(rate_tmp,
                                rate_tmp1)
        }
        print('@@@@@@@@@@@@@@@@@@@@@ 22222222222222222')
        # print(rate_tmp)
        # return(rate_tmp)
        df_dept_rate_future <- bind_rows(df_dept_rate_future,
                                         rate_tmp) %>%
          group_by(div, date, deptid) %>%
          summarise(total_hour_dept_func = sum(unique(total_hour_dept_func)),
                    attendance_emp = sum(unique(attendance_emp)),
                    emp_cnt = sum(unique(emp_cnt)),
                    add_hour_pct = sum(add_hour_pct, na.rm = T),
                    add_attendance_emp = sum(add_attendance_emp, na.rm = T),
                    hc = sum(hc, na.rm = T),
                    
                    total_hour_by_dep_func_cal = (total_hour_dept_func + add_hour_pct),
                    uti_rate_cal_dept = round(total_hour_by_dep_func_cal / (attendance_emp * emp_cnt + add_attendance_emp), 2)) %>%
          ungroup()
      }
      print('@@@@@@@@@@@@@@@@@@@@@ 33333333333333333333')
      # print(df_dept_rate_future)
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
