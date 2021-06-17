# setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
# df_prj <- fread('./data/df_prj.csv', 
#                 na.strings = c('', 'NA', NA, 'NULL'))
# df_hc <- fread('./data/df_HC.csv', 
#                na.strings = c('', 'NA', NA, 'NULL'))
# df <- read_xlsx('./data/powerapp_df_2021-06-10_v1.xlsx', sheet = 1)


hr_dept <- function(df_prj, 
                    df_hc,
                    division = NA){
  
  ## Package names
  packages <- c('data.table', 
                'tidyverse', 
                'dplyr', 
                'readxl',
                'lubridate', 
                'reshape2')
  
  ## Install packages not yet installed
  installed_packages <- packages %in% rownames(installed.packages())
  if (any(installed_packages == FALSE)) {
    install.packages(packages[!installed_packages])
  }
  
  ## Packages loading
  invisible(lapply(packages, library, character.only = TRUE))
  
  ## Do not show warnings
  options(warn = -1) 
  
  
  ###----- Load Data ------------------------------
  df <- read_xlsx('./data/powerapp_df.xlsx', sheet = 1)
  last_update_date = '2021-05-01'
  if (is.na(division) == T){
    df <- df %>%
      filter(ymd(date) <= last_update_date)
  } else {
    df <- df %>%
      filter(div == division,
             ymd(date) <= last_update_date)
  }
  
  ## data from powerApp
  # df_prj <- fread(paste0(input_path, '/df_prj.csv'), 
  #                 na.strings = c('', 'NA', NA, 'NULL'))
  # df_hc <- fread(paste0(input_path, '/df_HC.csv'), 
  #                 na.strings = c('', 'NA', NA, 'NULL'))
    
  ## settings
  # proj_code = unique(df_prj$project_code_old)
  proj_code = '1PD05K550001'
  new_hr = unique(df_prj$execute_hour[df_prj$project_code == df_prj$project_code_old])
  new_mon = unique(df_prj$execute_month[df_prj$project_code == df_prj$project_code_old])
  head_count <- df_hc %>%
    group_by(deptid, sub_job_family) %>%
    summarise(hc = sum(HC)) %>%
    ungroup()
  
  
  ###----- Calculate ------------------------------
  ## Project stage days
  df_proj <- df %>%
    filter(!is.na(project_code),
           project_type == 'old') %>%
    group_by(project_code, customer_name, product_type) %>%
    summarise(C0 = mean(C0_day),
              C1 = mean(C1_day),
              C2 = mean(C2_day),
              C3 = mean(C3_day),
              C4 = mean(C4_day),
              C5 = mean(C5_day),
              C6 = mean(C6_day),
              execute_hour = mean(execute_hour),
              execute_month = mean(execute_month))
  
  #--------------#
  #- Department -#
  #--------------#
  ## Monthly by department
  df_dept_mon <- df %>%
    filter(!is.na(project_code)) %>%
    group_by(div, project_code, customer_name, product_type, deptid, sub_job_family, date) %>%
    summarise(total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
    arrange(desc(date)) %>%
    slice(1) %>%
    slice(rep(1:n(), each = 3)) %>%
    ungroup() %>%
    select(-date) %>%
    mutate(date = rep(seq.Date(max(ymd(df$date)) %m+% months(1), max(ymd(df$date)) %m+% months(3), by = "month"), n()/3)) %>%
    left_join(., df %>%
                group_by(project_code, customer_name, product_type, stage, deptid, sub_job_family) %>%
                summarise(total_hour = sum(total_hour),
                          total_hour_by_dep_func = mean(total_hour_by_dep_func))) %>%
    distinct()
  
  df_dept <- df %>%
    group_by(project_code, customer_name, product_type, stage, deptid, sub_job_family) %>%
    summarise(total_hour = sum(total_hour),
              total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
    filter(project_code == proj_code, stage != 'NA') %>%
    left_join(df_proj %>%
                melt(id = 'project_code', 
                     measure = c('C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6'),
                     variable.name = 'stage',
                     value.name = 'days'),
              by = c('project_code', 'stage')) %>%
    left_join(., df_proj %>%
                select(project_code, execute_hour, execute_month)) %>%
    mutate(hr_ratio = ifelse(is.na(new_hr), 1, new_hr / execute_hour),
           mon_ratio = ifelse(is.na(new_mon), 1, new_mon / execute_month),
           
           total_hour = total_hour * hr_ratio,
           days = days * mon_ratio,
           
           exp_hr = total_hour / days) %>%
    slice(rep(1:n(), days)) %>%
    group_by(sub_job_family) %>%
    mutate(num_d = row_number(),
           gp_num = (num_d - 1) %/% 30 + 1) %>%
    ungroup() %>%
    group_by(project_code, customer_name, product_type, sub_job_family, gp_num) %>%
    summarise(exp_hr = round(sum(exp_hr), 1)) %>%
    ungroup() %>%
    filter(gp_num <= 3) %>%
    mutate(date = case_when(gp_num == 1 ~ max(ymd(df$date)) %m+% months(1),
                            gp_num == 2 ~ max(ymd(df$date)) %m+% months(2),
                            gp_num == 3 ~ max(ymd(df$date)) %m+% months(3),
                            TRUE ~ NA_Date_)) %>%
    select(-gp_num)
  
  ## calculate head count
  df_dept_func_cnt <- df %>%
    group_by(project_code, deptid, sub_job_family) %>%
    summarise(cnt = n()) %>%
    ungroup() %>%
    left_join(., head_count %>% mutate(deptid = as.character(deptid))) %>% 
    replace_na(list(hc = 0)) %>%
    mutate(cnt = cnt + hc) %>%
    group_by(project_code, sub_job_family) %>%
    mutate(tlt_cnt = sum(cnt),
           pct = cnt / tlt_cnt)
  
  cal_dept <- df_dept_mon %>%
    filter(project_code == proj_code) %>%
    select(div, project_code, customer_name, product_type, date, deptid, sub_job_family, total_hour_by_dep_func) %>%
    left_join(., df_dept) %>%
    left_join(., df_dept_func_cnt) %>%
    mutate(HC = cnt,
           exp_total_hr = round((total_hour_by_dep_func + exp_hr) * pct),
           type = 'DEP') %>%
    replace_na(list(exp_total_hr = 0)) %>%
    rename(title = deptid) %>%
    group_by(div, type, title, date) %>%
    summarise(exp_total_hr = sum(exp_total_hr)) %>%
    ungroup() %>%
    spread(date, exp_total_hr)
  
  pre_dept <- df %>%
    filter(ymd(date) <= last_update_date & ymd(date) > as.Date(last_update_date) %m-% months(3)) %>%
    group_by(div, project_code, customer_name, product_type, deptid, date) %>%
    summarise(total_hour_by_dep = mean(total_hour_by_dep)) %>%
    arrange(desc(date)) %>%
    filter(project_code == proj_code) %>%
    mutate(type = 'DEP') %>%
    rename(title = deptid) %>%
    select(div, type, title, date, total_hour_by_dep) %>%
    spread(date, total_hour_by_dep)
  
  if (sum(is.na(pre_dept)) == 0){
    pre_dept <- cal_dept %>%
      select(div, type, title)
    pre_dept[, as.character(as.Date(last_update_date) %m-% months(2))] <- 0
    pre_dept[, as.character(as.Date(last_update_date) %m-% months(1))] <- 0
    pre_dept[, last_update_date] <- 0
  } else {
    pre_dept <- pre_dept
  }
  
  out_dept <- pre_dept %>%
    left_join(., cal_dept)
  
  
  #------------#
  #- Function -#
  #------------#
  ## Monthly by Function
  df_func_mon <- df %>%
    filter(!is.na(project_code)) %>%
    group_by(div, project_code, customer_name, product_type, sub_job_family, date) %>%
    summarise(total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
    arrange(desc(date)) %>%
    slice(1) %>%
    slice(rep(1:n(), each = 3)) %>%
    ungroup() %>%
    select(-date) %>%
    mutate(date = rep(seq.Date(max(ymd(df$date)) %m+% months(1), max(ymd(df$date)) %m+% months(3), by = "month"), n()/3)) %>%
    left_join(df_proj,
              by = c('project_code', 'customer_name', 'product_type')) %>%
    distinct()
  
  df_func <- df %>%
    group_by(div, project_code, customer_name, product_type, stage, sub_job_family) %>%
    summarise(total_hour = sum(total_hour),
              total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
    filter(project_code == proj_code, stage != 'NA') %>%
    left_join(df_proj %>%
                melt(id = 'project_code', 
                     measure = c('C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6'),
                     variable.name = 'stage',
                     value.name = 'days'),
              by = c('project_code', 'stage')) %>%
    left_join(., df_proj %>%
                select(project_code, execute_hour, execute_month)) %>%
    mutate(hr_ratio = ifelse(is.na(new_hr), 1, new_hr / execute_hour),
           mon_ratio = ifelse(is.na(new_mon), 1, new_mon / execute_month),
           
           total_hour = total_hour * hr_ratio,
           days = days * mon_ratio,
           
           exp_hr = total_hour / days) %>%
    slice(rep(1:n(), days)) %>%
    group_by(sub_job_family) %>%
    mutate(num_d = row_number(),
           gp_num = (num_d - 1) %/% 30 + 1) %>%
    ungroup() %>%
    group_by(project_code, customer_name, product_type, sub_job_family, gp_num) %>%
    summarise(exp_hr = round(sum(exp_hr), 1)) %>%
    ungroup() %>%
    filter(gp_num <= 3) %>%
    mutate(date = case_when(gp_num == 1 ~ max(ymd(df$date)) %m+% months(1),
                            gp_num == 2 ~ max(ymd(df$date)) %m+% months(2),
                            gp_num == 3 ~ max(ymd(df$date)) %m+% months(3),
                            TRUE ~ NA_Date_)) %>%
    select(-gp_num)
  
  ## calculate head count
  df_func_cnt <- df %>%
    group_by(project_code, sub_job_family) %>%
    summarise(cnt = n()) %>%
    ungroup() %>%
    left_join(., head_count) %>% 
    replace_na(list(hc = 0)) %>%
    mutate(cnt = cnt + hc)
  
  cal_func <- df_func_mon %>%
    filter(project_code == proj_code) %>%
    select(div, project_code, customer_name, product_type, date, sub_job_family, total_hour_by_dep_func) %>%
    left_join(., df_func) %>%
    mutate(exp_total_hr = round(total_hour_by_dep_func + exp_hr)) %>%
    left_join(., df_func_cnt) %>%
    mutate(HC = cnt,
           type = 'SUB') %>%
    replace_na(list(exp_total_hr = 0)) %>%
    rename(title = sub_job_family) %>%
    group_by(div, type, title, date) %>%
    summarise(exp_total_hr = sum(exp_total_hr)) %>%
    ungroup() %>%
    spread(date, exp_total_hr)
  
  pre_func <- df %>%
    filter(ymd(date) <= last_update_date & ymd(date) > as.Date(last_update_date) %m-% months(3)) %>%
    group_by(div, project_code, customer_name, product_type, sub_job_family, date) %>%
    summarise(total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
    arrange(desc(date)) %>%
    filter(project_code == proj_code) %>%
    mutate(type = 'SUB') %>%
    rename(title = sub_job_family) %>%
    select(div, type, title, date, total_hour_by_dep_func) %>%
    spread(date, total_hour_by_dep_func)
  
  if (sum(is.na(pre_func)) == 0){
    pre_func <- cal_func %>%
      select(div, type, title)
    pre_func[, as.character(as.Date(last_update_date) %m-% months(2))] <- 0
    pre_func[, as.character(as.Date(last_update_date) %m-% months(1))] <- 0
    pre_func[, last_update_date] <- 0
  } else {
    pre_func <- pre_func
  }
  
  out_func <- pre_func %>%
    left_join(., cal_func)
  
  out <- bind_rows(out_dept,
                   out_func)
  return(out)
}
