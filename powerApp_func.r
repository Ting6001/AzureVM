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
  
  ## Do not show warnings & Suppress summarise info
  options(warn = -1,
          dplyr.summarise.inform = FALSE)
  
  ###----- Load Data ------------------------------
  # setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
  # df_prj <- fread('./data/df_prj_0621.csv', 
  #                 na.strings = c('', 'NA', NA, 'NULL'))
  # df_hc <- fread('./data/df_HC_0624.csv', 
  #                na.strings = c('', 'NA', NA, 'NULL'))
  df <- read_xlsx('./data/powerapp_df.xlsx', sheet = 1)
  
  # last_update_date = '2021-05-01'
  if (is.na(division) == F){
    df <- df %>%
      filter(div == division)
  }
  
  ## data from powerApp
  ## settings
  proj_code = unique(df_prj$project_code_old)
  new_hr = unique(df_prj$execute_hour[df_prj$project_code == 0])
  new_mon = unique(df_prj$execute_month[df_prj$project_code == 0])
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
              execute_month = mean(execute_month),
              execute_day = max(execute_day),
              project_start_stage = unique(project_start_stage)) %>%
    ungroup()
  old_stage <- df_proj$project_start_stage[df_proj$project_code == proj_code]
  new_stage <- 'C0'
  
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
    mutate(date = rep(seq.Date(max(ymd(df$date)) %m-% months(2), max(ymd(df$date)), by = "month"), n()/3)) %>%
    left_join(df %>%
                group_by(project_code, customer_name, product_type, stage, deptid, sub_job_family) %>%
                summarise(total_hour = sum(total_hour),
                          total_hour_by_dep_func = mean(total_hour_by_dep_func)),
              by = c("project_code", "customer_name", "product_type", "deptid", "sub_job_family", "total_hour_by_dep_func")) %>%
    distinct()
  
  df_dept <- df %>%
    group_by(project_code, product_type, stage, deptid, sub_job_family) %>%
    summarise(total_hour = sum(total_hour),
              total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
    filter(!is.na(project_code)) %>%
    # filter(project_code == proj_code) %>%
    left_join(df_proj %>%
                melt(id = 'project_code', 
                     measure = c('C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6'),
                     variable.name = 'stage',
                     value.name = 'days'),
              by = c('project_code', 'stage')) %>%
    left_join(df_proj %>%
                select(project_code, execute_hour, execute_month),
              by = c("project_code")) %>%
    mutate(hr_ratio = ifelse(new_hr == 0, 1, new_hr / execute_hour),
           mon_ratio = ifelse(new_mon == 0, 1, new_mon / execute_month),
           
           total_hour = total_hour * hr_ratio,
           days = days * mon_ratio,
           
           exp_hr = total_hour / days) %>%
    replace_na(list(days = 1)) %>%
    slice(rep(1:n(), days)) %>%
    group_by(project_code, deptid, sub_job_family) %>%
    mutate(num_d = row_number(),
           gp_num = (num_d - 1) %/% 30 + 1) %>%
    ungroup() %>%
    group_by(project_code, product_type, sub_job_family, gp_num) %>%
    summarise(exp_hr = round(sum(exp_hr), 1)) %>%
    ungroup() %>%
    filter(gp_num <= 3) %>%
    mutate(date = case_when(gp_num == 1 ~ max(ymd(df$date)) %m-% months(2),
                            gp_num == 2 ~ max(ymd(df$date)) %m-% months(1),
                            gp_num == 3 ~ max(ymd(df$date)),
                            TRUE ~ NA_Date_)) %>%
    select(-c(gp_num))
  
  ## calculate head count
  df_dept_func_cnt <- df %>%
    group_by(project_code, deptid, sub_job_family) %>%
    summarise(cnt = n()) %>%
    ungroup() %>%
    left_join(head_count %>% mutate(deptid = as.character(deptid)),
              by = c("deptid", "sub_job_family")) %>% 
    replace_na(list(hc = 0)) %>%
    mutate(cnt = cnt + hc) %>%
    group_by(project_code, sub_job_family) %>%
    mutate(tlt_cnt = sum(cnt),
           pct = cnt / tlt_cnt)
  
  cal_dept <- df_dept_mon %>%
    # filter(project_code == proj_code) %>%
    select(div, project_code, customer_name, product_type, date, deptid, sub_job_family, total_hour_by_dep_func) %>%
    left_join(df_dept,
              by = c("project_code", "product_type", "date", "sub_job_family")) %>%
    left_join(df_dept_func_cnt,
              by = c("project_code", "deptid", "sub_job_family")) %>%
    mutate(HC = cnt,
           exp_total_hr = round((total_hour_by_dep_func + exp_hr) * pct),
           type = 'DEP') %>%
    replace_na(list(exp_total_hr = 0)) %>%
    rename(title = deptid) %>%
    group_by(div, type, title, date) %>%
    summarise(exp_total_hr = sum(exp_total_hr)) %>%
    ungroup() %>%
    spread(date, exp_total_hr)
  
  
  if (old_stage > new_stage){
    pre_dept <- df %>%
      filter(stage != 'NA') %>%
      mutate(type = 'DEP') %>%
      rename(title = deptid) %>%
      group_by(type, div, title, stage) %>%
      summarise(total_hour_by_dep = mean(total_hour_by_dep)) %>%
      ungroup() %>%
      group_by(type, div, title) %>%
      mutate(n = row_number()) %>%
      filter(n <= 3) %>%
      select(div, type, title, n, total_hour_by_dep) %>%
      spread(n, total_hour_by_dep)
  } else{
    pre_dept <- df %>%
      arrange(desc(date)) %>%
      mutate(type = 'DEP') %>%
      rename(title = deptid) %>%
      group_by(type, div, project_code, customer_name, product_type, title, date) %>%
      summarise(total_hour_by_dep = mean(total_hour_by_dep)) %>%
      ungroup() %>%
      group_by(type, div, title) %>%
      mutate(gp_num = row_number()) %>%
      filter(gp_num <= 3) %>%
      select(div, type, title, gp_num, total_hour_by_dep) %>%
      spread(gp_num, total_hour_by_dep)
  }
  
  out_dept <- pre_dept %>%
    left_join(head_count %>% 
                group_by(deptid) %>%
                summarise(hc = sum(hc)),
              by = c('title' = 'deptid')) %>%
    left_join(cal_dept,
              by = c("div", "type", "title"))
  names(out_dept) <- c('Div', 'Type', 'Title', 'b_1', 'b_2', 'b_3', 'hc', 'a_1', 'a_2', 'a_3')
  
  
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
    mutate(date = rep(seq.Date(max(ymd(df$date)) %m-% months(2), max(ymd(df$date)), by = "month"), n()/3)) %>%
    left_join(df_proj,
              by = c('project_code', 'customer_name', 'product_type')) %>%
    distinct()
  
  df_func <- df %>%
    group_by(div, project_code, customer_name, product_type, stage, sub_job_family) %>%
    summarise(total_hour = sum(total_hour),
              total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
    filter(!is.na(project_code)) %>%
    # filter(project_code == proj_code, stage != 'NA') %>%
    left_join(df_proj %>%
                melt(id = 'project_code', 
                     measure = c('C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6'),
                     variable.name = 'stage',
                     value.name = 'days'),
              by = c('project_code', 'stage')) %>%
    left_join(df_proj %>%
                select(project_code, execute_hour, execute_month),
              by = "project_code") %>%
    mutate(hr_ratio = ifelse(new_hr == 0, 1, new_hr / execute_hour),
           mon_ratio = ifelse(new_mon == 0, 1, new_mon / execute_month),
           
           total_hour = total_hour * hr_ratio,
           days = days * mon_ratio,
           
           exp_hr = total_hour / days) %>%
    replace_na(list(days = 1)) %>%
    slice(rep(1:n(), days)) %>%
    group_by(project_code, sub_job_family) %>%
    mutate(num_d = row_number(),
           gp_num = (num_d - 1) %/% 30 + 1) %>%
    ungroup() %>%
    group_by(project_code, customer_name, product_type, sub_job_family, gp_num) %>%
    summarise(exp_hr = round(sum(exp_hr), 1)) %>%
    ungroup() %>%
    filter(gp_num <= 3) %>%
    mutate(date = case_when(gp_num == 1 ~ max(ymd(df$date)) %m-% months(2),
                            gp_num == 2 ~ max(ymd(df$date)) %m-% months(1),
                            gp_num == 3 ~ max(ymd(df$date)),
                            TRUE ~ NA_Date_)) %>%
    select(-gp_num)
  
  ## calculate head count
  df_func_cnt <- df %>%
    group_by(project_code, sub_job_family) %>%
    summarise(cnt = n()) %>%
    ungroup() %>%
    left_join(head_count,
              by = "sub_job_family") %>% 
    replace_na(list(hc = 0)) %>%
    mutate(cnt = cnt + hc)
  
  cal_func <- df_func_mon %>%
    # filter(project_code == proj_code) %>%
    select(div, project_code, customer_name, product_type, date, sub_job_family, total_hour_by_dep_func) %>%
    left_join(df_func,
              by = c("project_code", "customer_name", "product_type", "date", "sub_job_family")) %>%
    mutate(exp_total_hr = round(total_hour_by_dep_func + exp_hr)) %>%
    left_join(df_func_cnt,
              by = c("project_code", "sub_job_family")) %>%
    mutate(HC = cnt,
           type = 'SUB') %>%
    replace_na(list(exp_total_hr = 0)) %>%
    rename(title = sub_job_family) %>%
    group_by(div, type, title, date) %>%
    summarise(exp_total_hr = sum(exp_total_hr)) %>%
    ungroup() %>%
    spread(date, exp_total_hr)
  
  
  if (old_stage > new_stage){
    pre_func <- df %>%
      filter(stage != 'NA') %>%
      mutate(type = 'SUB') %>%
      rename(title = sub_job_family) %>%
      group_by(type, div, title, stage) %>%
      summarise(total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
      ungroup() %>%
      group_by(type, div, title) %>%
      mutate(n = row_number()) %>%
      filter(n <= 3) %>%
      select(div, type, title, n, total_hour_by_dep_func) %>%
      spread(n, total_hour_by_dep_func)
  } else{
    pre_func <- df %>%
      arrange(desc(date)) %>%
      mutate(type = 'SUB') %>%
      rename(title = sub_job_family) %>%
      group_by(type, div, title, date) %>%
      summarise(total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
      ungroup() %>%
      group_by(type, div, title) %>%
      mutate(gp_num = row_number()) %>%
      filter(gp_num <= 3) %>%
      select(div, type, title, gp_num, total_hour_by_dep_func) %>%
      spread(gp_num, total_hour_by_dep_func)
  }
  
  out_func <- pre_func %>%
    left_join(head_count %>% select(sub_job_family, hc),
              by = c('title' = 'sub_job_family')) %>%
    left_join(cal_func,
              by = c("div", "type", "title")) 
  names(out_func) <- c('Div', 'Type', 'Title', 'b_1', 'b_2', 'b_3', 'hc', 'a_1', 'a_2', 'a_3')
  
  out <- bind_rows(out_dept,
                   out_func)
  return(out)
}
