hr_cal <- function(df_prj,
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
  df_all <- read_xlsx('./data/powerapp_df.xlsx', sheet = 1)
  
  # last_update_date = '2021-05-01'
  if (is.na(division) == F){
    df <- df_all %>%
      filter(div == division)
  }
  
  ## data from powerApp
  ## Project stage days
  df_proj <- df %>%
    filter(!is.na(project_code)) %>%
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
  
  ## settings
  if (sum(dim(df_prj)) != 0){
    proj_code = unique(df_prj$project_code_old)
    new_hr = unique(df_prj$execute_hour[df_prj$project_code == 0])
    new_mon = unique(df_prj$execute_month[df_prj$project_code == 0])
    if ((df_prj$project_code == 0) == F){
      new_hr = NA_real_
      new_mon = NA_real_
    }
    old_stage <- df_proj$project_start_stage[df_proj$project_code == proj_code]
    new_stage <- 'C0'
  } else{
    new_hr = NA_real_
    new_mon = NA_real_
    old_stage = min(df_proj$project_start_stage)
    new_stage = old_stage
  }
  
  if (sum(dim(df_hc)) != 0){
    head_count <- df_hc %>%
      group_by(div, deptid, sub_job_family) %>%
      summarise(hc = sum(HC)) %>%
      ungroup()
  } else{
    head_count <- data.frame(deptid = NA, sub_job_family = NA, hc = NA)
  }
  
  
  ###----- Calculate ------------------------------
  ## Attendance
  df_dept_attend <- df %>%
    group_by(div, deptid, date) %>%
    summarise(attendance = sum(attendance)) %>%
    ungroup()
  
  df_func_attend <- df %>%
    group_by(div, sub_job_family, date) %>%
    summarise(attendance = sum(attendance)) %>%
    ungroup()
  
  df_attend <- df %>%
    group_by(date) %>%
    summarise(attendance = mean(attendance)) %>%
    ungroup()
  
  
  #--------------#
  #- Department -#
  #--------------#
  ## Monthly by department
  df_dept_mon <- df %>%
    filter(!is.na(project_code)) %>%
    group_by(div, deptid, sub_job_family, project_code, date) %>%
    summarise(total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
    arrange(date) %>%
    ungroup() %>%
    distinct()
  
  if (old_stage > new_stage){
    df_dept <- df %>%
      filter(!is.na(project_code)) %>%
      group_by(div, deptid, sub_job_family, project_code, stage) %>%
      summarise(total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
      ungroup() %>%
      left_join(df_proj %>%
                  melt(id = 'project_code', 
                       measure = c('C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6'),
                       variable.name = 'stage',
                       value.name = 'days'),
                by = c('project_code', 'stage')) %>%
      left_join(df_proj %>%
                  select(project_code, execute_hour, execute_month),
                by = c('project_code')) %>%
      group_by(div, deptid, sub_job_family, stage) %>%
      mutate(total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
      ungroup() %>%
      mutate(hr_ratio = ifelse(new_hr == 0 | is.na(new_hr), 1, new_hr / execute_hour),
             mon_ratio = ifelse(new_mon == 0 | is.na(new_mon), 1, new_mon / execute_month),
             
             total_hour = total_hour_by_dep_func * hr_ratio,
             days = days * mon_ratio,
             
             exp_hr = total_hour_by_dep_func / days) %>%
      replace_na(list(days = 1)) %>%
      slice(rep(1:n(), days)) %>%
      group_by(project_code, deptid, sub_job_family) %>%
      mutate(num_d = row_number(),
             gp_num = (num_d - 1) %/% 30 + 1) %>%
      ungroup() %>%
      group_by(project_code, sub_job_family, gp_num) %>%
      summarise(exp_hr = round(sum(exp_hr), 1)) %>%
      ungroup() %>%
      filter(gp_num <= 3) %>%
      mutate(date = case_when(gp_num == 1 ~ max(ymd(df$date)) %m-% months(2),
                              gp_num == 2 ~ max(ymd(df$date)) %m-% months(1),
                              gp_num == 3 ~ max(ymd(df$date)),
                              TRUE ~ NA_Date_)) %>%
      select(-c(gp_num))
  } else{
    df_dept <- df %>%
      filter(!is.na(project_code)) %>%
      group_by(div, deptid, sub_job_family, project_code, stage) %>%
      summarise(total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
      ungroup() %>%
      left_join(df_proj %>%
                  melt(id = 'project_code', 
                       measure = c('C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6'),
                       variable.name = 'stage',
                       value.name = 'days'),
                by = c('project_code', 'stage')) %>%
      left_join(df_proj %>%
                  select(project_code, execute_hour, execute_month),
                by = c('project_code')) %>%
      mutate(hr_ratio = ifelse(new_hr == 0 | is.na(new_hr), 1, new_hr / execute_hour),
             mon_ratio = ifelse(new_mon == 0 | is.na(new_mon), 1, new_mon / execute_month),
             
             total_hour = total_hour_by_dep_func * hr_ratio,
             days = days * mon_ratio,
             
             exp_hr = total_hour_by_dep_func / days) %>%
      replace_na(list(days = 1)) %>%
      slice(rep(1:n(), days)) %>%
      group_by(project_code, deptid, sub_job_family) %>%
      mutate(num_d = row_number(),
             gp_num = (num_d - 1) %/% 30 + 1) %>%
      ungroup() %>%
      group_by(project_code, sub_job_family, gp_num) %>%
      summarise(exp_hr = round(sum(exp_hr), 1)) %>%
      ungroup() %>%
      filter(gp_num <= 3) %>%
      mutate(date = case_when(gp_num == 1 ~ max(ymd(df$date)) %m-% months(2),
                              gp_num == 2 ~ max(ymd(df$date)) %m-% months(1),
                              gp_num == 3 ~ max(ymd(df$date)),
                              TRUE ~ NA_Date_)) %>%
      select(-c(gp_num))
  }
  
  
  ## calculate head count
  if (sum(dim(df_hc)) != 0){
    df_dept_func_cnt <- df %>%
      group_by(div, deptid, sub_job_family) %>%
      summarise(cnt = n()) %>%
      ungroup() %>%
      left_join(head_count %>% 
                  mutate(deptid = as.character(deptid)),
                by = c('div', 'deptid', 'sub_job_family')) %>% 
      replace_na(list(hc = 0)) %>%
      mutate(cnt = cnt + hc) %>%
      group_by(sub_job_family) %>%
      mutate(tlt_cnt = sum(cnt),
             pct = cnt / tlt_cnt) %>%
      ungroup()
  } else{
    df_dept_func_cnt <- df %>%
      group_by(div, deptid, sub_job_family) %>%
      summarise(cnt = n()) %>%
      ungroup() %>%
      group_by(sub_job_family) %>%
      mutate(tlt_cnt = sum(cnt),
             pct = cnt / tlt_cnt)
  }
  
  
  cal_dept <- df_dept_mon %>%
    select(div, date, deptid, sub_job_family, project_code, total_hour_by_dep_func) %>%
    left_join(df_dept,
              by = c("project_code", "date", "sub_job_family")) %>%
    left_join(df_dept_func_cnt,
              by = c("div", "deptid", "sub_job_family")) %>%
    replace_na(list(exp_hr = 0)) %>%
    mutate(HC = cnt,
           exp_total_hr = round((total_hour_by_dep_func + exp_hr) * pct),
           type = 'DEP') %>%
    replace_na(list(exp_total_hr = 0)) %>%
    rename(title = deptid) %>%
    group_by(div, type, title, date) %>%
    summarise(exp_total_hr = sum(exp_total_hr)) %>%
    ungroup() %>%
    left_join(df_dept_attend %>%
                rename(dept_att = attendance),
              by = c('div', 'title' = 'deptid', 'date')) %>%
    left_join(head_count %>% 
                group_by(deptid) %>%
                summarise(hc = sum(hc)),
              by = c('title' = 'deptid')) %>%
    left_join(df_attend,
              by = c('date')) %>%
    replace_na(list(dept_att = 0,
                    hc = 0,
                    attendance = 0)) %>%
    mutate(uti_rate = exp_total_hr / (dept_att + attendance * hc)) %>%
    mutate(uti_rate = ifelse(is.infinite(uti_rate) | is.na(uti_rate), 0, uti_rate)) %>%
    select(div, type, title, date, hc, uti_rate) %>%
    spread(date, uti_rate)
  
  out_names <- c("div", "type", "title", 
                 as.character(max(ymd(df_all$date)) %m-% months(5)),
                 as.character(max(ymd(df_all$date)) %m-% months(4)),
                 as.character(max(ymd(df_all$date)) %m-% months(3)),
                 "hc", 
                 as.character(max(ymd(df_all$date)) %m-% months(2)),
                 as.character(max(ymd(df_all$date)) %m-% months(1)),
                 as.character(max(ymd(df_all$date))))
 
  if (length(setdiff(out_names, names(cal_dept))) == 0){
    out_dept <- cal_dept %>% 
      select(out_names)
  } else{
    out_dept1 <- setNames(data.frame(matrix(ncol = 10, nrow = 1)), 
                         out_names) %>%
      mutate_all(as.character())
    
    out_dept <- bind_rows(out_dept1, 
              cal_dept %>% 
                select(div, type, title, hc)) %>%
      filter(!is.na(div))
    out_dept[is.na(out_dept)] <- 0
  }
  names(out_dept) <- c('Div', 'Type', 'Title', 'b_1', 'b_2', 'b_3', 'hc', 'a_1', 'a_2', 'a_3')
  
  
  #------------#
  #- Function -#
  #------------#
  ## Monthly by Function
  df_func_mon <- df %>%
    filter(!is.na(project_code)) %>%
    group_by(div, project_code, sub_job_family, date) %>%
    summarise(total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
    arrange(date) %>%
    ungroup() %>%
    distinct()
  
  if (old_stage > new_stage){
    df_func <- df %>%
      filter(!is.na(project_code)) %>%
      group_by(div, project_code, stage, sub_job_family) %>%
      summarise(total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
      left_join(df_proj %>%
                  melt(id = 'project_code', 
                       measure = c('C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6'),
                       variable.name = 'stage',
                       value.name = 'days'),
                by = c('project_code', 'stage')) %>%
      left_join(df_proj %>%
                  select(project_code, execute_hour, execute_month),
                by = "project_code") %>%
      group_by(div, sub_job_family, stage) %>%
      mutate(total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
      ungroup() %>%
      mutate(hr_ratio = ifelse(new_hr == 0 | is.na(new_hr), 1, new_hr / execute_hour),
             mon_ratio = ifelse(new_mon == 0 | is.na(new_mon), 1, new_mon / execute_month),
             
             total_hour = total_hour_by_dep_func * hr_ratio,
             days = days * mon_ratio,
             
             exp_hr = total_hour_by_dep_func / days) %>%
      replace_na(list(days = 1)) %>%
      slice(rep(1:n(), days)) %>%
      group_by(project_code, sub_job_family) %>%
      mutate(num_d = row_number(),
             gp_num = (num_d - 1) %/% 30 + 1) %>%
      ungroup() %>%
      group_by(project_code, sub_job_family, gp_num) %>%
      summarise(exp_hr = round(sum(exp_hr), 1)) %>%
      ungroup() %>%
      filter(gp_num <= 3) %>%
      mutate(date = case_when(gp_num == 1 ~ max(ymd(df$date)) %m-% months(2),
                              gp_num == 2 ~ max(ymd(df$date)) %m-% months(1),
                              gp_num == 3 ~ max(ymd(df$date)),
                              TRUE ~ NA_Date_)) %>%
      select(-gp_num)
  } else{
    df_func <- df %>%
      filter(!is.na(project_code)) %>%
      group_by(div, project_code, stage, sub_job_family) %>%
      summarise(total_hour_by_dep_func = mean(total_hour_by_dep_func)) %>%
      left_join(df_proj %>%
                  melt(id = 'project_code', 
                       measure = c('C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6'),
                       variable.name = 'stage',
                       value.name = 'days'),
                by = c('project_code', 'stage')) %>%
      left_join(df_proj %>%
                  select(project_code, execute_hour, execute_month),
                by = "project_code") %>%
      mutate(hr_ratio = ifelse(new_hr == 0 | is.na(new_hr), 1, new_hr / execute_hour),
             mon_ratio = ifelse(new_mon == 0 | is.na(new_mon), 1, new_mon / execute_month),
             
             total_hour = total_hour_by_dep_func * hr_ratio,
             days = days * mon_ratio,
             
             exp_hr = total_hour_by_dep_func / days) %>%
      replace_na(list(days = 1)) %>%
      slice(rep(1:n(), days)) %>%
      group_by(project_code, sub_job_family) %>%
      mutate(num_d = row_number(),
             gp_num = (num_d - 1) %/% 30 + 1) %>%
      ungroup() %>%
      group_by(project_code, sub_job_family, gp_num) %>%
      summarise(exp_hr = round(sum(exp_hr), 1)) %>%
      ungroup() %>%
      filter(gp_num <= 3) %>%
      mutate(date = case_when(gp_num == 1 ~ max(ymd(df$date)) %m-% months(2),
                              gp_num == 2 ~ max(ymd(df$date)) %m-% months(1),
                              gp_num == 3 ~ max(ymd(df$date)),
                              TRUE ~ NA_Date_)) %>%
      select(-gp_num)
  }
  
  
  ## calculate head count
  if (sum(dim(df_hc)) != 0){
    df_func_cnt <- df %>%
      group_by(sub_job_family) %>%
      summarise(cnt = n()) %>%
      ungroup() %>%
      left_join(head_count %>%
                  group_by(sub_job_family) %>%
                  summarise(hc = sum(hc)),
                by = "sub_job_family") %>% 
      replace_na(list(hc = 0)) %>%
      mutate(cnt = cnt + hc)
  } else{
    df_func_cnt <- df %>%
      group_by(sub_job_family) %>%
      summarise(cnt = n()) %>%
      ungroup()
  }
  
  
  cal_func <- df_func_mon %>%
    filter(!(is.na(project_code))) %>%
    select(div, project_code, date, sub_job_family, total_hour_by_dep_func) %>%
    left_join(df_func,
              by = c("project_code", "date", "sub_job_family")) %>%
    replace_na(list(exp_hr = 0)) %>%
    mutate(exp_total_hr = round(total_hour_by_dep_func + exp_hr)) %>%
    left_join(df_func_cnt,
              by = c("sub_job_family")) %>%
    mutate(HC = cnt,
           type = 'SUB') %>%
    replace_na(list(exp_total_hr = 0)) %>%
    rename(title = sub_job_family) %>%
    group_by(div, type, title, date) %>%
    summarise(exp_total_hr = sum(exp_total_hr)) %>%
    ungroup() %>%
    left_join(df_func_attend %>%
                rename(func_att = attendance),
              by = c('div', 'title' = 'sub_job_family', 'date')) %>%
    left_join(head_count %>% 
                group_by(sub_job_family) %>%
                summarise(hc = sum(hc)),
              by = c('title' = 'sub_job_family')) %>%
    left_join(df_attend,
              by = c('date')) %>%
    replace_na(list(func_att = 0,
                    hc = 0,
                    attendance = 0)) %>%
    mutate(uti_rate = exp_total_hr / (func_att + attendance * hc)) %>%
    select(div, type, title, date, hc, uti_rate) %>%
    spread(date, uti_rate)
  
  
  out_names <- c("div", "type", "title", 
                 as.character(max(ymd(df_all$date)) %m-% months(5)),
                 as.character(max(ymd(df_all$date)) %m-% months(4)),
                 as.character(max(ymd(df_all$date)) %m-% months(3)),
                 "hc", 
                 as.character(max(ymd(df_all$date)) %m-% months(2)),
                 as.character(max(ymd(df_all$date)) %m-% months(1)),
                 as.character(max(ymd(df_all$date))))
  
  if (length(setdiff(out_names, names(cal_func))) == 0){
    out_func <- cal_func %>% select(out_names)
  } else{
    out_func1 <- setNames(data.frame(matrix(ncol = 10, nrow = 1)), 
                         out_names) %>%
      mutate_all(as.character())
    
    out_func <- bind_rows(out_func1, 
                          cal_func %>% 
                            select(div, type, title, hc)) %>%
      filter(!is.na(div))
    out_func[is.na(out_func)] <- 0
  }
  names(out_func) <- c('Div', 'Type', 'Title', 'b_1', 'b_2', 'b_3', 'hc', 'a_1', 'a_2', 'a_3')
  
  out <- bind_rows(out_dept,
                   out_func)
  
  return(out)
}
