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
  df_all <- fread('./data/UtilizationRateInfo_0630.csv',
                  na.strings = c('', 'NA', NA, 'NULL'))
  
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
    new_hr = unique(df_prj$execute_hour)
    new_mon = unique(df_prj$execute_month)
    old_stage <- ifelse(proj_code == "" | is.na(proj_code), min(df_proj$project_start_stage), df_proj$project_start_stage[df_proj$project_code == proj_code])
    new_stage <- 'C0'
  } else{
    new_hr = NA_real_
    new_mon = NA_real_
    old_stage = min(df_proj$project_start_stage)
    new_stage = old_stage
  }
  
  old_stage = ifelse(old_stage == 'C0' | is.na(old_stage), 0,
                     ifelse(old_stage == 'C1', 1,
                            ifelse(old_stage == 'C2', 2,
                                   ifelse(old_stage == 'C3', 3,
                                          ifelse(old_stage == 'C4', 4,
                                                 ifelse(old_stage == 'C5', 5,
                                                        ifelse(old_stage == 'C6', 6)))))))
  new_stage = ifelse(new_stage == 'C0'| is.na(new_stage), 0,
                     ifelse(new_stage == 'C1', 1,
                            ifelse(new_stage == 'C2', 2,
                                   ifelse(new_stage == 'C3', 3,
                                          ifelse(new_stage == 'C4', 4,
                                                 ifelse(new_stage == 'C5', 5,
                                                        ifelse(new_stage == 'C6', 6)))))))
  
  if (sum(dim(df_hc)) != 0){
    head_count <- df_hc %>%
      group_by(div, deptid, sub_job_family) %>%
      summarise(hc = sum(HC)) %>%
      ungroup() %>%
      group_by(sub_job_family) %>%
      mutate(hc_pct = hc / sum(hc)) %>%
      ungroup()
  } else{
    head_count <- data.frame(deptid = NA, sub_job_family = NA, hc = 0, hc_pct = 0)
  }
  
  
  ###----- Calculate ------------------------------
  ## setting standing month
  stand_date = '2021-05-01'
  
  ## Attendance by department and function
  df_attend <- df %>%
    group_by(div, date) %>%
    summarise(attendance_emp = mean(attendance)) %>%
    ungroup() %>%
    left_join(df %>%
                group_by(div, deptid, sub_job_family, date) %>%
                summarise(cnt_dept_func = n_distinct(emplid)) %>%
                ungroup(),
              by = c("div", "date")) %>%
    left_join(df %>%
                group_by(div, sub_job_family, date) %>%
                summarise(cnt_func = n_distinct(emplid)) %>%
                ungroup(),
              by = c('div', 'date', 'sub_job_family')) %>%
    mutate(attendance_dept_func = attendance_emp * cnt_dept_func,
           attendance_func = attendance_emp * cnt_func)
  
  
  #------------#
  #- Function -#
  #------------#
  ## Monthly hours by Function
  df_func_mon <- df %>%
    filter(!is.na(project_code)) %>%
    group_by(div, sub_job_family, date) %>%
    summarise(total_hour_by_func = sum(total_hour)) %>%
    arrange(date) %>%
    ungroup()
  
  ## Previous utilization rate
  df_func_rate <- df %>%
    filter(!is.na(project_code)) %>%
    group_by(div, sub_job_family, date) %>%
    summarise(uti_rate = mean(utilization_rate_by_div_func)) %>%
    arrange(date) %>%
    ungroup()
  
  
  ## New and Old project proportion
  proj_prop <- df_proj %>%
    filter(project_code == proj_code) %>%
    select(project_code, execute_hour, execute_month, project_start_stage) %>%
    mutate(new_hour = new_hr,
           new_month = new_mon) %>%
    mutate(hr_ratio = ifelse(new_hr == 0 | is.na(new_hr), 1, new_hr / execute_hour),
           mon_ratio = ifelse(new_mon == 0 | is.na(new_mon), 1, new_mon / execute_month))
  
  
  # ### Average hour by stage
  # df_func_stage_mon <- df %>%
  #   filter(!is.na(project_code)) %>%
  #   group_by(div, sub_job_family, stage) %>%
  #   summarise(total_hour_by_func_stage = sum(total_hour)) %>%
  #   ungroup() %>%
  #   group_by(sub_job_family, stage) %>%
  #   mutate(avg_totle_hour_stage = mean(total_hour_by_func_stage)) %>%
  #   ungroup() 
  
  
  ## Calculate next three month utilization rate
  df_func_rate_future <- df_proj %>%
    filter(project_code == proj_code) %>%
    melt(id = 'project_code', 
         measure = c('C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6'),
         variable.name = 'stage',
         value.name = 'days') %>%
    replace_na(list(days = 1)) %>%
    slice(rep(1:n(), days)) %>%
    mutate(months = round(days / 30, 1),
           gp_num = (row_number() - 1) %/% 30 + 1) %>%
    filter(gp_num <= 3) %>%
    mutate(date = case_when(gp_num == 1 ~ as.Date(stand_date) %m+% months(1),
                            gp_num == 2 ~ as.Date(stand_date) %m+% months(2),
                            gp_num == 3 ~ as.Date(stand_date) %m+% months(3),
                            TRUE ~ NA_Date_)) %>%
    left_join(df_func_mon %>%
                left_join(df_attend %>%
                            select(div, date, attendance_emp, sub_job_family, attendance_func),
                          by = c('div', 'sub_job_family', 'date')) %>%
                distinct() %>% 
                filter(date >= as.Date(stand_date) %m+% months(1)),
              by = c('date')) %>%
    left_join(head_count %>%
                select(-deptid),
              by = c("div", "sub_job_family")) %>%
    replace_na(list(hc = 0,
                    hc_pct = 1)) %>%
    group_by(div, date, sub_job_family) %>%
    mutate(cal_hour = (total_hour_by_func * proj_prop$hr_ratio) / (months * proj_prop$mon_ratio),
           total_hour_by_func_cal = (total_hour_by_func + cal_hour) * hc_pct) %>%
    ungroup() %>%
    distinct() %>%
    group_by(div, date, sub_job_family) %>%
    summarise(total_hour_by_func_cal = sum(total_hour_by_func_cal),
              hc = mean(hc),
              uti_rate = round(sum(total_hour_by_func_cal) / (attendance_func + (attendance_emp * hc)), 2)) %>%
    ungroup() %>%
    distinct()
  
  
  ## Combine previous and future rate
  out_func <- df_func_mon %>%
    mutate(type = 'SUB') %>%
    select(div, type, sub_job_family) %>%
    distinct() %>%
    left_join(df_func_rate_future %>%
                select(div, sub_job_family, date, uti_rate) %>%
                spread(date, uti_rate),
              by = c('div', 'sub_job_family')) %>%
    left_join(head_count %>%
                group_by(div, sub_job_family) %>%
                summarise(hc = sum(hc)) %>%
                ungroup(),
              by = c('div', 'sub_job_family')) %>%
    left_join(df_func_rate %>%
                filter(date >= as.Date(stand_date) %m-% months(2) & date <= as.Date(stand_date)) %>%
                select(div, sub_job_family, date, uti_rate) %>%
                spread(date, uti_rate),
              by = c('div', 'sub_job_family'))
  out_func[is.na(out_func)] <- 0
  names(out_func) <- c('Div', 'Type', 'Title', 'b_1', 'b_2', 'b_3', 'hc', 'a_1', 'a_2', 'a_3')
  
  
  if (nrow(out_func) != 0){
    out_func <- out_func
  } else{
    out_func <- df_func_mon %>%
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
  ## Monthly hours by Function
  df_dept_mon <- df %>%
    filter(!is.na(project_code)) %>%
    group_by(div, deptid, sub_job_family, date) %>%
    summarise(total_hour_by_dep_func = sum(total_hour)) %>%
    arrange(date) %>%
    ungroup()
  
  ## Previous utilization rate
  df_dept_rate <- df %>%
    filter(!is.na(project_code)) %>%
    group_by(div, deptid, date) %>%
    summarise(uti_rate = mean(utilization_rate_by_dep)) %>%
    arrange(date) %>%
    ungroup()
  
  ## Calculate next three month utilization rate
  df_dept_rate_future <- df_proj %>%
    filter(project_code == proj_code) %>%
    melt(id = 'project_code', 
         measure = c('C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6'),
         variable.name = 'stage',
         value.name = 'days') %>%
    replace_na(list(days = 1)) %>%
    slice(rep(1:n(), days)) %>%
    mutate(months = round(days / 30, 1),
           gp_num = (row_number() - 1) %/% 30 + 1) %>%
    filter(gp_num <= 3) %>%
    mutate(date = case_when(gp_num == 1 ~ as.Date(stand_date) %m+% months(1),
                            gp_num == 2 ~ as.Date(stand_date) %m+% months(2),
                            gp_num == 3 ~ as.Date(stand_date) %m+% months(3),
                            TRUE ~ NA_Date_)) %>%
    left_join(df_dept_mon %>%
                left_join(df_attend %>%
                            select(div, date, attendance_emp, deptid, sub_job_family, attendance_dept_func),
                          by = c('div', 'deptid', 'sub_job_family', 'date')) %>%
                distinct() %>% 
                filter(date >= as.Date(stand_date) %m+% months(1)),
              by = c('date')) %>%
    left_join(head_count,
              by = c("div", "deptid", "sub_job_family")) %>%
    replace_na(list(hc = 0,
                    hc_pct = 1)) %>%
    group_by(div, date, deptid, sub_job_family) %>%
    mutate(cal_hour = (total_hour_by_dep_func * proj_prop$hr_ratio) / (months * proj_prop$mon_ratio),
           total_hour_by_dep_func_cal = (total_hour_by_dep_func + cal_hour) * hc_pct) %>%
    ungroup() %>%
    distinct() %>%
    group_by(div, date, deptid) %>%
    summarise(total_hour_by_dep_cal = sum(total_hour_by_dep_func_cal),
              attendance_dept = sum(attendance_dept_func), 
              hc = mean(hc),
              uti_rate = round(total_hour_by_dep_cal / (attendance_dept + (attendance_emp * hc)), 2)) %>%
    ungroup() %>%
    distinct()
  
  
  ## Combine previous and future rate
  out_dept <-  df_dept_mon %>%
    mutate(type = 'DEP') %>%
    select(div, type, deptid) %>%
    distinct() %>%
    left_join(df_dept_rate_future %>%
                select(div, deptid, date, uti_rate) %>%
                spread(date, uti_rate),
              by = c('div', 'deptid')) %>%
    left_join(head_count %>%
                group_by(div, deptid) %>%
                summarise(hc = sum(hc)) %>%
                ungroup(),
              by = c('div', 'deptid')) %>%
    left_join(df_dept_rate %>%
                filter(date >= as.Date(stand_date) %m-% months(2) & date <= as.Date(stand_date)) %>%
                select(div, deptid, date, uti_rate) %>%
                spread(date, uti_rate),
              by = c('div', 'deptid'))
  out_dept[is.na(out_dept)] <- 0
  names(out_dept) <- c('Div', 'Type', 'Title', 'b_1', 'b_2', 'b_3', 'hc', 'a_1', 'a_2', 'a_3')
  
  if (nrow(out_dept) != 0){
    out_dept <- out_dept
  } else{
    out_dept <- df_dept_mon %>%
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
  
  return(out)
}
