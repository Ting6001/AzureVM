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
  df_all <- fread('./data/UtilizationRateInfo_0701.csv',
                  na.strings = c('', 'NA', NA, 'NULL'))
  
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
      group_by(div, deptid, sub_job_family) %>%
      summarise(hc = sum(HC)) %>%
      ungroup() %>%
      group_by(sub_job_family) %>%
      mutate(hc_pct = hc / sum(hc)) %>%
      ungroup()
  } else{
    head_count <- data.frame(div = NA, deptid = NA, sub_job_family = NA, hc = 0, hc_pct = 0)
  }
  
  
  ###----- Calculate ------------------------------
  if (length(setdiff(df_prj$project_code_old, df_proj$project_code)) != 0){
    print('Invalid project code!!')
  } else {
    #------------#
    #- Function -#
    #------------#
    ## Previous utilization rate
    df_func_rate <- df %>%
      group_by(div, sub_job_family, date) %>%
      summarise(uti_rate = mean(utilization_rate_by_div_func),
                total_hour_func = sum(total_hour),
                attendance_func = sum(attendance),
                emp_cnt = n_distinct(emplid) - sum(termination_n),
                attendance_emp = mean(attendance)) %>%
      arrange(date) %>%
      ungroup()
    
    ## New and Old project proportion
    proj_prop <- df_proj %>%
      filter(project_code == unique(df_prj$project_code_old)) %>%
      select(project_code, execute_hour, execute_month, project_start_stage) %>%
      mutate(hr_ratio = ifelse(unique(df_prj$execute_hour) == 0 | is.na(unique(df_prj$execute_hour)), 1, unique(df_prj$execute_hour) / execute_hour),
             mon_ratio = ifelse(unique(df_prj$execute_month) == 0 | unique(df_prj$execute_month), 1, unique(df_prj$execute_month) / execute_month))
    
    ## Calculate next three month utilization rate
    df_func_rate_future <- df_func_rate %>%
      left_join(head_count %>%
                  select(-deptid),
                by = c("div", "sub_job_family")) %>%
      replace_na(list(hc = 0,
                      hc_pct = 1)) %>%
      group_by(div, date, sub_job_family) %>%
      mutate(add_hour = (total_hour_func * proj_prop$hr_ratio) / (proj_prop$mon_ratio),
             total_hour_by_func_cal = (total_hour_func + add_hour) * hc_pct,
             uti_rate_cal = round(total_hour_by_func_cal / (attendance_func + (attendance_emp * hc)), 2)) %>%
      ungroup() %>%
      distinct()
    
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
    df_dept_rate_future <- df_dept_rate %>%
      left_join(head_count,
                by = c("div", "deptid", "sub_job_family")) %>%
      replace_na(list(hc = 0,
                      hc_pct = 1)) %>%
      group_by(div, date, deptid, sub_job_family) %>%
      mutate(add_hour = (total_hour_dept_func * proj_prop$hr_ratio) / (proj_prop$mon_ratio),
             total_hour_by_dep_func_cal = (total_hour_dept_func + add_hour) * hc_pct,
             uti_rate_cal = round(total_hour_by_dep_func_cal / (attendance_emp * emp_cnt + (attendance_emp * hc)), 2)) %>%
      ungroup() %>%
      distinct() %>%
      group_by(div, date, deptid) %>%
      mutate(uti_rate_cal_dept = sum(uti_rate_cal)) %>%
      ungroup() %>%
      distinct()
    
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
}
