library(openxlsx)
library(tidyverse)
library(lubridate)

placements <- read.xlsx("Test.xlsx", sheet = 1)

date_format <- as.Date(as.numeric(colnames(placements)[-c(1:2)]), origin = "1899-12-30")

placement_date <- read.xlsx("Test.xlsx", sheet = 2)
placement_date$placement.date <- as.Date(as.numeric(placement_date$placement.date), origin = "1899-12-30")


my_fun <- function(nric, num_months_later){
  p_date = placement_date$placement.date[placement_date$nric == nric]
  uen = placement_date$uen[placement_date$nric == nric]
  later_date <- p_date + months(num_months_later)
  
  ind0 = which(date_format == p_date)
  S0 <- as.numeric(placements[placements$nric == nric & placements$uen == uen, ind0+2])
  if (is_empty(S0)) S0 <- NA
  
  ind = which(date_format == later_date)
  S <- as.numeric(placements[placements$nric == nric & placements$uen == uen, ind+2])
  if (is_empty(S)) S <- NA
  
  placements_filter <- placements$nric == nric
  check <- sum(!is.na(unlist(placements[placements_filter, ind+2]))) > 1
  return(list(S0=S0, S=S, Check=check))
}

long_table <- expand.grid(nric = placement_date$nric, wind = 4:6) %>% 
  mutate(S = map2(nric, wind, my_fun) %>% map("S") %>% unlist(),
         S0 = map2(nric, wind, my_fun) %>% map("S0") %>% unlist(),
         check = map2(nric, wind, my_fun) %>% map("Check") %>% unlist() %>% as.numeric())

table_check <- select(long_table, -S, -S0) %>%
  spread(wind, check) %>%  
  mutate(four_months_later = `4`,
         four_six_months_later = as.numeric((`4`+`5`+`6`)>=1)) %>% 
  select(-c(2:4)) %>% 
  left_join(placement_date, ., by = "nric")

table_growth <- select(long_table, -check) %>% 
  spread(wind, S) %>% 
  mutate(growth_4months = (`4` - S0)/S0,
         growth_6months = (`6` - S0)/S0) %>% 
  select(-c(2:5)) %>% 
  left_join(placement_date, ., by = "nric")


wb = createWorkbook()
addWorksheet(wb, sheetName = "Check")
writeData(wb, sheet = 1, table_check)
addWorksheet(wb, sheetName = "Wage growth")
writeData(wb, sheet = 2, table_growth)
saveWorkbook(wb, "Output.xlsx", overwrite = T)
