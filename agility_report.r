## This reads in two Agility reports which are in XLSX Worksheet Format
## It then runs some algorithms to the current state metrics and report variations
## These variations show how the blueprints are changing over time (adds and deletes)

## (C) Paul Colmer 2016

## 4 hrs

## Must have Java installed to work as XLConnect and XLConnectJars are written in Java
## Installed Java V8 Update 71

## Setup Libraries
##install.packages("rJava")
library(rJava)

##install.packages("XLConnect")
library(XLConnect)

##install.packages("XLConnectJars")
library(XLConnectJars)

##install.packages("sqldf")
library(sqldf)

##Setup Working Directory
setwd("C:/Users/pcolmer/OneDrive for Business/R Studio/Agility Report")

## Check to see the correct baseline file exists
x <- file.exists("agility_baseline_report_03Feb2016.xls")
if(x == FALSE) {
    print("Agility Baseline Report Filename INVALID")
    stop()
} else {
    print("Agility Baseline Report GOOD")
    print("Press [enter] to continue")
    line <- readline()
}

## Load in the Baseline Report
wb <- loadWorkbook("agility_baseline_report_03Feb2016.xls")
## Load in the Sheets
bp_index_sheet <- readWorksheet(wb, sheet=2)
bp_wklds_sheet <- readWorksheet(wb, sheet=3)
bp_wklds_resources_sheet <- readWorksheet(wb, sheet=4)
bp_vars_all_sheet <- readWorksheet(wb, sheet=5)
bp_pkg_scrpt_ver_sheet <- readWorksheet(wb, sheet=6)

## Clean Up bp_index_sheet and setup column headings
## bp_index_sheet[1,x] = column headings
bp_index_sheet <- bp_index_sheet[-1, ]
colnames(bp_index_sheet) <- c(bp_index_sheet[1,1], bp_index_sheet[1,2], bp_index_sheet[1,3], 
                              bp_index_sheet[1,4], bp_index_sheet[1,5], bp_index_sheet[1,6], 
                              bp_index_sheet[1,7], bp_index_sheet[1,8])
bp_index_sheet <- bp_index_sheet[-1, ]

## Convert BP IDs to numeric values and sort
bp_index_sheet[ ,1] <- as.numeric(as.character(bp_index_sheet[, 1]))
attach(bp_index_sheet)
bp_sorted <- bp_index_sheet[order(bp_id), ]

## User to enter new report name
new_file <- readline(prompt="Enter filename: ")
new_filename <- paste(new_file, ".xls", sep = "")

## Check to see file exists
x <- file.exists(new_filename)
if(x == FALSE) {
  print("Agility Compare Report Filename INVALID")
  stop()
} else {
  print("Agility Compare Report GOOD")
  print("Press [enter] to continue")
  line <- readline()
}

## Load in the Baseline Report
wb2 <- loadWorkbook(new_filename)
## Load in the Sheets
bp_index_new <- readWorksheet(wb2, sheet=2)
bp_wklds_new <- readWorksheet(wb2, sheet=3)
bp_wklds_resources_new <- readWorksheet(wb2, sheet=4)
bp_vars_all_new <- readWorksheet(wb2, sheet=5)
bp_pkg_scrpt_ver_new <- readWorksheet(wb2, sheet=6)

## Clean Up bp_index_sheet and setup column headings
## bp_index_new[1,x] = column headings
bp_index_new <- bp_index_new[-1, ]
colnames(bp_index_new) <- c(bp_index_new[1,1], bp_index_new[1,2], bp_index_new[1,3], 
                              bp_index_new[1,4], bp_index_new[1,5], bp_index_new[1,6], 
                              bp_index_new[1,7], bp_index_new[1,8])
bp_index_new <- bp_index_new[-1, ]

## Convert BP IDs to numeric values and sort
bp_index_new[ ,1] <- as.numeric(as.character(bp_index_new[, 1]))
attach(bp_index_new)
bp_sorted_new <- bp_index_new[order(bp_id), ]

## BluePrint Stats from Baseline
bp_tot_base <- sum(complete.cases(bp_index_sheet[ ,1]))
print("Baseline Blueprints Total:")
print(bp_tot_base)

## BluePrint Stats from New Report
bp_tot_new <- sum(complete.cases(bp_index_new[ ,1]))
print("New Report Blueprints Total:")
print(bp_tot_new)

## Show Added Blueprint in New Report
added_new <- sqldf('SELECT * FROM bp_index_new EXCEPT SELECT * FROM bp_index_sheet')
print("New Blueprints")
print(added_new[, 1])

## Show Deleted Blueprints in New Report ()
missing_new <- sqldf('SELECT * FROM bp_index_sheet EXCEPT SELECT * FROM bp_index_new')
print("Deleted Blueprints")
print(missing_new[, 1])
