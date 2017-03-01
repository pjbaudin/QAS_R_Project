# QAS Dept monitoring - Expense Exploratory analysis
# Result are planned for use in RMarkdown report

# set working directory
setwd("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data")

# Load library
library(readxl)
library(dplyr)
library(ggplot2)
library(reshape2)
library(lubridate)

# Import data
FileName <- "Employee_absence_list.xlsx"
QASAttendance <- read_excel(FileName)

# Convert to factor
ind <- c("EmployeeName", "EmployeeNumber", "Type_Leave")
QASAttendance[ind] <- lapply(QASAttendance[ind], factor)

# Clean up date format
QASAttendance$Start_date <- ymd(QASAttendance$Start_date)
QASAttendance$End_date <- ymd(QASAttendance$End_date)

# Mutate to get month year format
# and filter employee who resigned

ind <- t(QASAttendance[QASAttendance$Type_Leave == "End date", "EmployeeName"])
QASAttendance <- QASAttendance %>%
        mutate(WeekNumber = as.Date(cut(Start_date, breaks = "week"))) %>%
        mutate(Month = as.Date(cut(Start_date, breaks = "month"))) %>%
        mutate(Quarter = as.Date(cut(Start_date, breaks = "quarter"))) %>%
        filter(!(EmployeeName %in% ind)) %>%
        droplevels()

# month of report
MonthAttendanceQAS <- paste(month((month(Sys.Date()) - 1), label = TRUE), year(Sys.Date()), sep = "-")


#######################################
# QAS Attendance Monitoring
########################################

ggplot(QASAttendance, aes(x = Quarter, y = Days, colour = EmployeeName)) +
        geom_point() +
        facet_wrap(~ EmployeeName)

ggplot(QASAttendance, aes(x = EmployeeName)) +
        geom_bar(aes(fill = Days)) +
        facet_wrap(~ Quarter)

QASAttendanceMonth <- QASAttendance %>%
        group_by(EmployeeName, Month) %>%
        summarise(TotalDays = sum(Days), MeanMonth = mean(Days))

QASAttendanceQuarter <- QASAttendance %>%
        group_by(Quarter, EmployeeName) %>%
        summarise(TotalDays = sum(Days), MeanQuarter = mean(Days)) %>%
        filter(!is.na(TotalDays))

ggplot(QASAttendanceQuarter, aes(x = Quarter, y = TotalDays)) +
        geom_point() +
        facet_wrap(~ EmployeeName)

ggplot(QASAttendanceQuarter, aes(x = Quarter, y = TotalDays)) +
        geom_col(position = "identity", aes(fill = EmployeeName)) +
        geom_hline(yintercept = mean(QASAttendanceQuarter$TotalDays),
                   col = "red", lwd = 1.3, lty = 2, show.legend = TRUE) +
        facet_wrap(~ EmployeeName)

# Monitor hours of OT and weekend OT