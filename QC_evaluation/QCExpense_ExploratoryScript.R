# QC Bonus monitoring - Exploratory analysis
# Result are planned for use in RMarkdown report

# set working directory
setwd("C:/Users/PB/SkyDrive/DG Evolt/QAS Dept/Monthly Bonus")

# Load library
library(readxl)
library(dplyr)
library(ggplot2)
library(zoo)
library(reshape2)
library(lubridate)

# Import data
FileName <- "QC_Bonus_monitoring_R.xlsx"
QCMonitor <- read_excel(FileName)

# Convert to factor
ind <- c("WorkNumber", "QCOperatorName", "Note")
QCMonitor[ind] <- lapply(QCMonitor[ind], factor)

# Mutate to get month year format
QCMonitor$Month <- QCMonitor$Date %>%
      format("%b-%Y") %>%
      as.yearmon("%b-%Y")
# Clean up date format
QCMonitor$Date <- ymd(QCMonitor$Date)

# month of report
MonthQC <- paste(month((month(Sys.Date())-1), label = TRUE), year(Sys.Date()), sep = "-")
# Filter the QC Operator working for the last months
ind <- QCMonitor %>%
      filter(month(Date) == month(Sys.Date()) - 1 &
                   year(Date) == year(Sys.Date())) %>%
      select(QCOperatorName) %>%
      unique() %>%
      t() %>% c() %>% as.factor()

QCMonitor_2m <- QCMonitor %>% 
      filter(QCOperatorName %in% ind)

# Summary
summary(QCMonitor)
summary(QCMonitor_2m)

########################
# Monthly total monitoring

QCMonitor_Dept <- QCMonitor %>%
      select(Month, BasicSalary, OT, WKD, QCBonus) %>%
      group_by(Month) %>%
      summarise_all(sum) %>%
      melt(id = "Month")

# Plotting the QC operator Dept. Expense
ggplot(QCMonitor_Dept,aes(x = Month, y = value )) + 
      geom_bar(aes(fill = variable), stat = "identity") +
      ylab("Expenses (RMB)") +
      ggtitle("QAS Dept. - QC Operator Expense Monitoring")

########################
# Individual Monitoring

# Using reshape2 package
QCMonitor_2m_r <- QCMonitor_2m %>%
      select(Month, QCOperatorName, BasicSalary, OT, WKD, QCBonus) %>%
      melt(id.vars = c("Month", "QCOperatorName"))

# Plotting for individual monitoring
ggplot(QCMonitor_2m_r,aes(x = Month, y = value)) + 
      geom_bar(aes(fill = variable), stat = "identity") +
      facet_wrap(~ QCOperatorName) +
      ylab("Expenses (RMB)") +
      ggtitle(paste("Individual QC Operator Expense Monitoring", MonthQC, sep = " - "))


############################################
# For future implementation
############################################

# Add staff salary to get global QAS expense
# 
# Add Absence and vacation data
# Add Start date
# 
# Monitor hours of OT and weekend OT