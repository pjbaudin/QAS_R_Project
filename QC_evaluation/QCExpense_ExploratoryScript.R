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

# Filter the QC Operator working for the last 2 months
ind <- QCMonitor %>%
      filter(Date > Sys.Date() - 76) %>%
      select(QCOperatorName) %>%
      unique() %>%
      t() %>%
      c() %>%
      as.factor()

QCMonitor_2m <- filter(QCMonitor, QCOperatorName %in% ind)

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
      ggtitle("Individual QC Operator Expense Monitoring")