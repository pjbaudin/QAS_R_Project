# QAS Dept monitoring - Exploratory analysis
# Result are planned for use in RMarkdown report

# set working directory
setwd("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data")

# Load library
library(readxl)
library(dplyr)
library(ggplot2)
library(zoo)
library(reshape2)
library(lubridate)
library(plotly)

# Import data
FileName <- "QAS_Expense_monitoring.xlsx"
QASMonitor <- read_excel(FileName)

# Convert to factor
ind <- c("WorkNumber", "Beneficiary", "Category", "Note")
QASMonitor[ind] <- lapply(QASMonitor[ind], factor)

# Clean up date format
QASMonitor$Date <- ymd(QASMonitor$Date)

# Mutate to get month year format
QASMonitor <- QASMonitor %>%
        mutate(WeekNumber = as.Date(cut(Date, breaks = "week"))) %>%
        mutate(Month = as.Date(cut(Date, breaks = "month")))

# Replace all the NA for staff as zeros
QASMonitor$OT[is.na(QASMonitor$OT)] <- 0 
QASMonitor$WKD[is.na(QASMonitor$WKD)] <- 0
QASMonitor$QCBonus[is.na(QASMonitor$QCBonus)] <- 0

# recalculate total (never trust excel)
QASMonitor <- QASMonitor %>% 
        mutate(Total = BaseSalary + OT + WKD + QCBonus)

# month of report
MonthQAS <- paste(month((month(Sys.Date()) - 1), label = TRUE), year(Sys.Date()), sep = "-")

# Filter the QC Operator working for the last months
ind <- QASMonitor %>%
      filter(month(Date) == month(Sys.Date()) - 1 &
                   year(Date) == year(Sys.Date()) &
                     Category == "Operator salary") %>%
      select(Beneficiary) %>%
      unique() %>%
      t() %>% c() %>% as.factor()

QCMonitor_2m <- QASMonitor %>% 
      filter(Beneficiary %in% ind)

# Summary
# summary(QASMonitor)
# summary(QCMonitor_2m)


#######################################
# QC Operator Monitoring
########################################
# Monthly total monitoring for QC Operator

QCMonitor_Dept <- QASMonitor %>%
      filter(Category == "Operator salary") %>%
      select(Month, BaseSalary, OT, WKD, QCBonus) %>%
      group_by(Month) %>%
      summarise_all(sum) %>%
      melt(id = "Month")

# Plotting the QC operator Dept. Expense
ggplot(QCMonitor_Dept,aes(x = Month, y = value )) + 
      geom_bar(aes(fill = variable), stat = "identity") +
        facet_wrap(~ variable) +
        geom_smooth(se = FALSE) +
      ylab("Expenses (RMB)") +
      ggtitle("QAS Dept. - QC Operator Expense Monitoring")

########################
# Individual Monitoring

# Using reshape2 package
QCMonitor_2m_r <- QCMonitor_2m %>%
      select(Month, Beneficiary, BaseSalary, OT, WKD, QCBonus) %>%
      melt(id.vars = c("Month", "Beneficiary"))

# Plotting for individual monitoring
ggplot(QCMonitor_2m_r,aes(x = Month, y = value)) + 
        geom_bar(aes(fill = variable), stat = "identity") +
        facet_wrap(~ Beneficiary) +
        geom_smooth() +
        ylab("Expenses (RMB)") +
        ggtitle(paste("Individual QC Operator Expense Monitoring", MonthQAS, sep = " - "))

#######################################
# QAS Expense Monitoring
########################################

# The palette with grey:
cbPalette <- c("#999999", "#E69F00", "#56B4E9", "#009E73", "#F0E442", "#0072B2", "#D55E00", "#CC79A7")

# Plot by category of expense with trend
ggplot(QASMonitor, aes(x = Month, y = Total)) +
        geom_bar(aes(fill = Category), stat = "identity", position = "dodge",
                 alpha = 0.8, colour = "#0072B2") +
        geom_smooth(method = "loess", se = FALSE) +
        facet_wrap(~ Category) +
        scale_fill_manual(values = cbPalette) +
        scale_x_date(date_breaks = "1 month", date_labels = "%b-%y")

# Plot for overal spending and trend
QAS_Sum <-  QASMonitor %>%
        group_by(Month) %>%
        summarise(GrandTotal = sum(Total))

ggplot(QAS_Sum, aes(x = Month, y = GrandTotal)) +
        geom_line(color = "green", lwd = 1.2) +
        geom_point(size = 5, alpha = 0.8, colour = "darkgreen") +
        geom_smooth(se = FALSE, colour = "grey", alpha = 0.8) +
        scale_x_date(date_breaks = "1 month", date_labels = "%b-%y") +
        ylim(0, max(QAS_Sum$GrandTotal) + 5000)

############################################
# For future implementation
############################################


p <- ggplotly(p)

# Linear model trial
QAS_Sum_gran <-  QASMonitor %>%
        group_by(Month) %>%
        select(Month, BaseSalary, OT, WKD, QCBonus) %>%
        summarise_each(funs(sum))

QAS_model_all <- lm(QCBonus ~ ., data = QAS_Sum_gran[, -1])
QAS_model_all_step <- step(QAS_model_all, k=log(nrow(QAS_Sum_gran[, -1])))

summary(QAS_model_all_step)

# Add staff salary to get global QAS expense
# 
# Add Absence and vacation data
# Add Start date
# 
# Monitor hours of OT and weekend OT