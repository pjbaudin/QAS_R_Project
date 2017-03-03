# QC Bonus calculation

# QC Bonus worth in RMB
QCLevel_RMB <- data.frame(QCLevel = c("QC0", "QC1", "QC2", "QC3", "QC4", "QC5", "QC6", "QCLeader"),
                          Amount = c(0, 250, 390, 460, 550, 675, 800, 1000))

QCLevel_RMB$QCLevel <- as.factor(QCLevel_RMB$QCLevel)

# Hourly rate
Hrate <- 10.05
Dayhour <- 8
MonthDay <- 21.75

# OT rate
Week_OT <- 1.5
WKD_OT <- 2

# Set working directory
setwd("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data/QAS_Dept_Admin")

# Load library
library(readxl)
library(dplyr)
library(ggplot2)
library(reshape2)
library(lubridate)
library(chron)
library(tidyr)

# Import data for OT time
FileName <- "OT_list.xlsx"
OTList <- read_excel(FileName)

# Clean up class for OTList
OTList$Date <- ymd(OTList$Date)
OTList$NightShift <- as.factor(OTList$NightShift)
OTList <- OTList %>%
        mutate(WeekNumber = as.Date(cut(Date, breaks = "week"))) %>%
        mutate(Month = as.Date(cut(Date, breaks = "month"))) %>%
        mutate(WeekDay = weekdays(Date)) %>%
        mutate(OTRate = 0)

for(i in 1:nrow(OTList)) {
        if(OTList$WeekDay[i] == "Sunday" | OTList$WeekDay[i] == "Saturday") {
                OTList$OTRate[i] <- "WKD" 
        } else {
                OTList$OTRate[i] <- "OT"
        }
}

# Adjust OTRate for Chinese New Year
ind <- c("2017-02-04", "2017-02-05")
OTList$OTRate[OTList$Date == ind] <- "OT"

OTList$OTRate <-  as.factor(OTList$OTRate)

# melt data for easier manipulation
OTList <- melt(OTList, id = c("Date", "NightShift", "WeekNumber", "Month", "WeekDay", "OTRate"))

# correction for Maisie
ind <- c("2017-02-05")
OTList[OTList$Date == "2017-02-05" & OTList$variable == "Maisie", "OTRate"] <- "WKD"

OTListMonthSum <- OTList %>%
        select(Month, variable, OTRate, value) %>%
        group_by(Month, variable, OTRate) %>%
        summarise(OThours = sum(value, na.rm = TRUE)) %>%
        spread(OTRate, OThours) %>%
        mutate(OTvalue = OT * Week_OT * Hrate) %>%
        mutate(WKDvalue = WKD * WKD_OT * Hrate)

colnames(OTListMonthSum)[2] <- "EmployeeName"

# Import absence data
FileName <- "Employee_absence_list.xlsx"
EmployeeAbsenceList <- read_excel(FileName)

EmployeeAbsenceList <- EmployeeAbsenceList %>%
        mutate(Month = as.Date(cut(Start_date, breaks = "month"))) %>%
        group_by(Month, EmployeeName) %>%
        summarise(TotalAbs = sum(Days, na.rm = TRUE))

QAS_QCOp <- left_join(OTListMonthSum, EmployeeAbsenceList, by = c("Month", "EmployeeName"))

ind <- is.na(QAS_QCOp$TotalAbs)

QAS_QCOp$TotalAbs[ind] <- 0

QAS_QCOp$Workdays <- MonthDay - QAS_QCOp$TotalAbs
QAS_QCOp$BaseSalary <- QAS_QCOp$Workdays * Dayhour * Hrate

QAS_QCOp$EmployeeName <-  as.factor(QAS_QCOp$EmployeeName)

# Import data for QC level
FileName <- "QCBonusLevel_Master.xlsx"
QCBonuslevel <- read_excel(FileName)

QCBonuslevel$StartDate <- ymd(QCBonuslevel$StartDate)
QCBonuslevel$DateQCBonus <- ymd(QCBonuslevel$DateQCBonus)
QCBonuslevel$EvalDate <- as.Date(cut(QCBonuslevel$EvalDate, breaks = "month"))

# Convert to factor
ind <- c("EmployeeNumber", "EmployeeName", "ChineseName", "Class", "Position", "QCLevel", "Status")
QCBonuslevel[ind] <- lapply(QCBonuslevel[ind], factor)

# Add the QC level bonus in RMB with left join to QCLevel_RMB
QCBonuslevel <- left_join(QCBonuslevel, QCLevel_RMB, by = "QCLevel")
QCBonuslevel$QCLevel <- as.factor(QCBonuslevel$QCLevel)

QCBonuslevel <- QCBonuslevel[-nrow(QCBonuslevel), ]

QCBonuslevel <- QCBonuslevel %>%
        filter(Status == "Active")

QAS_QCOp <- left_join(QAS_QCOp, QCBonuslevel, by = c("EmployeeName", "Month" = "EvalDate"))
QAS_QCOp$EmployeeName <- as.factor(QAS_QCOp$EmployeeName)

QAS_QCOp$ServiceTime <- floor(as.numeric(QAS_QCOp$Month - QAS_QCOp$StartDate) / 365.25)
ind <- QAS_QCOp$ServiceTime <= 0 | is.na(QAS_QCOp$ServiceTime)
QAS_QCOp$ServiceTime[ind] <- 0

QAS_QCOp$HoldQCLevelTime <- duration(
        as.numeric(difftime(QAS_QCOp$Month, QAS_QCOp$DateQCBonus)),
        "seconds")

QAS_QCOp$ServiceBonus <- 0

for(i in 1:nrow(QAS_QCOp)) {
        if(QAS_QCOp$ServiceTime[i] <= 0) {
                QAS_QCOp$ServiceBonus[i] <- 0
        } else if(QAS_QCOp$ServiceTime[i] == 1) {
                QAS_QCOp$ServiceBonus[i] <- 100
        } else if(QAS_QCOp$ServiceTime[i] == 2) {
                QAS_QCOp$ServiceBonus[i] <- 150
        } else if(QAS_QCOp$ServiceTime[i] == 3){
                QAS_QCOp$ServiceBonus[i] <- 200
        } else if(QAS_QCOp$ServiceTime[i] == 4){
                QAS_QCOp$ServiceBonus[i] <- 250
        } else if(QAS_QCOp$ServiceTime[i] == 5){
                QAS_QCOp$ServiceBonus[i] <- 300
        } else if(QAS_QCOp$ServiceTime[i] == 6){
                QAS_QCOp$ServiceBonus[i] <- 320
        } else if(QAS_QCOp$ServiceTime[i] == 7){
                QAS_QCOp$ServiceBonus[i] <- 340
        } else if(QAS_QCOp$ServiceTime[i] == 8){
                QAS_QCOp$ServiceBonus[i] <- 360
        } else if(QAS_QCOp$ServiceTime[i] == 9){
                QAS_QCOp$ServiceBonus[i] <- 380
        } else if(QAS_QCOp$ServiceTime[i] == 10){
                QAS_QCOp$ServiceBonus[i] <- 400
        }
}

# Reference for long service bonus
# 1 nian	100
# 2 nian	150
# 3 nian	200
# 4 nian	250
# 5 nian	300
# 6 nian	320
# 7 nian	340
# 8 nian	360
# 9 nian	380
# 10 nian	400

QAS_QCOp$QCEvalAmount <- (QAS_QCOp$ServiceBonus + QAS_QCOp$Amount) * QAS_QCOp$QCEval

QAS_QCOp$QCBonusAmount <- QAS_QCOp$ServiceBonus +
        QAS_QCOp$Amount +
        QAS_QCOp$QCEvalAmount

QAS_QCOp$GrandTotal <- QAS_QCOp$OTvalue +
        QAS_QCOp$WKDvalue +
        QAS_QCOp$BaseSalary +
        QAS_QCOp$QCEvalAmount
        

QAS_QCOp <- na.omit(QAS_QCOp)

write.csv(QAS_QCOp, file = "MonthlySummary.csv")






