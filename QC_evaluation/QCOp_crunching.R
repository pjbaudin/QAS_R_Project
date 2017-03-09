##################
# QAS Expense Master Script
# 
## Data processing:
# - QC operator OT and bonus calculation
#       OT_list.xlsx
#       Employee_absence_list.xlsx
#       QCBonusLevel_Master.xlsx
# 
# - QAS Office expense
#       QAS_Office_Expense.xlsx


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

##########################################
# QC operator OT and bonus calculation
##########################################

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

##########
# Adjust OTRate for Chinese New Year
ind <- c("2017-02-04", "2017-02-05")
OTList$OTRate[OTList$Date == ind] <- "OT"
OTList$OTRate <-  as.factor(OTList$OTRate)
###########

# melt data for easier manipulation
OTList <- melt(OTList, id = c("Date", "NightShift", "WeekNumber", "Month", "WeekDay", "OTRate"))

##########
# correction for Maisie
ind <- c("2017-02-05")
OTList[OTList$Date == "2017-02-05" & OTList$variable == "Maisie", "OTRate"] <- "WKD"
##########

# Mutate to calculate the OT values
OTListMonthSum <- OTList %>%
        select(Month, variable, OTRate, value) %>%
        group_by(Month, variable, OTRate) %>%
        summarise(OThours = sum(value, na.rm = TRUE)) %>%
        spread(OTRate, OThours) %>%
        mutate(OTvalue = OT * Week_OT * Hrate) %>%
        mutate(WKDvalue = WKD * WKD_OT * Hrate)

# Rename column for easier left_join
colnames(OTListMonthSum)[2] <- "EmployeeName"

# Import absence data
FileName <- "Employee_absence_list.xlsx"
EmployeeAbsenceList <- read_excel(FileName)

# Process Employee absence list
EmployeeAbsenceList <- EmployeeAbsenceList %>%
        mutate(Month = as.Date(cut(Start_date, breaks = "month"))) %>%
        group_by(Month, EmployeeName) %>%
        summarise(TotalAbs = sum(Days, na.rm = TRUE))

# Attach absence records to QC Operator OT list
QAS_QCOp <- left_join(OTListMonthSum, EmployeeAbsenceList, by = c("Month", "EmployeeName"))
QAS_QCOp$EmployeeName <-  as.factor(QAS_QCOp$EmployeeName)

# Set all NA of total absence to zero
ind <- is.na(QAS_QCOp$TotalAbs)
QAS_QCOp$TotalAbs[ind] <- 0

# Calculate the number of working days and base salary
QAS_QCOp$Workdays <- MonthDay - QAS_QCOp$TotalAbs
QAS_QCOp$BaseSalary <- QAS_QCOp$Workdays * Dayhour * Hrate

# Import data for QC level
FileName <- "QCBonusLevel_Master.xlsx"
QCBonuslevel <- read_excel(FileName)

# CLean up date to correct format for manipulation
QCBonuslevel$StartDate <- ymd(QCBonuslevel$StartDate)
QCBonuslevel$DateQCBonus <- ymd(QCBonuslevel$DateQCBonus)
QCBonuslevel$EvalDate <- as.Date(cut(QCBonuslevel$EvalDate, breaks = "month"))

# Convert to factor
ind <- c("EmployeeNumber", "EmployeeName", "ChineseName", "Class", "Position", "QCLevel", "Status")
QCBonuslevel[ind] <- lapply(QCBonuslevel[ind], factor)

# Add the QC level bonus in RMB with left join to QCLevel_RMB
QCBonuslevel <- left_join(QCBonuslevel, QCLevel_RMB, by = "QCLevel")
QCBonuslevel$QCLevel <- as.factor(QCBonuslevel$QCLevel)

# Remove NA line added at the joint level
QCBonuslevel <- QCBonuslevel[-nrow(QCBonuslevel), ]

# Join to QAS_QCOp table
QAS_QCOp <- left_join(QAS_QCOp, QCBonuslevel, by = c("EmployeeName", "Month" = "EvalDate"))
QAS_QCOp$EmployeeName <- as.factor(QAS_QCOp$EmployeeName)

# Calculate the service time
QAS_QCOp$ServiceTime <- floor(as.numeric(QAS_QCOp$Month - QAS_QCOp$StartDate) / 365.25)
ind <- QAS_QCOp$ServiceTime <= 0 | is.na(QAS_QCOp$ServiceTime)
QAS_QCOp$ServiceTime[ind] <- 0

# Calculate the time since last change of QC level
QAS_QCOp$HoldQCLevelTime <- duration(
        as.numeric(difftime(QAS_QCOp$Month, QAS_QCOp$DateQCBonus)),
        "seconds")

# Set the Service Bonus to zera
QAS_QCOp$ServiceBonus <- 0

# Loop to assign the relevant amount for service bonus
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

# Removing Non-operator from the df
QAS_QCOp <- QAS_QCOp[!is.na(QAS_QCOp$Class), ]

################
# Adding the QAS staff salaary and other expenses

# Import data
FileName <- "QAS_Office_Expense.xlsx"
QASMonitor <- read_excel(FileName)

# Process and cleanup data
QASMonitor$Date <- as.Date(cut(QASMonitor$Date, breaks = "month"))
QASMonitor$StartDate <- ymd(QASMonitor$StartDate)
ind <- c("WorkNumber", "Beneficiary", "Class", "Position", "Status")
QASMonitor[ind] <- lapply(QASMonitor[ind], factor)

# Rename columns for easy full join
colnames(QASMonitor) <- c("Month", "EmployeeNumber", "EmployeeName", "BaseSalary",
                          "QCBonusAmount", "Class", "Position", "Status",
                          "Notes", "StartDate")

# List key for the full join operation
ind <- c("Month", "EmployeeNumber", "EmployeeName", "BaseSalary",
         "Class", "Position", "Status", "StartDate")

# Full join to match the Operator data and the Staff/Dept data
QAS_QCOp <- full_join(QAS_QCOp, QASMonitor, by = ind)
ind <- c("EmployeeName", "EmployeeNumber", "Class", "Position")
QAS_QCOp[ind] <- lapply(QAS_QCOp[ind], factor)

# (Re) Calculate the service time in years
QAS_QCOp$ServiceTime <- floor(as.numeric(QAS_QCOp$Month - QAS_QCOp$StartDate) / 365.25)
ind <- QAS_QCOp$ServiceTime <= 0 | is.na(QAS_QCOp$ServiceTime)
QAS_QCOp$ServiceTime[ind] <- 0

# Set the NA to zero for expenses related variables
ind <- is.na(QAS_QCOp$ServiceBonus)
QAS_QCOp$ServiceBonus[ind] <- 0

ind <- is.na(QAS_QCOp$Amount)
QAS_QCOp$Amount[ind] <- 0

ind  <- is.na(QAS_QCOp$QCEval)
QAS_QCOp$QCEval[ind] <- 0

ind <- is.na(QAS_QCOp$OTvalue)
QAS_QCOp$OTvalue[ind] <- 0

ind <- is.na(QAS_QCOp$WKDvalue)
QAS_QCOp$WKDvalue[ind] <- 0


# Calculate the QC Evaluation Bonus (top up) for QC Operator
# Calculate the total QC Bonus amount for QC Operator
ind <- QAS_QCOp$Class == "QC Operator"

# Set QC Evaluation Bonus variable
QAS_QCOp$QCEvalAmount <- 0
QAS_QCOp$QCEvalAmount[ind] <- (QAS_QCOp$ServiceBonus[ind] + QAS_QCOp$Amount[ind]) * QAS_QCOp$QCEval[ind]
QAS_QCOp$QCBonusAmount[ind] <- QAS_QCOp$ServiceBonus[ind] +
        QAS_QCOp$Amount[ind] +
        QAS_QCOp$QCEvalAmount[ind]

# Calculate the Grand total for the QC operator
QAS_QCOp$GrandTotal <- QAS_QCOp$OTvalue +
        QAS_QCOp$WKDvalue +
        QAS_QCOp$BaseSalary +
        QAS_QCOp$QCBonusAmount

# filter for the relevant period with all accurate data
QAS_QCOp <- filter(QAS_QCOp, Month >= "2016-09-01" )

write.csv(QAS_QCOp, file = "MonthlySummary.csv")

# Export for the month QC Bonus Operator
Export_QCBonus <- QAS_QCOp[QAS_QCOp$Month == "2017-02-01" & QAS_QCOp$Class == "QC Operator",
                           c("Month", "EmployeeName", "EmployeeNumber", "OT", "WKD",
                             "OTvalue", "WKDvalue", "TotalAbs", "StartDate", "Position",
                             "QCLevel", "Amount", "ServiceBonus", "QCEvalAmount", "QCBonusAmount")]

write.csv(Export_QCBonus, file = "2017-02_MonthlyQCBonusSummary.csv")


# The palette with grey:
cbPalette <- c("#999999", "#E69F00", "#56B4E9", "#009E73", "#0072B2", "#D55E00", "#CC79A7")

ggplot(QAS_QCOp, aes(x = Month, y = GrandTotal, group = Month)) +
        geom_col(aes(fill = Class), colour = "darkgrey") +
        scale_fill_manual(values = cbPalette) +
        scale_x_date(date_breaks  = "1 month", date_labels = "%b-%y")

ggplot(QAS_QCOp, aes(x = Month, y = GrandTotal, group = Month)) +
        geom_col(aes(fill = Class), colour = "darkgrey") +
        scale_fill_manual(values = cbPalette) +
        facet_wrap(~ Class) +
        scale_x_date(date_breaks  = "1 month", date_labels = "%b-%y")




