# R script to monitor the IQC issue on weekly basis
# (to be integrated into a Rmd file)

# Load packages
library(readxl)
library(dplyr)
library(lubridate)
library(ggplot2)
library(RODBC)      # to query SAP database
library(DBI)        # DBI convention for MySQL
library(reshape2)

###########################
# Data import

# File path name
fileName <- "C:/Users/PB/SkyDrive/DG Evolt/QASsvn/IQC Monitoring/IQC_defect_table.xlsx"

# Set working directory
setwd("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data")

# Read the sheet names
IQCMonitoring_sheet <- excel_sheets(fileName)

# Import AOI dataset
IQC_df <- suppressWarnings(read_excel(fileName, sheet = 1))

######################
# Data Cleaning

# Remove last line of excel rubish
ind <- dim(IQC_df)[1]
IQC_df <- IQC_df[-ind, ]

# Rename variables
names(IQC_df) <- c("InspDate", "Area", "PartCategory", "ItemCode",
                   "PartName", "Supplier", "DocNum", "PI",
                   "NCN", "NCNType", "Severity", "IssueCategory",
                   "IssueNote", "Result", "Redmine", "RedmineStatus",
                   "SamplingSize", "RJ_NGCount", "OQCReport", "MaterialLabel",
                   "Who")

# Filter inccorrect Doc Number (Good Receipt PO) / must start with 3xxxx
ind <- grepl("^[3].", IQC_df$DocNum)
ind <- IQC_df$DocNum[ind]
IQC_df <- IQC_df %>% filter(DocNum %in% ind)

# Convert to factor
ind <- c("Area", "PartCategory", "ItemCode", "PartName", "Supplier", "DocNum",
         "PI", "NCN", "NCNType", "IssueCategory", "Result", "Redmine", "RedmineStatus",
         "OQCReport", "MaterialLabel", "Who")
IQC_df[ind] <- lapply(IQC_df[ind], factor)

# Convert to numeric
ind <- c("SamplingSize", "RJ_NGCount")
IQC_df[ind] <- suppressWarnings(lapply(IQC_df[ind], as.numeric))

# Convert date in dmy
IQC_df$InspDate <- ymd(IQC_df$InspDate)

# Final observation selection from IQC Monitoring table
IQC_df <- IQC_df %>% select(InspDate, Area,ItemCode, DocNum, PI, NCN, NCNType,
                            Severity, IssueCategory, IssueNote, Result, Redmine,
                            RedmineStatus, SamplingSize, RJ_NGCount, OQCReport,
                            MaterialLabel, Who)

# Date for the report
WIQC <- paste(week(Sys.Date()) - 1, "-", year(Sys.Date()), sep = "")


###########################
# Import data from SAP (Good receipt PO)

con <- odbcDriverConnect("ODBC;Description=SBO;DRIVER=SQL Server Native Client 11.0;SERVER=DB0MRL82;
                         UID=evoltENG;PWD=Engineering123;DATABASE=ITesting;ApplicationIntent=READONLY;")

#List table on database
odbcGetInfo(con)

SQL <- paste("SELECT
             OPDN.DocNum, PDN1.ItemCode, OITM.ItemName, PDN1.Quantity, OPDN.DocDate, OPDN.CardCode, 
             OCRD.CardFName, OITB.ItmsGrpNam, POR1.Price, POR1.ShipDate
             FROM
             OPDN
             INNER JOIN OCRD ON OPDN.CardCode = OCRD.CardCode
             INNER JOIN PDN1 ON OPDN.DocEntry = PDN1.DocEntry
             INNER JOIN OITM ON PDN1.ItemCode = OITM.ItemCode
             INNER JOIN OITB ON OITM.ItmsGrpCod = OITB.ItmsGrpCod
             INNER JOIN POR1 ON PDN1.ItemCode = POR1.ItemCode AND PDN1.BaseEntry = POR1.DocEntry
             WHERE
             OPDN.DocNum LIKE '%'
             AND PDN1.ItemCode LIKE '%'")

# Get results from query
SAPres <- sqlQuery(con, SQL)
odbcClose(con)

# Formating for analysis
SAPres$DocNum <- as.factor(SAPres$DocNum)

# Convert to date format
ind <- c("DocDate", "ShipDate")
SAPres[ind] <- lapply(SAPres[ind], ymd)

########################
# Merge IQC table and SAP database

IQC_df <- suppressWarnings(full_join(IQC_df, SAPres, by = c("DocNum", "ItemCode")))

# Convert to factor
ind <- c("ItemCode", "DocNum")
IQC_df[ind] <- lapply(IQC_df[ind], as.factor)

# Align dates with doc dates if NA
ind <- is.na(IQC_df$InspDate)
IQC_df[ind, "InspDate"] <- IQC_df[ind, "DocDate"]

# Mutate to get week inspection number
IQC_df <- IQC_df %>%
        mutate(WeekNumber = as.Date(cut(InspDate, breaks = "week"))) %>%
        mutate(MonthInsp = as.Date(cut(InspDate, breaks = "month")))

# If Na replace line by lot Accepted
ind <- is.na(IQC_df$Result)
IQC_df[ind, "Result"] <- "Accepted"

################
# Weekly report for the system week - 1

IQC_week <- IQC_df %>%
      filter(week(WeekNumber) == week(Sys.Date()) - 1 &
                   year(WeekNumber) == year(Sys.Date()))

SAP_weekKPI <- SAPres %>%
      filter(week(DocDate) == week(Sys.Date()) - 1 &
                   year(DocDate) == year(Sys.Date())) %>%
      droplevels() %>%
      group_by(DocNum)

IQC_weekKPI <- IQC_week %>%
      filter(week(WeekNumber) == week(Sys.Date()) - 1 &
                     year(WeekNumber) == year(Sys.Date())) %>%
      filter(Result == "Rejected" | Result == "Deviation")

###############
# Weekly summary

weekKPI <- sum(IQC_weekKPI$Result == "Rejected") / length(unique(SAP_weekKPI$DocNum))
outA <- paste("Week", WIQC, "IQC Percentage of lot rejected:",
           round((weekKPI) * 100, digits = 2),"%", sep=" ");
outB <- paste("Number of part received in SAP: ", length(SAP_weekKPI$ItemCode), sep = "");
outB2 <- paste("Number of part received and inspected: ", length(IQC_week$ItemCode), sep = "");
outC <- paste("Number of lot received: ", length(unique(SAP_weekKPI$DocNum)), sep = "");
outD <- paste("Average number of part per lot received: ",
      round(length(SAP_weekKPI$ItemCode) / length(unique(SAP_weekKPI$DocNum)), digits = 2),
      sep = "");
outE <- paste("Number of rejected lots: ", sum(IQC_week$Result == "Rejected"), sep = "");
outF <- paste("Number of lots accepted with deviation: ", sum(IQC_week$Result == "Deviation"), sep = "");


# On-time receiving
# Need to pull the ITR transfer date from SAP
IQCon_time <- mean(IQC_df$InspDate - IQC_df$DocDate, na.rm = TRUE)
outG <- paste("Mean inspection time: ", round(IQCon_time, digits = 2), " days", sep = "")

IQCweekResult <- data.frame(Result = c(outA, outB, outB2, outC, outD, outE, outF, outG))

##########
# PPAP monitoring for review and action

IQC_PPAP <- IQC_week %>% 
      filter(IssueCategory == "PPAP") %>%
      select(ItemCode, ItemName)

######
# export result and PPAP list in csv
mainDir <- "C:/Users/PB/SkyDrive/DG Evolt/QAS_Data"
subDir <- WIQC

dir.create(file.path(mainDir, subDir), showWarnings = FALSE)

setwd(file.path(mainDir, subDir))
File_Name <- paste("IQC_", WIQC, "PPAP_follow_up.csv", sep = "")
write.csv(IQC_PPAP, File_Name)

File_Name <- paste("IQC_", WIQC, "Monitoring_Result.csv", sep = "")
write.csv(IQCweekResult, File_Name)


# Weekly report for the system week

IQC_week <- IQC_df %>%
        filter(WeekNumber == week(Sys.Date()) &
                       year(WeekNumber) == year(Sys.Date()))

SAP_weekKPI <- SAPres %>%
        filter(week(DocDate) == week(Sys.Date()) &
                       year(DocDate) == year(Sys.Date())) %>%
        droplevels() %>%
        group_by(DocNum)

IQC_weekKPI <- IQC_week %>%
        filter(WeekNumber == week(Sys.Date()) &
                       year(WeekNumber) == year(Sys.Date())) %>%
        filter(Result == "Rejected" | Result == "Deviation")

###############
# Weekly summary

weekKPI <- sum(IQC_weekKPI$Result == "Rejected") / length(unique(SAP_weekKPI$DocNum))
outA <- paste("Week", "ThisWeek", "IQC Percentage of lot rejected:",
              round((weekKPI) * 100, digits = 2),"%", sep=" ");
outB <- paste("Number of part received in SAP: ", length(SAP_weekKPI$ItemCode), sep = "");
outB2 <- paste("Number of part received and inspected: ", length(IQC_week$ItemCode), sep = "");
outC <- paste("Number of lot received: ", length(unique(SAP_weekKPI$DocNum)), sep = "");
outD <- paste("Average number of part per lot received: ",
              round(length(SAP_weekKPI$ItemCode) / length(unique(SAP_weekKPI$DocNum)), digits = 2),
              sep = "");
outE <- paste("Number of rejected lots: ", sum(IQC_week$Result == "Rejected"), sep = "");
outF <- paste("Number of lots accepted with deviation: ", sum(IQC_week$Result == "Deviation"), sep = "");


# On-time receiving
# Need to pull the ITR transfer date from SAP
IQCon_time <- mean(IQC_df$InspDate - IQC_df$DocDate, na.rm = TRUE)
outG <- paste("Mean inspection time: ", round(IQCon_time, digits = 2), " days", sep = "")

IQCweekResult <- data.frame(Result = c(outA, outB, outB2, outC, outD, outE, outF, outG))

### Add in future Rmd report with template

##########
# PPAP monitoring for review and action

IQC_PPAP <- IQC_week %>% 
        filter(IssueCategory == "PPAP") %>%
        select(ItemCode, ItemName)

######
# export result and PPAP list in csv
mainDir <- "C:/Users/PB/SkyDrive/DG Evolt/QAS_Data"
subDir <- "ThisWeek"

dir.create(file.path(mainDir, subDir), showWarnings = FALSE)

setwd(file.path(mainDir, subDir))
File_Name <- paste("IQC_", "ThisWeek", "PPAP_follow_up.csv", sep = "")
write.csv(IQC_PPAP, File_Name)

File_Name <- paste("IQC_", "ThisWeek", "Monitoring_Result.csv", sep = "")
write.csv(IQCweekResult, File_Name)


#########################################
# Weekly performance monitoting (rolling over 4 weeks)

IQC_rolling4w <- IQC_df %>%
        filter(WeekNumber >= Sys.Date() - as.difftime(10, units = "weeks") &
                       weeks(WeekNumber) < weeks(Sys.Date()))
        
        # filter(WeekNumber > week(Sys.Date()) - 5 &
        #                WeekNumber < week(Sys.Date()) &
        #                year(InspDate) == year(Sys.Date()))
        

# IQC_rolling4w_unique <- aggregate(ItemCode ~ WeekNumber, data = IQC_rolling4w,
#                            FUN = function(x) length(unique(x)))

# Calculate how many parts received per week
IQC_rolling4w_total <- aggregate(ItemCode ~ WeekNumber, data = IQC_rolling4w,
                                 FUN = length)

mean(IQC_rolling4w_total$ItemCode)

testdf <- IQC_rolling4w %>%
        select(WeekNumber, ItemCode, Result) %>%
        group_by(WeekNumber) %>%
        summarize(total_Accepted = sum(Result == "Accepted"),
                  total_Deviation = sum(Result == "Deviation"),
                  total_Rejected = sum(Result == "Rejected")) %>%
        melt(id = "WeekNumber")

# The palette with grey:
cbPalette <- c("#009E73", "#E69F00", "#D55E00")

######
# export plot
mainDir <- "C:/Users/PB/SkyDrive/DG Evolt/QAS_Data"
subDir <- WIQC

dir.create(file.path(mainDir, subDir), showWarnings = FALSE)

setwd(file.path(mainDir, subDir))
File_Name <- paste("IQC_", WIQC, "10week_rolling.png", sep = "")

png(File_Name, width = 1102, height = 720, units = "px")
ggplot(testdf, aes(WeekNumber, value )) +
        geom_bar(aes(fill = variable), stat = "identity", position = "stack",
                 alpha = 0.6, colour = "darkgrey") +
        scale_fill_manual(values = cbPalette) +
        geom_smooth(aes(colour = variable), method = "auto", se = FALSE) +
        scale_colour_manual(values = cbPalette) +
        scale_x_date(date_breaks = "1 week", date_labels = "%W") +
        ggtitle("IQC - Quality Performance Monitoring") +
        ylab("Number of parts inspected")
dev.off()


#######################################
# For future implementation
#######################################

# Cluster test
# 
# IQC_clus <- IQC_df %>% select(PartNumber, Supplier,Severity, Result) %>%
#       group_by(Result) %>%
#       na.omit()
# hc <- hclust(dist(IQC_clus$Supplier), "ave")
# plot(hc)