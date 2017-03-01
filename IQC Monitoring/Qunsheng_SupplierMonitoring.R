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
IQC_df <- select(IQC_df, InspDate, Area, ItemCode, DocNum, PI, NCN, NCNType,
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

# If severity is NA, replace by 0
ind <- is.na(IQC_df$Severity)
IQC_df[ind, "Severity"] <- 0

#################################
# IQC results for selected parts

# File name and import for Qunsheng
FileName <- "./Supplier Data/QunshengPartList.csv"

ind <- read.csv(FileName, header = FALSE, col.names = "PartNumber")
ind$PartNumber <- as.factor(ind$PartNumber)

IQC_Qunsheng_List <- IQC_df %>%
        filter(ItemCode %in% ind$PartNumber) %>%
        select(InspDate, ItemCode, NCN, Severity, IssueCategory, Result, ItemName,
               Quantity, CardCode, CardFName, Price, WeekNumber, MonthInsp) %>%
        mutate(DocPrice = Quantity * Price) %>%
        group_by(ItemCode) %>%
        droplevels()

# Plot by parts and export graph and summary from Daniel's List

SinceDate <- min(IQC_Qunsheng_List$InspDate)

File_Name <- paste("./Supplier Data/QunshengSummaryGraphDanielList_Since_", SinceDate, ".png", sep = "")

png(File_Name, width = 1102, height = 720, units = "px")
ggplot(IQC_Qunsheng, aes(x = Quantity, y = DocPrice, colour = Severity)) +
        geom_point(size = 4, alpha = 0.8) +
        scale_color_continuous(low = "#009E73", high = "red") +
        geom_smooth(method = "lm", se = FALSE) +
        facet_wrap(~ ItemCode) +
        ggtitle(paste("S801805 - Incoming Quality Performance - Since ", SinceDate, sep = "")) +
        xlab("Quantity per incoming lot") +
        ylab("Lot Value")
dev.off()

SummaryQunshengList <- IQC_Qunsheng_List %>%
        group_by(ItemCode, ItemName) %>%
        summarise(TotalQty = sum(Quantity),
                  MeanSeverity = round(mean(Severity), digit = 2),
                  total_Accepted = sum(Result == "Accepted"),
                  total_Deviation = sum(Result == "Deviation"),
                  total_Rejected = sum(Result == "Rejected")) %>%
        arrange(desc(MeanSeverity))

FileName <- paste("./Supplier Data/QunshengSummaryDanielList_Since_", SinceDate, ".csv", sep = "")
write.csv(SummaryQunshengList, file = FileName)

# Plot by parts and export graph and summary from SAP data

IQC_Qunsheng <- IQC_df %>%
        filter(CardCode == "S801805") %>%
        select(InspDate, ItemCode, NCN, Severity, IssueCategory, Result, ItemName,
               Quantity, CardCode, CardFName, Price, WeekNumber, MonthInsp) %>%
        mutate(DocPrice = Quantity * Price) %>%
        group_by(ItemCode) %>%
        droplevels()

SinceDate <- min(IQC_Qunsheng$InspDate)

File_Name <- paste("./Supplier Data/QunshengSummaryGraphSAP_Since_", SinceDate, ".png", sep = "")

png(File_Name, width = 1102, height = 720, units = "px")
ggplot(IQC_Qunsheng, aes(x = Quantity, y = DocPrice, colour = Severity)) +
        geom_point(size = 4, alpha = 0.8) +
        scale_color_continuous(low = "#009E73", high = "red") +
        geom_smooth(method = "lm", se = FALSE) +
        facet_wrap(~ ItemCode) +
        ggtitle(paste("S801805 - Incoming Quality Performance - Since ", SinceDate, sep = "")) +
        xlab("Quantity per incoming lot") +
        ylab("Lot Value")
dev.off()

SummaryQunsheng <- IQC_Qunsheng %>%
        group_by(ItemCode, ItemName) %>%
        summarise(TotalQty = sum(Quantity),
                  MeanSeverity = round(mean(Severity), digit = 2),
                  total_Accepted = sum(Result == "Accepted"),
                  total_Deviation = sum(Result == "Deviation"),
                  total_Rejected = sum(Result == "Rejected")) %>%
        arrange(desc(MeanSeverity))

FileName <- paste("./Supplier Data/QunshengSummarySAP_Since_", SinceDate, ".csv", sep = "")
write.csv(SummaryQunsheng, file = FileName)

# Investigation on additional parts not listed
DanielList <- SummaryQunshengList %>% select(ItemCode, ItemName) %>% droplevels()
SAPList <- SummaryQunsheng %>% select(ItemCode, ItemName) %>% droplevels()

NotinDanielList <- anti_join(SAPList, DanielList)

FileName <- paste("./Supplier Data/ExportNotinDanielList.csv", sep = "")
write.csv(NotinDanielList, file = FileName)

