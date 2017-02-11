# Load packages
library(readxl)
library(dplyr)
library(lubridate)
library(ggplot2)
library(RODBC)      # to query SAP database
library(DBI)        # DBI convention for MySQL

###########################
# Data import

# File path name
fileName <- "C:/Users/PB/SkyDrive/DG Evolt/QASsvn/IQC Monitoring/IQC_defect_table.xlsx"

# Set working directory
setwd("D:/Data Science/QAS_R_Project/IQC Monitoring")

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
MIQC <- paste(month((month(Sys.Date()) - 1), label = TRUE), "-", year(Sys.Date()), sep = "")


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

IQC_df <- suppressWarnings(left_join(IQC_df, SAPres, by = c("DocNum", "ItemCode")))

# Convert to factor
ind <- c("ItemCode", "DocNum")
IQC_df[ind] <- lapply(IQC_df[ind], as.factor)

################
# Monthly report for the system month - 1

SAP_KPI <- SAPres %>%
      filter(month(DocDate) == month(Sys.Date()) - 1 &
                   year(DocDate) == year(Sys.Date())) %>%
      droplevels() %>%
      group_by(DocNum)

IQC_KPI <- IQC_df %>%
      filter(month(DocDate) == month(Sys.Date()) - 1 &
                   year(DocDate) == year(Sys.Date())) %>%
      filter(Result == "Rejected" | Result == "Deviation")


KPI <- sum(IQC_KPI$Result == "Rejected") / length(unique(SAP_KPI$DocNum))
paste(MIQC, "IQC Percentage of lot rejected:", round((KPI)*100, digits = 2),"%", sep=" ")
paste("Number of part received: ", length(SAP_KPI$ItemCode), sep = "")
paste("Number of lot received: ", length(unique(SAP_KPI$DocNum)), sep = "")
paste("Number of rejected lots: ", sum(IQC_KPI$Result == "Rejected"), sep = "")
paste("Number of lots accepted with deviation: ", sum(IQC_KPI$Result == "Deviation"), sep = "")

# Plotting

# filter supplier with rejected and deviation lot and observe the trend
# since start of records 

#######################################
# For future implementation
#######################################

# On-time receiving
# Need to pull the ITR transfer date from SAP
IQCon_time <- mean(IQC_df$InspDate - IQC_df$DocDate, na.rm = TRUE)
IQCon_time

# Cluster test
# 
# IQC_clus <- IQC_df %>% select(PartNumber, Supplier,Severity, Result) %>%
#       group_by(Result) %>%
#       na.omit()
# hc <- hclust(dist(IQC_clus$Supplier), "ave")
# plot(hc)