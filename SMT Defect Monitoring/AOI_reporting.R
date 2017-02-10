# Load packages
library(readxl)
library(dplyr)

###########################
# Data import

# File path name
fileName <- "C:/Users/PB/SkyDrive/DG Evolt/QASsvn/QA Monitoring/QA report monitoring list.xlsx"

# Set workign directory
setwd("D:/Data Science/QAS_R_Project/SMT Defect Monitoring")

# copy filen in the working directory for manuipulation
file.copy(fileName, getwd())

# Read the sheet names
QAMonitoring_sheet <- excel_sheets(fileName)

# Import AOI dataset
AOI_df <- suppressWarnings(read_excel(fileName, sheet = 4))

######################
# Data Cleaning

# Rename variables
names(AOI_df) <- c("Date", "MONumber", "SAPMONumber", "PartNumber",
                   "PartName", "Quantity", "ReportRef", "A-sidePPM",
                   "B-sidePPM", "Average", "Notes")
# Convert to factor
ind <- c("MONumber", "SAPMONumber", "PartNumber", "PartName", "ReportRef")
AOI_df[ind] <- lapply(AOI_df[ind], factor)