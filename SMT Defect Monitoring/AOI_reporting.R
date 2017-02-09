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

