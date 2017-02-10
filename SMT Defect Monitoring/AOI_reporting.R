# Load packages
library(readxl)
library(dplyr)
library(lubridate)

###########################
# Data import

# File path name
fileNameImport <- "C:/Users/PB/SkyDrive/DG Evolt/QASsvn/QA Monitoring/QA report monitoring list.xlsx"

# Set workign directory
setwd("D:/Data Science/QAS_R_Project/SMT Defect Monitoring")

# copy filen in the working directory for manuipulation
file.copy(fileNameImport, getwd())

# Filename
fileName <- "QA report monitoring list.xlsx"

# Read the sheet names
QAMonitoring_sheet <- excel_sheets(fileName)

# Import AOI dataset
AOI_df <- suppressWarnings(read_excel(fileName, sheet = 4))

######################
# Data Cleaning

# Remove last line of excel rubish
ind <- dim(AOI_df)[1]
AOI_df <- AOI_df[-ind, ]

# Rename variables
names(AOI_df) <- c("Date", "MONumber", "SAPMONumber", "PartNumber",
                   "PartName", "Quantity", "ReportRef", "A-sidePPM",
                   "B-sidePPM", "Average", "Notes")
# Convert to factor
ind <- c("MONumber", "SAPMONumber", "PartNumber", "PartName", "ReportRef")
AOI_df[ind] <- lapply(AOI_df[ind], factor)

# Convert date in dmy
AOI_df$Date <- ymd(AOI_df$Date)

################
# Monthly report

AOI_KPI <- AOI_df %>%
      filter(month(Date) == month(Sys.Date()) - 1 &
                   year(Date) == year(Sys.Date())) %>%
      select(Average) %>%
      t() %>% mean () %>% round(digits = 2)

paste(month(Sys.Date()) - 1, "-", year(Sys.Date()), " KPI : ", AOI_KPI, " PPM", sep = "")

###############
# Plotting
ind <- AOI_df %>%
      filter(month(Date) == month(Sys.Date()) - 1 &
                   year(Date) == year(Sys.Date())) %>%
      select(PartNumber) %>%
      unique()

AOI_plot <- AOI_df %>%
      filter(ind %in% PartNumber) %>%
      select

