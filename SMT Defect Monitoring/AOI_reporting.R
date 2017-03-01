# Load packages
library(readxl)
library(dplyr)
library(lubridate)
library(ggplot2)

###########################
# Data import

# File path name
fileName <- "C:/Users/PB/SkyDrive/DG Evolt/QASsvn/QA Monitoring/QA report monitoring list.xlsx"

# Set working directory
setwd("D:/Data Science/QAS_R_Project/SMT Defect Monitoring")

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
                   "PartName", "Quantity", "ReportRef", "AsidePPM",
                   "BsidePPM", "Average", "Notes")
# Convert to factor
ind <- c("MONumber", "SAPMONumber", "PartNumber", "PartName", "ReportRef")
AOI_df[ind] <- lapply(AOI_df[ind], factor)

# Convert date in dmy
AOI_df$Date <- ymd(AOI_df$Date)

# Date for the report
MAOI <- paste(month((month(Sys.Date()) - 1), label = TRUE), "-", year(Sys.Date()), sep = "")

################
# Monthly report for the system month - 1

AOI_KPI <- AOI_df %>%
      filter(month(Date) == month(Sys.Date()) - 1 &
                   year(Date) == year(Sys.Date())) %>%
      select(Average) %>%
      t() %>% mean () %>% round(digits = 2)

paste(month(Sys.Date()) - 1, "-", year(Sys.Date()), " KPI : ", AOI_KPI, " PPM", sep = "")


###############
# Plotting for the system month - 1

ind <- AOI_df %>%
      filter(month(Date) == month(Sys.Date()) - 1 &
                   year(Date) == year(Sys.Date())) %>%
      select(PartNumber) %>%
      unique() %>% t()

AOI_plot <- AOI_df %>%
      filter(PartNumber %in% ind) %>%
      select(Date, PartNumber, AsidePPM, BsidePPM)

fileplot <- paste(MAOI, ".png", sep = "")
 
png(fileplot, width = 1200, height = 830)
ggplot(AOI_plot, aes(x = Date)) +
        geom_point(aes(y = AsidePPM), colour = "blue") +
        geom_line(aes(y = AsidePPM), colour = "lightblue") +
        geom_point(aes(y = BsidePPM), colour = "green") +
        geom_line(aes(y = BsidePPM), colour = "lightgreen") +
        geom_hline(yintercept = 500, colour = "red") +
        facet_wrap(~ PartNumber) +
        ggtitle(paste(MAOI, "AOI performance", sep = " ")) +
        ylab("Defective rate in PPM")
dev.off()

# Export list of part for this month
MonthPart <- AOI_df %>%
        filter(month(Date) == month(Sys.Date()) - 1 &
                       year(Date) == year(Sys.Date())) %>%
        select(PartNumber, PartName, Quantity)

write.csv(MonthPart, file = paste(MAOI, "Part_list.csv", sep = "_"))
