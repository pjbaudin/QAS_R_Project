# Load packages
library(readxl)
library(dplyr)
library(lubridate)
library(ggplot2)

###########################
# Data import

# File path name
fileName <- "C:/Users/PB/SkyDrive/DG Evolt/QASsvn/QA Monitoring/QA report monitoring list.xlsx"

# Read the sheet names
QAMonitoring_sheet <- excel_sheets(fileName)

# Import AOI dataset
AOI_df <- suppressWarnings(read_excel(fileName, sheet = 4))


###################################################
# Set working directory
setwd("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data")

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

# Convert date in ymd
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

mainDir <- getwd()
subDir <- MAOI

if (file.exists(subDir)){
        setwd(file.path(mainDir, subDir))
} else {
        dir.create(file.path(mainDir, subDir))
        setwd(file.path(mainDir, subDir))
}


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
 
AOImonthly <- ggplot(AOI_plot, aes(x = Date)) +
        geom_point(aes(y = AsidePPM), colour = "blue") +
        geom_line(aes(y = AsidePPM), colour = "lightblue") +
        geom_point(aes(y = BsidePPM), colour = "green") +
        geom_line(aes(y = BsidePPM), colour = "lightgreen") +
        geom_hline(yintercept = 250, colour = "red") +
        facet_wrap(~ PartNumber) +
        ggtitle(paste(MAOI, "AOI performance", sep = " ")) +
        ylab("Defective rate in PPM")

ggsave(fileplot, AOImonthly)

# Export list of part for this month
MonthPart <- AOI_df %>%
        filter(month(Date) == month(Sys.Date()) - 1 &
                       year(Date) == year(Sys.Date())) %>%
        select(PartNumber, PartName, Quantity, AsidePPM, BsidePPM) %>%
        group_by(PartNumber, PartName) %>%
        summarize(TotalQuantity = sum(Quantity),
                  AsidePPM_avg = mean(AsidePPM),
                  BsidePPM_avg = mean(BsidePPM))

write.csv(MonthPart, file = paste(MAOI, "AOI_Resutlts_Part_list.csv", sep = "_"))

# Export list for investigation
MonthPart_Inv <- AOI_df %>%
        filter(month(Date) == month(Sys.Date()) - 1 &
                       year(Date) == year(Sys.Date())) %>%
        select(MONumber, PartNumber, PartName, Quantity, AsidePPM, BsidePPM) %>%
        group_by(MONumber, PartNumber, PartName) %>%
        summarize(TotalQuantity = sum(Quantity),
                  AsidePPM_avg = mean(AsidePPM),
                  BsidePPM_avg = mean(BsidePPM)) %>%
        filter(AsidePPM_avg > 250 | BsidePPM_avg > 250)

write.csv(MonthPart_Inv, file = paste(MAOI, "AOI_Resutlts_toInvestigate.csv", sep = "_"))
