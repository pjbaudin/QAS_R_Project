# Load libraries()
library(readxl)     # to import and read excel files
library(qcc)        # to plot quality-specific graphs
library(purrr)      # to use map function
library(dplyr)      # to help with data manipulation
library(tidyr)      # to tidy data
library(stringr)    # to clean data
library(lubridate)  # to tidy the date and time

# Create path the file
path <- "20161215 SPC.xlsx"
SPC_sheet_nbr <- excel_sheets(path)

#create function to import excel sheets and plot each
XL_analysis <- function (path_name, sheet_nbr) {
  # import data from path into data frame for analysis
  USB_SPC <- read_excel(path = path_name, sheet = sheet_nbr)
  USB_SPC_df <- as.data.frame(USB_SPC)
  
  #check USB_SPC_data
  glimpse(USB_SPC_df)
  
  # Plot time vs AB dimension
  #plot(USB_SPC_df$AB ~ USB_SPC_df$Time)
  #boxplot(USB_SPC_df$AB ~ USB_SPC_df$Time)
  
  # Use qcc package to further analysis for AB
  USB_SPC_AB <- qcc.groups(USB_SPC_df$AB, USB_SPC_df$Time)
  obj_AB <- qcc(USB_SPC_AB, type="xbar", limits =c(21.60, 21.85), center = 21.80)
  summary(obj_AB)
  process.capability(obj_AB, spec.limits = c(21.60, 21.85), target = 21.80)
  
  USB_SPC_CD <- qcc.groups(USB_SPC_df$CD, USB_SPC_df$Time)
  obj_CD <- qcc(USB_SPC_CD, type="xbar", limits = c(21.35, 21.60), center = 21.40)
  summary(obj_CD)
  process.capability(obj_CD, spec.limits = c(21.35, 21.60), target = 21.40)

}

map(excel_sheets(path), XL_analysis, path_name = path)
#SUMMARY <- map((map(excel_sheets(path), read_excel, path = path)), summary)

#write.table(SUMMARY, "SPC data summary.csv")