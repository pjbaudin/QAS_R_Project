# Load libraries()
library(readxl)     # to import and read excel files
library(qcc)        # to plot quality-specific graphs
library(purrr)      # to use map function
library(dplyr)      # to help with data manipulation
library(tidyr)      # to tidy data
library(stringr)    # to clean data
library(lubridate)  # to tidy the date and time

# Set working directory
setwd("D:/Data Science/QAS_R_Project/SPC_dimension")

# Create path the file
path <- "SPC_dataset_2017.xlsx"
SPC_sheet_nbr <- excel_sheets(path)

#create function to import excel sheets and plot each
XL_analysis <- function (path_name, sheet_nbr) {
  # import data from path into data frame for analysis
  USB_SPC <- read_excel(path = path_name, sheet = sheet_nbr)
  USB_SPC <- as.data.frame(USB_SPC)
  
  #check USB_SPC_data
  #glimpse(USB_SPC_df)
  
  # Plot time vs AB dimension
  #plot(USB_SPC_df$AB ~ USB_SPC_df$Time)
  #boxplot(USB_SPC_df$AB ~ USB_SPC_df$Time)
  
  # Use qcc package to further analysis for AB
  # Dimensions and analysis for specification of the AB dimensions
  USB_SPC_AB <- qcc.groups(USB_SPC$AB, USB_SPC$Time)
  GraphTitle <- paste("30USBCM - Xbar chart for AB ", sheet_nbr)
  obj_AB <- qcc(USB_SPC_AB, type="xbar", limits =c(21.60, 21.85), center = 21.80, title = GraphTitle)

  # Create JPEG file name using sheet name
  FileName <- paste("AB_", sheet_nbr, ".jpeg")
  # Output a JPEG of the graph for CD
  jpeg(file = FileName, width = 800, height = 600)
  plot(obj_AB)
  dev.off()
  
  # Process capability for AB dimension
  # Create JPEG file name using sheet name
  FileName <- paste("AB_ProcessReport_", sheet_nbr, ".jpeg")
   # Output a JPEG of the graph for CD process capability
  jpeg(file = FileName, width = 800, height = 600)
  process.capability(obj_AB, spec.limits = c(21.60, 21.85), target = 21.80)
  dev.off()
  
  
  # Dimensions and analysis for specification of the CD dimensions
  USB_SPC_CD <- qcc.groups(USB_SPC$CD, USB_SPC$Time)
  GraphTitle <- paste("30USBCM - Xbar chart for CD ", sheet_nbr)
  obj_CD <- qcc(USB_SPC_CD, type="xbar", limits = c(21.35, 21.60), center = 21.40, title = GraphTitle)
  
  # Create JPEG file name using sheet name
  FileName <- paste("CD_", sheet_nbr, ".jpeg")
  
  # Output a JPEG of the graph for CD
  jpeg(file = FileName, width = 800, height = 600)
  plot(obj_CD)
  dev.off()
  
  # Process capability for CD dimension
  # Create JPEG file name using sheet name
  FileName <- paste("CD_ProcessReport_", sheet_nbr, ".jpeg")
  # Output a JPEG of the graph for CD process capability
  jpeg(file = FileName, width = 800, height = 600)
  process.capability(obj_CD, spec.limits = c(21.35, 21.60), target = 21.40)
  dev.off()

  # Create CSV file name using sheet name
  # FileName <- paste("CD_Summary", sheet_nbr, ".csv")
  # Report <- summary(obj_CD)
  # # Write data in a summary .csv file
  # write.table(Report, file = FileName)
  # /!\ This is not working at the moment
  # => need to study the structure of the output of qcc package
}

# Loop over the excell sheet with data to perform analysis
# using the function XL_analysis
map(excel_sheets(path), XL_analysis, path_name = path)

