# Load libraries()
library(readxl)     # to import and read excel files
library(qcc)        # to plot quality-specific graphs
library(purrr)      # to use map function
library(dplyr)      # to help with data manipulation
library(tidyr)      # to tidy data
library(stringr)    # to clean data
library(lubridate)  # to tidy the date and time

# Set working directory
setwd("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data/USB_SPC")

# Create path the files
File_list <- list.files(pattern = ".xlsx")

# List the MO processed directory
MO_processed <- list.dirs("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data/USB_SPC")

# Extract the MO number processed
MO_processed <- MO_processed[2:length(MO_processed)]
MO_processed <- str_sub(MO_processed, start = -9L, end = -1L)

write.csv(MO_processed, file = "MO_processed_list.csv")

#create function to import excel sheets and plot each
XL_analysis <- function (path_name) {
      # Set working directory for data import
      setwd("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data/USB_SPC")
      
      # remove the file extension for naming
      MOnumber <- str_sub(path_name, start = 1L, end = -6L)
      
      # Delete existing folder
      Dir_to_remove <- paste("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data/USB_SPC/", MOnumber, sep = "")
      unlink(Dir_to_remove, recursive = TRUE)
      
      # Create a directory with MO number
      mainDir <- getwd()
      subDir <- MOnumber
      
      dir.create(file.path(mainDir, subDir), showWarnings = FALSE)
      
      # import data from path into data frame for analysis
      USB_SPC <- read_excel(path = path_name)
      USB_SPC <- as.data.frame(USB_SPC)
      
      # Set output working directory
      setwd(file.path(mainDir, subDir))
      
      # copy source file to sub directory
      File_to_copy <- paste("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data/USB_SPC/", path_name, sep = "")
      
      file.copy(File_to_copy, getwd())
      # Clean up directory
      file.remove(File_to_copy)
      
      # Use qcc package to further analysis for AB
      # Dimensions and analysis for specification of the AB dimensions
      USB_SPC_AB <- qcc.groups(USB_SPC$AB, USB_SPC$Time)
      GraphTitle <- paste("30USBCM - Xbar chart for AB - ", MOnumber)
      obj_AB <- qcc(USB_SPC_AB, type="xbar", limits =c(21.60, 21.85), center = 21.80, title = GraphTitle)

      # Create JPEG file name using sheet name
      FileName <- paste("AB_", MOnumber, ".jpeg")
      # Output a JPEG of the graph for CD
      jpeg(file = FileName, width = 800, height = 600)
      plot(obj_AB)
      dev.off()
      
      # Process capability for AB dimension
      # Create JPEG file name using sheet name
      FileName <- paste("AB_ProcessReport_", MOnumber, ".jpeg")
      # Output a JPEG of the graph for CD process capability
      jpeg(file = FileName, width = 800, height = 600)
      process.capability(obj_AB, spec.limits = c(21.60, 21.85), target = 21.80)
      dev.off()
      
      # Dimensions and analysis for specification of the CD dimensions
      USB_SPC_CD <- qcc.groups(USB_SPC$CD, USB_SPC$Time)
      GraphTitle <- paste("30USBCM - Xbar chart for CD ", MOnumber)
      obj_CD <- qcc(USB_SPC_CD, type="xbar", limits = c(21.35, 21.60), center = 21.40, title = GraphTitle)
  
      # Create JPEG file name using sheet name
      FileName <- paste("CD_", MOnumber, ".jpeg")
  
      # Output a JPEG of the graph for CD
      jpeg(file = FileName, width = 800, height = 600)
      plot(obj_CD)
      dev.off()
  
      # Process capability for CD dimension
      # Create JPEG file name using sheet name
      FileName <- paste("CD_ProcessReport_", MOnumber, ".jpeg")
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
map(File_list, XL_analysis)

MO_processed