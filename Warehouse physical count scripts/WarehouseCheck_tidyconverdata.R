# Simple script to convert warehouse .txt count to csv and remove NA

# Note: after running this script, need to update and commit the changes
# to the QASsvn folder

# load packages
library(stringr)
library(dplyr)

# set working Directory
setwd("C:/Users/PB/SkyDrive/DG Evolt/QASsvn/Warehouse check")

# List the files in the QASsvn local directory with .txt extension
File_list <- list.files(pattern = ".txt")

# Loop to tidy and convert all the txt files from checking
for(i in length(File_list)) {
      # Import data set
      df <- read.delim(File_list[1], header = FALSE, sep = "\t")
      
      # Select and rename the first three columns 
      df <- df[, 1:3]
      names(df) <- c("PartNumber", "BinLocation", "Quantity")
      
      # Basic filter to remove incorrect part number (must start with 6 or 7)
      # Can be revise to check the lenght of the part
      ind <- ind <- grepl("^[6-7]", df$PartNumber)
      df <- df[ind, ]
      
      # Create file name .csv
      File_Name <- str_sub(File_list[i], start = 1L, end = -5L)
      File_Name <- paste(File_Name, ".csv", sep = "")
      
      # Write .csv file of tidy data
      write.csv(df, file = File_Name)
      
      file.remove(File_list[i])
}