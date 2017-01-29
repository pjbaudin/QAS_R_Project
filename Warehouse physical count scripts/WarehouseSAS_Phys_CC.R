# Load libraries
library(readxl)
library(dplyr)

# Set working directory
setwd("C:/Users/PB/Desktop/WH-SAP")

# Import tables
PhysCheck <- read_excel("Physical count 2017-01-24 1352PM.xlsx")

SAPCheck <- read_excel("blcl 20170124.xlsx")

# Rename column names
names(SAPCheck) <- c("BinLocation", "PartNumber", "PartName", "SAPQuantity")
names(PhysCheck) <- c("PartNumber", "PartName", "PhysQuantity",
                      "BinLocation", "Diff", "Check", "Occ", "SAPin")

PhysSAPmerge <- left_join(PhysCheck, SAPCheck,
                          by = c("PartNumber", "BinLocation"))

write.csv(PhysSAPmerge, file = "Physical count Merged Updated.csv")

PhysUpdate <- PhysSAPmerge[is.na(PhysSAPmerge$SAPQuantity), ] %>%
        select(-PartName.y)

write.csv(PhysUpdate, file = "Physical count Updated.csv")
