# Preanalysis of AOI reports
tic()
setwd("C:/Users/PB/SkyDrive/DG Evolt/QASsvn/AOI Report")

# Load libraries
library(readxl)
library(stringr)

# import data

FileName <- "AOI-1706-01 MO_4000492.xlsx"

#  Import report information
df <- read_excel(FileName, skip = 2)
dfinfo <- df[1:2, c(1, 5, 6, 8, 9, 14)]

colnames(dfinfo) <- c("Product", "PCBSide", "Quantity", "DotQuantity", "TotalDotQuantity", "MONumber")

dfinfo$MONumber <- str_sub(dfinfo$MONumber, start = 1L, end = 7L)

# Import report defect summary
defect <- read_excel(FileName, skip = 6 )

# for side A
defectA <- defect[1:21, c(-1, -14)]
colnames(defectA) <- c("Location", "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10",
                       "A11", "A12", "P1", "P2", "P3", "P4", "P5", "P6")
# reclass the data
defectA$Location <- as.factor(defectA$Location)
ind <- c("A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10",
         "A11", "A12", "P1", "P2", "P3", "P4", "P5", "P6")
defectA[ind] <- lapply(defectA[ind], as.numeric)

# remove NA observations (no location)
defectA <- defectA[!is.na(defectA$Location), ]

# Calculate the PPM rate for side A
round(sum(defectA[, -1], na.rm = TRUE)
      / (as.numeric(dfinfo$Quantity[1]) * as.numeric(dfinfo$DotQuantity[1]))
      * 1000000, digits = 2)

# for side B
defectB <- defect[23:44, c(-1, -14)]
colnames(defectB) <- c("Location", "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10",
                       "A11", "A12", "P1", "P2", "P3", "P4", "P5", "P6")
ind <- c("A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10",
         "A11", "A12", "P1", "P2", "P3", "P4", "P5", "P6")
defectB[ind] <- lapply(defectB[ind], as.numeric)

# remove NA observations (no location)
defectB <- defectB[!is.na(defectA$Location), ]

# Calculate the PPM rate for side A
round(sum(defectB[, -1], na.rm = TRUE)
      / (as.numeric(dfinfo$Quantity[2]) * as.numeric(dfinfo$DotQuantity[2]))
      * 1000000, digits = 2)
toc()