# Preanalysis of AOI reports
tic()

# Load libraries for SQL query and connection to SAP database
library(RODBC)      # to query SAP database
library(DBI)        # DBI convention for MySQL

###########################
# Import data from SAP (MO number and details)

con <- odbcDriverConnect("ODBC;Description=SBO;DRIVER=SQL Server Native Client 11.0;SERVER=DB0MRL82;
                         UID=evoltENG;PWD=Engineering123;DATABASE=ITesting;ApplicationIntent=READONLY;")

#List table on database
odbcGetInfo(con)

SQL <- paste("SELECT
             OWOR.DocNum, OWOR.PlannedQTY, OWOR.ItemCode, OITT.Code, OITM.ItemCode, OITM.ItemName
             FROM
             OWOR
             INNER JOIN OITT ON OWOR.ItemCode = OITT.Code
             INNER JOIN OITM ON OITT.Code = OITM.ItemCode")

# Get results from query
SAPres <- sqlQuery(con, SQL)

odbcClose(con)

SAPres <- select(SAPres, DocNum, PlannedQTY, ItemCode, ItemName)
colnames(SAPres) <- c("MONumber", "MOQuantity", "ItemCode", "ItemName")

setwd("C:/Users/PB/SkyDrive/DG Evolt/QASsvn/AOI Report")

# Load libraries
library(readxl)
library(stringr)
library(dplyr)

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

# Replace NA with 0 quantity
defectA[is.na(defectA)] <- 0

# Add PCB side and MO number
defectA$PCBside <- as.factor("A")
defectA$MONumber <- dfinfo$MONumber[1]

# Calculate the PPM rate for side A
APPM <- round(sum(defectA[, -c(1, 20, 21)], na.rm = TRUE)
      / (as.numeric(dfinfo$Quantity[1]) * as.numeric(dfinfo$DotQuantity[1]))
      * 1000000, digits = 2)

# Add PPM result to SAPres for A side
SAPres$A_PPM[SAPres$MONumber == dfinfo$MONumber[1]] <- APPM

# for side B
defectB <- defect[23:44, c(-1, -14)]
colnames(defectB) <- c("Location", "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10",
                       "A11", "A12", "P1", "P2", "P3", "P4", "P5", "P6")
ind <- c("A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10",
         "A11", "A12", "P1", "P2", "P3", "P4", "P5", "P6")
defectB[ind] <- lapply(defectB[ind], as.numeric)

# remove NA observations (no location)
defectB <- defectB[!is.na(defectB$Location), ]

# Add PCB side and MO number
if (is.na(dfinfo$MONumber[2]) == FALSE) {
        defectB$MONumber <- dfinfo$MONumber[2]
        defectB$PCBside <- as.factor("B")
        # Replace NA with 0 quantity
        defectB[is.na(defectB)] <- 0
        
        # Calculate the PPM rate for side A
        BPPM <- round(sum(defectB[, -c(1, 20, 21)], na.rm = TRUE)
                      / (as.numeric(dfinfo$Quantity[2]) * as.numeric(dfinfo$DotQuantity[2]))
                      * 1000000, digits = 2)
        # Add PPM result to SAPres for B side
        SAPres$B_PPM[SAPres$MONumber == dfinfo$MONumber[2]] <- suppressWarnings(BPPM)
        
}

# Bind the defect data set
MOdefect <- suppressWarnings(bind_rows(defectA, defectB))
ind <- c("PCBside", "MONumber", "Location")
MOdefect[ind] <- lapply(MOdefect[ind], as.factor)


toc()