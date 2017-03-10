library(tictoc)
# Preanalysis of AOI reports
tic()

# Load libraries for SQL query and connection to SAP database
library(RODBC)      # to query SAP database
library(DBI)        # DBI convention for MySQL
library(dplyr)

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
library(purrr)

# import list of files to process
FileList <- list.files()

# Setup dataframe to store the final result
AOIResultOut <- data.frame()

# Creating function for importing and cleaning the data
AOIresultProcessing <- function(FileName) {

        #  Import report information
        dfinfo <- read_excel(FileName, skip = 2)
        dfinfo <- dfinfo[1:2, c(1, 5, 6, 8, 9, 14)]
        
        colnames(dfinfo) <- c("Product", "PCBSide", "Quantity", "DotQuantity", "TotalDotQuantity", "MONumber")
        
        dfinfo$MONumber <- str_sub(dfinfo$MONumber, start = 1L, end = 7L)
        
        # Import report defect summary side A
        defectA <- read_excel(FileName, skip = 6)
        
        # for side A
        defectA <- defectA[1:21, c(-1, -15)]
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
        if (dim(defectA)[1] != 0) {
                defectA$PCBside <- as.factor("A")
                defectA$MONumber <- dfinfo$MONumber[1]
        } else {
                defectA <- data.frame("Location" = as.factor(0), "A1" = 0, "A2" = 0, "A3" = 0,
                                      "A4" = 0, "A5" = 0, "A6" = 0 , "A7" = 0 , "A8" = 0,
                                      "A9" = 0, "A10" = 0, "A11" = 0, "A12" = 0, "P1" = 0,
                                      "P2" = 0, "P3" = 0, "P4" = 0, "P5" = 0, "P6" = 0,
                                      "PCBside" = "A", "MONumber" = dfinfo$MONumber[1])
        }
        
        # Calculate the PPM rate for side A
        APPM <- round(sum(defectA[, -c(1, 20, 21)], na.rm = TRUE)
                      / (as.numeric(dfinfo$Quantity[1]) * as.numeric(dfinfo$DotQuantity[1]))
                      * 1000000, digits = 2)
        
        # Add PPM result to SAPres for A side
        SAPres$A_PPM[SAPres$MONumber == dfinfo$MONumber[1]] <- APPM
        
        # Add Dot Quantity for A side
        SAPres$A_DotQty[SAPres$MONumber == dfinfo$MONumber[1]] <- dfinfo$DotQuantity[1]
        
        defectB <- data.frame()
        
        # Add PCB side and MO number
        if (is.na(dfinfo$MONumber[2]) == FALSE) {
                # Import report defect summary side B
                defectB <- read_excel(FileName, skip = 28)
                
                # for side B
                defectB <- defectB[1:20, c(-1, -15)]
                colnames(defectB) <- c("Location", "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10",
                                       "A11", "A12", "P1", "P2", "P3", "P4", "P5", "P6")
                # reclass the data
                defectB$Location <- as.factor(defectB$Location)
                ind <- c("A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10",
                         "A11", "A12", "P1", "P2", "P3", "P4", "P5", "P6")
                defectB[ind] <- lapply(defectB[ind], as.numeric)
                
                # remove NA observations (no location)
                defectB <- defectB[!is.na(defectB$Location), ]
                
                # Replace NA with 0 quantity
                defectB[is.na(defectB)] <- 0
                
                # Add PCB side and MO number
                if (dim(defectB)[1] != 0) {
                        defectB$MONumber <- dfinfo$MONumber[2]
                        defectB$PCBside <- as.factor("B")
                } else {
                        defectB <- data.frame("Location" = as.factor(0), "A1" = 0, "A2" = 0, "A3" = 0,
                                              "A4" = 0, "A5" = 0, "A6" = 0 , "A7" = 0 , "A8" = 0,
                                              "A9" = 0, "A10" = 0, "A11" = 0, "A12" = 0, "P1" = 0,
                                              "P2" = 0, "P3" = 0, "P4" = 0, "P5" = 0, "P6" = 0,
                                              "PCBside" = "B", "MONumber" = dfinfo$MONumber[2])

                }
                
                # Calculate the PPM rate for side A
                BPPM <- round(sum(defectB[, -c(1, 20, 21)], na.rm = TRUE)
                              / (as.numeric(dfinfo$Quantity[2]) * as.numeric(dfinfo$DotQuantity[2]))
                              * 1000000, digits = 2)
                # Add PPM result to SAPres for B side
                SAPres$B_PPM[SAPres$MONumber == dfinfo$MONumber[2]] <- suppressWarnings(BPPM)
                
                # Add Dot Quantity for B side
                SAPres$B_DotQuantity[SAPres$MONumber == dfinfo$MONumber[2]] <- dfinfo$DotQuantity[2]
                
        }
        
        # Bind the defect data set
        df <- suppressWarnings(bind_rows(defectA, defectB))
        
        # Final bind to summary df
        AOIResultOut <- suppressWarnings(bind_rows(AOIResultOut, df))

}

AOIResultOut <- map_df(FileList, AOIresultProcessing)

ind <- c("PCBside", "MONumber", "Location")
AOIResultOut[ind] <- lapply(AOIResultOut[ind], as.factor)

summary(AOIResultOut)

library(reshape2)

df <- melt(AOIResultOut, id = c("MONumber", "PCBside", "Location"))

library(ggplot2)

ggplot(df, aes(variable, value)) +
        geom_bar(aes(fill = variable), stat = "identity", position = "stack",
                 alpha = 0.6, colour = "darkgrey")

toc()