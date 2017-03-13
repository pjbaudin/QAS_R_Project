library(tictoc)
# Preanalysis of AOI reports
tic()

# Load libraries for SQL query and connection to SAP database
library(RODBC)      # to query SAP database
library(DBI)        # DBI convention for MySQL
library(dplyr)
# Load package for plotting
library(ggplot2)
# Load package for easy melting of the dataset
library(reshape2)
library(qcc)

###########################
# Import data from SAP (MO number and details)

con <- odbcDriverConnect("ODBC;Description=SBO;DRIVER=SQL Server Native Client 11.0;SERVER=DB0MRL82;
                         UID=evoltENG;PWD=Engineering123;DATABASE=ITesting;ApplicationIntent=READONLY;")

#List table on database
odbcGetInfo(con)

SQL <- paste("SELECT
             OWOR.DocNum, OWOR.PlannedQTY, OWOR.ItemCode, OITM.ItemName
             FROM
             OWOR
             INNER JOIN OITT ON OWOR.ItemCode = OITT.Code
             INNER JOIN OITM ON OITT.Code = OITM.ItemCode")

# Get results from query
SAPres <- sqlQuery(con, SQL)

odbcClose(con)

colnames(SAPres) <- c("MONumber", "MOQuantity", "ItemCode", "ItemName")

ind <- c("MONumber", "ItemCode", "ItemName")
SAPres[ind] <- lapply(SAPres[ind], as.factor) 

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
                colnames(defectB) <- c("Location", "A1", "A2", "A3", "A4", "A5", "A6",
                                       "A7", "A8", "A9", "A10", "A11", "A12", "P1", "P2",
                                       "P3", "P4", "P5", "P6")
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
                        defectB <- data.frame("Location" = as.factor(0), "A1" = 0, "A2" = 0,
                                              "A3" = 0, "A4" = 0, "A5" = 0, "A6" = 0 ,
                                              "A7" = 0 , "A8" = 0, "A9" = 0, "A10" = 0, "A11" = 0,
                                              "A12" = 0, "P1" = 0, "P2" = 0, "P3" = 0, "P4" = 0,
                                              "P5" = 0, "P6" = 0, "PCBside" = "B",
                                              "MONumber" = dfinfo$MONumber[2])

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

# Loop to extract data for the AOI excel report using the AOIresultProcessing function
AOIResultOut <- map_df(FileList, AOIresultProcessing)

# Joining AOIResult and SAP results info
AOIResultOut <- SAPres %>%
        select(MONumber, ItemCode, ItemName) %>%
        left_join(AOIResultOut, by = "MONumber")

# Adjusting the class of the variable
ind <- c("PCBside", "MONumber", "Location")
AOIResultOut[ind] <- lapply(AOIResultOut[ind], as.factor)

# Removing MO without Location (meaning not a PCB MO or no results digitized)
ind <- !is.na(AOIResultOut$Location)

AOIResultOut <- AOIResultOut[ind, ]

# Summary of the AOI info output
summary(AOIResultOut)

# melting
AOIsum <- AOIResultOut %>%
        melt(id = c("MONumber", "ItemCode", "ItemName", "PCBside", "Location")) %>%
        filter(value > 0) %>%
        group_by(variable) %>%
        summarize(TotalDefect = sum(value)) %>%
        droplevels()

# Reorder the Location based on the number of counts
AOIsum$variable <- with(AOIsum, factor(variable, 
                                       levels = variable[order(TotalDefect, decreasing = TRUE)]))

# Rename columns
colnames(AOIsum) <- c("DefectType", "TotalDefect")

# Basic plot of the type of defect accross all PCB
ggplot(AOIsum, aes(DefectType, TotalDefect)) +
        geom_bar(aes(fill = TotalDefect), stat = "identity", position = "stack",
                 alpha = 0.6, colour = "darkgrey") +
        ggtitle("Type of defect accross all PCB produced since Jan-17") +
        xlab("Count") +
        ylab("Defect Type")

# Plot a Pareto chart using the qcc package
AOIsumdefect <- AOIsum$TotalDefect 
names(AOIsumdefect) <- as.character(AOIsum$DefectType)

pareto.chart(AOIsumdefect, main = "AOI defect since Jan-17",
             xlab = "Defect Type",
             ylab = "Frequency",
             cex.names = 0.6, las = 1)
abline(h = (sum(AOIsumdefect) * 0.8), col = "red", lwd = 4)

# Extract plotting for Ledfire Basic PCB side A and B
AOILedfireBasic <- AOIResultOut %>%
        filter(ItemCode == "607558" | ItemCode == "607560") %>%
        droplevels()

# Prepare dataset
AOILedfireBasic <- AOILedfireBasic %>%
        melt(id = c("MONumber", "ItemCode", "ItemName", "PCBside", "Location")) %>%
        filter(value > 0) %>%
        droplevels()

# Reorder the Location based on the number of counts
AOILedfireBasic$Location <- with(AOILedfireBasic, factor(Location,
                               levels = Location[order(value, decreasing = TRUE)]))

# Plot Ledfire basic by PCB side (aggregate of all the MO)
ggplot(AOILedfireBasic, aes(Location, value)) +
        geom_bar(aes(fill = variable), stat = "identity", position = "stack",
                 alpha = 0.6, colour = "darkgrey") +
        facet_grid(~ PCBside) +
        ggtitle("Defect Location and Type for LedfireBasic since Jan-17") +
        xlab("Location on PCB") +
        ylab("Count")

# Plot Ledfire basic by MO (aggregate of all the MO)
ggplot(AOILedfireBasic, aes(Location, value)) +
        geom_bar(aes(fill = variable), stat = "identity", position = "stack",
                 alpha = 0.6, colour = "darkgrey") +
        facet_wrap(~ MONumber) +
        ggtitle("Defect Location and Type vs. MO number for LedfireBasic since Jan-17") +
        xlab("Location on PCB") +
        ylab("Count")

# Creating function to plot all the PCB in the AOIresult dataset

# Collect the unique part number as indices to be processed in the plot loop
ind <- unique(AOIResultOut$ItemCode)

# Function to plot the AOI results
AOIPCBplot <- function(PartNumber) {
        # subset the AOIResultOut dataset with the part number to be plotted
        AOIItem <- AOIResultOut[AOIResultOut$ItemCode == PartNumber, ]
        
        # Ignore PCB without defect reported
        if (sum(AOIItem$Location != 0)) {
                # Set graph title
                GraphTitle <- paste("Defect Location and Type for",
                            AOIResultOut[AOIResultOut$ItemCode == PartNumber, 3],
                            "since Jan-17", sep = "")
                # Prepare the dataset
                df <- AOIItem %>% 
                        melt(id = c("MONumber", "ItemCode", "ItemName",
                                    "PCBside", "Location")) %>%
                        filter(value > 0) %>%
                        droplevels()
                # Reorder the location by number of defect counted
                df$Location <- with(df, factor(Location,
                                       levels = Location[order(value, decreasing = TRUE)]))
                
                # Plot the dataset for this item (aggregate of all MO defects)
                PCBplt <- ggplot(df, aes(Location, value)) +
                        geom_bar(aes(fill = variable), stat = "identity",
                                 position = "stack", alpha = 0.6, colour = "darkgrey") +
                        ggtitle(GraphTitle) +
                        xlab("Location on PCB") +
                        ylab("Count")
                # Set the plot filename
                FileName <- paste("AOIplot_", PartNumber, ".png", sep = "")
                # Set the working directory to save the plots
                setwd("C:/Users/PB/Desktop/AOIplots")
                # Save the plot
                ggsave(FileName, PCBplt)
                
                # Alternative plot if the number of MO is more than one
                if (length(unique(df$MONumber)) != 1) {
                        # Set the graph title
                        GraphTitle <- paste("Defect Location and Type vs. MO number ",
                                    AOIResultOut[AOIResultOut$ItemCode == PartNumber, 3],
                                    "since Jan-17", sep = "")
                        # Plot the graph for this part number, wrapped by MO number
                        PCBplt <- ggplot(df, aes(Location, value)) +
                                geom_bar(aes(fill = variable), stat = "identity",
                                         position = "stack", alpha = 0.6, colour = "darkgrey") +
                                facet_wrap(~ MONumber) +
                                ggtitle(GraphTitle) +
                                xlab("Location on PCB") +
                                ylab("Count")
                        # Set the graph filename
                        FileName <- paste("AOIplot_byMONumber_", PartNumber, ".png", sep = "")
                        # Set the working directory to save the plots
                        setwd("C:/Users/PB/Desktop/AOIplots")
                        # Save the file
                        ggsave(FileName, PCBplt)
                        # Restore the working intial directory
                        setwd("C:/Users/PB/SkyDrive/DG Evolt/QASsvn/AOI Report")
                
                }
        # Restore the working intial directory
        setwd("C:/Users/PB/SkyDrive/DG Evolt/QASsvn/AOI Report")
        
        }
        
}

# Loop of the part number using the plotting function
map(ind, AOIPCBplot)

# Isolate analysis of a certain type of defect
AOI_A6_MissingParts <- AOIResultOut[AOIResultOut$A6 > 0, ]

# Result on aggregate (all MO included)
AOI_A6_MissingParts <- AOI_A6_MissingParts %>%
        select(ItemCode, ItemName, Location, A6, MONumber) %>%
        group_by(ItemCode, ItemName) %>%
        summarise(Totaldefect_A6 = sum(A6)) %>%
        arrange(desc(Totaldefect_A6))

toc()