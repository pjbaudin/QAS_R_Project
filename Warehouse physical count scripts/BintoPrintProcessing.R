library(dplyr)

setwd("C:/Users/PB/Desktop/WH-SAP/Bin Summary Scanning/Bin_Summary")

mainDir <- "C:/Users/PB/Desktop/WH-SAP/Bin Summary Scanning/Bin_Summary"
subDir <- "outputDirectory"

Bintocut <- read.csv("MainListBinSummary.csv")

Bingroup <- Bintocut %>% select(BinLocation) %>% 
        unique() %>% droplevels()

dir.create(file.path(mainDir, subDir), showWarnings = FALSE)
setwd(file.path(mainDir, subDir))

BinLoc <- data.frame()

for (i in 1:length(Bingroup$BinLocation)) {
        BinLoc <- droplevels(Bingroup[i, 1])
        BinLoc <- as.character(BinLoc)
        Bintoprint <- Bintocut %>%
                filter(BinLocation == BinLoc)
        filename <- paste(BinLoc, "_toprint.csv", sep = "")
        write.csv(Bintoprint, file = filename)
}

