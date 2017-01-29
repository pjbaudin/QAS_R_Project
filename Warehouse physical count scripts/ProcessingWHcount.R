# Load packages for data cleaning
library(dplyr)
# library(lubridate)
# library(magrittr)
library(readxl)


filename <- "Judy 20170120 FG Cleaned.xlsx"

sheetNbr <- length(excel_sheets(filename))

DF <- data.frame()

#create function to import excel sheets in one dataset
for (i in 1:sheetNbr) {
        excel <- read_excel(path = filename, sheet = i,
                            col_names = TRUE)
        names(excel) <- c("PartNumber", "Quantity", "BinLocation")
        excel$PartNumber <- as.factor(excel$PartNumber)
        excel$Quantity <- as.numeric(excel$Quantity)
        excel$BinLocation<- as.factor(excel$BinLocation)
        DF <- bind_rows(DF, excel)
}

DF$PartNumber <- as.factor(DF$PartNumber)
DF$Quantity <- as.numeric(DF$Quantity)
DF$BinLocation<- as.factor(DF$BinLocation)

#summary(DF)

OutDF <- DF %>% group_by(BinLocation, PartNumber) %>%
        summarise(sum(Quantity))


# glimpse(OutDF)

outputfile <- paste(filename, "BinSummary.csv", sep = "_")

write.csv(OutDF, file = outputfile)
