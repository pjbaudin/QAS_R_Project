library(dplyr)

filename <- list.files()

binlistsum <- data.frame()

for (i in 1:length(filename)) {
        namedf <- filename[i]
        df <- read.csv(namedf)
        df <- select(df, BinLocation, PartNumber, sum.Quantity.)
        df$BinLocation <- as.factor(df$BinLocation)
        df$PartNumber <- as.factor(df$PartNumber)
        df$sum.Quantity. <- as.numeric(df$sum.Quantity.)
        binlistsum <- bind_rows(binlistsum, df)
}

binlistsum$BinLocation <- toupper(binlistsum$BinLocation)
binlistsum[binlistsum$BinLocation == "SMD ROOM UNLABELED RACK", 1] <- "UNLABELED"
ind <- grepl("^7", binlistsum$BinLocation) == TRUE
binlistsum[ind, 1] <- paste0("RM-",binlistsum[ind, 1])

binlistsum$BinLocation <- as.factor(binlistsum$BinLocation) 
binlistsum$PartNumber <- as.factor(binlistsum$PartNumber) 
binlistsum$sum.Quantity. <- as.numeric(binlistsum$sum.Quantity.)



summary(binlistsum)

write.csv(binlistsum, file = "MainListBinSummary.csv")

BinlocationList <- unique(binlistsum$BinLocation)
write.csv(BinlocationList, file = "BinLocationList.csv" )

PartNumberList <- unique(binlistsum$PartNumber)
write.csv(PartNumberList, file = "PartNumberList.csv")

