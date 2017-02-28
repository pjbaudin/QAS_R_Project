# RJ warehouse data processing

setwd("C:/Users/PB/Desktop/Evolt People Download/Sarah/2017-02 February")

library(readxl)

FileName <- "RJ.xlsx"
RJ_df <- read_excel(FileName, col_names = FALSE)

names(RJ_df) <- c("PartNumber", "Quantity", "Notes")

RJ_df$PartNumber <- as.factor(RJ_df$PartNumber)

RJ_sum <- RJ_df %>%
        filter(Notes == "-") %>%
        group_by(PartNumber) %>%
        summarise(TotalQty = sum(Quantity)) %>%
        arrange(desc(TotalQty)) %>%
        droplevels()

write.csv(RJ_sum, file = "RJ_count_final.csv")
        
