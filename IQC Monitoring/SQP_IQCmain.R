library(purrr)      # to use map function
library(dplyr)      # to help with data manipulation
library(tidyr)      # to tidy data
library(stringr)    # to clean data
library(lubridate)  # to tidy the date and time
library(magrittr)   # to use pipe operator
library(tictoc)     # To measure computation time
library(readxl)     # to import and read excel files
library(plotly)     # interactive plot
library(qcc)        # quality analysis and plot
library(ggplot)     # Plot package

#############################################################################################


#############################################################################################
# Import IQC table and cleanup
#############################################################################################
tic()

# Set working directory
setwd("C:/Users/PB/SkyDrive/DG Evolt/QASsvn/IQC Monitoring");
getwd()

filename <- "IQC_defect_table.xlsx"

# Get sheet number form the file

SPC_sheet_nbr <- excel_sheets(filename)

#Import sheet 1
IQC_df <- read_excel("IQC_defect_table.xlsx", col_types = c("date", "text", "text", "text", "text", "text",
                                          "text", "text", "text", "text", "numeric", "text",
                                          "text","text", "text", "text", "numeric", "numeric",
                                          "text"))
#check dimension of IQC_df0
dim_IQC_df <- dim(IQC_df) ;
#remove last row from excel table (calculation row from excel)
IQC_df <- as.data.frame(IQC_df[-(dim_IQC_df[1] - 1), ])

# Clean data as factor
c_names <- names(IQC_df)
colfact <- c_names[c(2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16, 19)]
IQC_df[ , colfact] <- map(IQC_df[ , colfact], factor)
#Convert Date into Year - Month - Day
IQC_df$Date <- ymd(IQC_df$Date)

#Rename column
IQC_df <- rename(IQC_df, DocNum = `SAP PO / MO no.`);
IQC_df <- rename(IQC_df, ItemCode = `Part number`);

# check dataframe and summary
#glimpse(IQC_df)
summary(IQC_df)

toc()
##########################################################################################
# End of import and cleaning
##########################################################################################



##########################################################################################
# IQC_Final input date for monthly analysis
##########################################################################################

## December results
IQC_Dec <- IQC_df %>%
      filter(Date >= "2016-12-01" & Date <= "2016-12-31")

## Define weight for lot results
Weight <- data.frame(Result = c("Accepted", "Deviation", "Rejected"), Weight = c(2, 4, 8))
      
IQC_Dec <- left_join(IQC_Dec, Weight, by = "Result")

IQC_Dec %>% mutate(Ratio = Severity * Weight) %>%
      group_by(Result, Supplier) %>%
      tally() %>%
      arrange(desc(Ratio))
head(IQC_Dec)

ggplot(IQC_Dec, aes(Result, Severity)) + geom_point() + facet_grid(.~ Supplier)