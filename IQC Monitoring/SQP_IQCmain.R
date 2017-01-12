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
library(ggplot2)     # Plot package


#############################################################################################


#############################################################################################
# Import IQC table and cleanup
#############################################################################################
tic()

# Set working directory (for future use)
# setwd("C:/Users/PB/SkyDrive/DG Evolt/QASsvn/IQC Monitoring");
# getwd()

# Test directory
setwd("D:/Data Science/QAS_R_Project/IQC Monitoring")

filename <- "IQC_defect_table_update.xlsx"

# Get sheet number form the file

SPC_sheet_nbr <- excel_sheets(filename)

#Import sheet 1
IQC_df <- read_excel(filename, col_types = c("date", "text", "text", "text", "text", "text",
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
IQC_df <- rename(IQC_df, IssueCat = `Issue cat.`);


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
## In the future use a define .csv table to set the result weight

Weight <- data.frame(Result = c("Accepted", "Deviation", "Rejected"), Weight = c(2, 4, 8))

# Combine dataset with Result weight
IQC_Dec <- left_join(IQC_Dec, Weight, by = "Result")


## Define weight for category
## In the future use a define .csv table to set the category weight
CatWeight <- data.frame( IssueCat = c("Appearance", "Construction",
                                      "Dimension", "Engineering",
                                      "Function", "Marking", "Material", "Packaging",
                                      "PPAP"),
                         CatWeight = c(2, 5, 3, 1, 5, 3, 3, 2, 1))
# Combine dataset with category weigth
IQC_Dec <- left_join(IQC_Dec, CatWeight, by = "IssueCat")


TallyDec16 <- IQC_Dec %>% 
      mutate(Ratio = Severity * Weight * CatWeight) %>%
      group_by(Result, Supplier, Ratio) %>%
      tally() %>%
      ungroup()

Top15Dec_Occurence <- TallyDec16 %>%
      arrange(desc(n), Ratio) %>%
      head(n = 15)

Top15Dec_Occurence

Top15Dec_Ratio <- TallyDec16 %>%
      arrange(desc(Ratio), n) %>%
      head(n = 15)

Top15Dec_Ratio

# compute teh intersection of the 2 calculated top 15
intersect(Top15Dec_Occurence$Supplier, Top15Dec_Ratio$Supplier)

# Thik about the output as the list of suppliers
# that need to be watched and monitored for the following month

<<<<<<< HEAD

## Supplier list

filename <- "Supplierlistcc.xlsx"

#Import sheet 1
Supplier <- read_excel(filename, col_types = c("text", "text", "text"))
Supplier$CardCode <- as.factor(Supplier$CardCode)
summary(Supplier)
=======
>>>>>>> caaefd5465aabc508a21703839692c037d4f85f2
