# Load libraries
library(lubridate)
library(dplyr)
library(stringr)
library(reshape2)
library(ggplot2)
library(purrr)

# Set working directory
setwd("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data/ORT_Gen3_TestResults")

# InputFile is the list of processed MO
InputFile <- list.files(pattern = "Gen3_ProcessedMO.csv")

# OutputDir is the list of folders with MO data
OutputDir <- list.dirs()
# Subset OutputDir to keep only the MO number
OutputDir <- str_subset(OutputDir, "MO")
OutputDir <- as.factor(str_sub(OutputDir, start = 3L, end = 11L))

# Only process the MO that have no analysis
if(length(InputFile) == 0) {
        MOtoProcess <- as.character(OutputDir)
} else {
        ProcessedMO <- read.csv(InputFile)
        MOtoProcess <- as.character(OutputDir[!(OutputDir %in% ProcessedMO$x)])
}

Data_Processing <- function(MONumber) {
        MONumber <- as.character(MONumber)
        subdir <- str_subset(list.dirs(), MONumber)
        FileName <- tolower(list.files(path = subdir))
        FileName <- paste(subdir, str_subset(FileName, ".csv"), sep = "/")
        
        df <- read.csv(FileName)
        
        df$DateTime <- ymd_hms(strptime(df$DateTime, "%Y/%m/%d %H:%M"))
        
        names(df) <- c("DateTime", "SerialNumber", "InputVoltage", "InputCurrent", "InputPower",
                       "PowerFactor", "BatteryLevel")
        
        ORTplot <- ggplot(df, aes(x = DateTime, y = BatteryLevel)) +
                geom_point(aes(colour = SerialNumber), alpha = 0.6, size = 0.8) +
                geom_line(aes(colour = SerialNumber)) +
                facet_wrap(~ SerialNumber) +
                ggtitle(paste("Charging level as function of time - ", MONumber, sep = ""))
        
        FileName <- paste(subdir, paste(MONumber, "-ORT_BatteryChargeLevel.png", sep = ""), sep = "/")
        ggsave(FileName, ORTplot)
        
        FileName <- paste(subdir, paste(MONumber, "-ORT_InputVoltage.png", sep = ""), sep = "/")
        ORTplot <- ggplot(df, aes(x= DateTime, y = InputVoltage)) +
                geom_jitter(alpha = 0.6) +
                geom_smooth(method = "loess")
        ggsave(FileName, ORTplot)
        
        FileName <- paste(subdir, paste(MONumber, "-ORT_InputCurrent.png", sep = ""), sep = "/")
        ORTplot <- ggplot(df, aes(x= DateTime, y = InputCurrent)) +
                geom_jitter(alpha = 0.6) +
                geom_smooth(method = "loess")
        ggsave(FileName, ORTplot)
        
        FileName <- paste(subdir, paste(MONumber, "-ORT_InputPower.png", sep = ""), sep = "/")
        ORTplot <- ggplot(df, aes(x= DateTime, y = InputPower)) +
                geom_jitter(alpha = 0.6) +
                geom_smooth(method = "loess")
        ggsave(FileName, ORTplot)
        
        FileName <- paste(subdir, paste(MONumber, "-ORT_PowerFactor.png", sep = ""), sep = "/")
        ORTplot <- ggplot(df, aes(x= DateTime, y = PowerFactor)) +
                geom_jitter(alpha = 0.6) +
                geom_smooth(method = "loess")
        ggsave(FileName, ORTplot)
        
        FileName <- paste(subdir, paste(MONumber, "-ORT_InputPowerVSVoltage.png", sep = ""), sep = "/")
        ORTplot <- ggplot(df, aes(x = InputVoltage, y = InputPower)) +
                geom_jitter(alpha = 0.5) +
                geom_smooth(method = "loess")
        ggsave(FileName, ORTplot)
        
}

map(MOtoProcess, Data_Processing)

write.csv(OutputDir, file = "Gen3_ProcessedMO.csv")


