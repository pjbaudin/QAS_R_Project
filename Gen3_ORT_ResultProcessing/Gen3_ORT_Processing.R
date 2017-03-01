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
        FileName <- paste(subdir,
                          str_subset(list.files(path = subdir), "PF9810"),
                          sep = "/")
        
        df <- read.csv(FileName)
        
        Bat1 <- paste(df$SerialNumber[1])
        Bat2 <- paste(df$SerialNumber.1[1])
        Bat3 <- paste(df$SerialNumber.2[1])
        Bat4 <- paste(df$SerialNumber.3[1])
        
        df <- df %>%
                select(Voltage.V., Current.A., Power.W., PowerFactor,
                     Battery1..., Battery2..., Battery3..., Battery4...) %>%
                mutate(Time = as_datetime(rep(1:nrow(df))))

        names(df) <- c("Voltage", "Current", "InputPower",
                       "PowerFactor", Bat1, Bat2, Bat3, Bat4, "Time")

        df <- melt(df, id=c("Time", "Voltage", "Current", "InputPower", "PowerFactor")) 
        
        ORTplot <- ggplot(df, aes(x = Time, y = value)) +
                geom_point(aes(colour = variable), alpha = 0.6, size = 0.8) +
                geom_line(aes(colour = variable)) +
                facet_wrap(~ variable) +
                ggtitle(paste("Charging level as function of time - ", MONumber, sep = ""))
        
        FileName <- paste(subdir, paste(MONumber, "-ORT_BatteryChargeLevel.png", sep = ""), sep = "/")
        ggsave(FileName, ORTplot)
        
        FileName <- paste(subdir, paste(MONumber, "-ORT_InputVoltage.png", sep = ""), sep = "/")
        png(FileName)
                plot(df$Voltage ~ df$Time)
        dev.off()
        
        FileName <- paste(subdir, paste(MONumber, "-ORT_InputPowerVSVoltage.png", sep = ""), sep = "/")
        ORTplot <- ggplot(df, aes(x = Voltage, y = InputPower)) +
                geom_point(alpha = 0.5)
        ggsave(FileName, ORTplot)
}

map(MOtoProcess, Data_Processing)

write.csv(OutputDir, file = "Gen3_ProcessedMO.csv")


