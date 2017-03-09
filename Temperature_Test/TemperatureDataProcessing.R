# R-script to output temperature test for USB charger product

# Load packages
library(ggplot2)
library(lubridate)
library(reshape2)
library(stringr)
library(dplyr)

# set workign directory
setwd("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data/USB_Temperature test")

# Import data set
FileName <- "4000637 USB temperature test 170305.csv"
Tempdf <- read.csv(FileName)

# Grep the MO number
MONumber <- str_sub(FileName, start = 1L, end = 7L)

# Rename variables
names(Tempdf) <- c("Time", "Probe1", "TempInCase", "Probe2", "TempInChamber")

# Convert Time into proper class
Tempdf$Time <- ymd_hm(Tempdf$Time)

# Rescale the time to actual test date
StartTime <- Sys.time() - ymd_hm("2017-02-28 00:00")
duration <- Sys.time() - Tempdf$Time[1] - StartTime

Tempdf$Time <- Tempdf$Time + duration

Tempdf <- select(Tempdf, Time, TempInChamber, TempInCase, Probe1, Probe2)

# summary of data set
summary(Tempdf)

# Melt data for easy plotting
Tempdfplot <- melt(Tempdf, id = "Time") 

# Plot data set
USBTempGraph <- ggplot(Tempdfplot, aes(x = Time, y = value, colour = variable)) +
      geom_line(lwd = 1.5) +
      geom_hline(yintercept = 85, colour = "red", lwd = 1.5) +
      scale_y_continuous(breaks = seq(0, 100, 5)) +
      xlab("Time") +
      ylab("Temperature in degree C") +
      ggtitle(paste("MO ", MONumber, " - 30USBCM Charger Temperature Test - ", Sys.Date(), sep = ""))

FileName <- paste(MONumber, "_USBTemperatureTest.png", sep = "")
ggsave(FileName, USBTempGraph)  
