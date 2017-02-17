# R-script to output temperature test for USB charger product

# Load packages
library(ggplot2)
library(lubridate)
library(reshape2)
library(stringr)

# set workign directory
setwd("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data/USB_Temperature test")

# Import data set
FileName <- "4000529 USB temperature test 170215.csv"
Tempdf <- read.csv(FileName)

# Grep the MO number
MONumber <- str_sub(FileName, start = 1L, end = 7L)

# Rename variables
names(Tempdf) <- c("Time", "TempInChamber", "TempInCase", "Probe1", "Probe2")

# Convert Time into proper class
Tempdf$Time <- ymd_hm(Tempdf$Time)

# Rescale the time to actual test date
StartTime <- Sys.time() - ymd_hm("2017-02-15 00:00")
duration <- Sys.time() - Tempdf$Time[1] - StartTime

Tempdf$Time <- Tempdf$Time + duration

# summary of data set
summary(Tempdf)

# Melt data for easy plotting
Tempdfplot <- melt(Tempdf, id = "Time") 

# Plot data set
ggplot(Tempdfplot, aes(x = Time, y = value, colour = variable)) +
      geom_line(lwd = 1.5) +
      geom_hline(yintercept = 85, colour = "red", lwd = 1.5) +
      xlab("Time") +
      ylab("Temperature in degree C") +
      ggtitle(paste("MO ", MONumber, " - 30USBCM Charger Temperature Test - ", Sys.Date(), sep = ""))
      
