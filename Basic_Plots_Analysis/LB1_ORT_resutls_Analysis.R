# Draft script to plot test results
 
# Set working directory
setwd("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data/LB1 - ORT")

# load libraries
library(ggplot2)
library(readxl)

# Import data
df <- read_excel(path = "ORT-1709-09 605422 Ledfire basic.xlsx",
                 sheet = 3, skip = 1)

df$SerialNumber <- as.factor(df$SerialNumber)

FileName <- "InputPower_vs_InvputVoltage.png"
ORTplot <- ggplot(df, aes(x = `Voltage(V)`, y = `Power(W)`)) +
        geom_jitter(alpha = 0.8) +
        geom_hline(yintercept = 2.5, colour = "red", lwd = 1.5) +
        geom_hline(yintercept = 1.9, colour = "red", lwd = 1.5) +
        stat_ellipse(level = 0.95, color = "steelblue", lty = 2, lwd = 1.2) +
        stat_ellipse(level = 0.99, color = " blue", lty = 4, lwd = 1.2) +
        stat_density_2d() +
        geom_point(aes(x = mean(`Voltage(V)`), y = mean(`Power(W)`)),
                   colour = "red", size = 2, shape = 3)

ggsave(FileName, ORTplot)


FileName <- "PowerFactor_vs_InvputVoltage.png"
ORTplot <- ggplot(df, aes(x = `Voltage(V)`, y = PowerFactor)) +
        geom_jitter(alpha = 0.8) +
        stat_ellipse(level = 0.95, color = "steelblue", lty = 2, lwd = 1.2) +
        stat_ellipse(level = 0.99, color = " blue", lty = 4, lwd = 1.2) +
        stat_density_2d() +
        geom_point(aes(x = mean(`Voltage(V)`), y = mean(PowerFactor)),
                   colour = "red", size = 2, shape = 3)

ggsave(FileName, ORTplot)

FileName <- "InputCurrent_vs_InvputVoltage.png"
ORTplot <- ggplot(df, aes(x = `Voltage(V)`, y = `Current(A)`)) +
        geom_jitter(alpha = 0.8) +
        stat_ellipse(level = 0.95, color = "steelblue", lty = 2, lwd = 1.2) +
        stat_ellipse(level = 0.99, color = " blue", lty = 4, lwd = 1.2) +
        stat_density_2d() +
        geom_point(aes(x = mean(`Voltage(V)`), y = mean(`Current(A)`)),
                   colour = "red", size = 2, shape = 3)

ggsave(FileName, ORTplot)


FileName <- "PowerFactor_vs_InputPower.png"
ORTplot <- ggplot(df, aes(x = `Power(W)`, y = PowerFactor)) +
        geom_jitter(alpha = 0.8) +
        geom_vline(xintercept = 2.5, colour = "red", lwd = 1.5) +
        geom_vline(xintercept = 1.9, colour = "red", lwd = 1.5) +
        stat_ellipse(level = 0.95, color = "steelblue", lty = 2, lwd = 1.2) +
        stat_ellipse(level = 0.99, color = " blue", lty = 4, lwd = 1.2) +
        stat_density_2d() +
        geom_point(aes(x = mean(`Power(W)`), y = mean(PowerFactor)),
                   colour = "red", size = 2, shape = 3)

ggsave(FileName, ORTplot)
