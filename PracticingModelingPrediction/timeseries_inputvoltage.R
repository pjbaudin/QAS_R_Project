# Load libraries
library(lubridate)
library(dplyr)
library(stringr)
library(reshape2)
library(ggplot2)
library(purrr)
library(xts)

# Set working directory
setwd("C:/Users/PB/SkyDrive/DG Evolt/QAS_Data/ORT_Gen3_TestResults")

# Practice prediction on a Input voltage dataset
subdir <- "./MO4000554"
FileName <- tolower(list.files(path = subdir))
FileName <- paste(subdir, str_subset(FileName, ".csv"), sep = "/")

df <- read.csv(FileName)

df$DateTime <- strptime(df$DateTime, "%d/%m/%Y %H:%M")

names(df) <- c("DateTime", "SerialNumber", "InputVoltage", "InputCurrent", "InputPower",
               "PowerFactor", "BatteryLevel")


# dfts <- xts(df$InputVoltage, df$DateTime)
dfts <- as.ts(df$InputVoltage, start = c(df$DateTime, 1), frequency = 24)

# 
Voltage_fit <- arima(dfts, order = c(1, 0, 0))

predict_Voltage <- predict(Voltage_fit)

acf(dfts)

plot(dfts)

Voltage_forecast <- predict(Voltage_fit, n.ahead = 200)$pred
Voltage_forecast_se <- predict(Voltage_fit, n.ahead = 200)$se

points(Voltage_forecast, type = "l", col = 2)
points(Voltage_forecast - 2*Voltage_forecast_se, type = "l", col = 2, lty = 2)
points(Voltage_forecast + 2*Voltage_forecast_se, type = "l", col = 2, lty = 2)

# Moving Average (MA) - Estimation and Forecasting
InputV <- as.ts(df$InputVoltage)
InputV_changes <- diff(InputV)

ts.plot(InputV)
ts.plot(InputV_changes)

acf(InputV_changes, lag.max = 50)

# Fitting a first order MA model
MA_InputV_changes <- arima(InputV, order = c(0, 0, 1))
AR_InputV_changes <- arima(InputV, order = c(1, 0, 0))
print(MA_InputV_changes)
print(AR_InputV_changes)

# Ploting the fitted model
ts.plot(InputV)
MA_InputV_changes_fitted <- InputV - residuals(MA_InputV_changes)
AR_InputV_changes_fitted <- InputV - residuals(AR_InputV_changes)

points(MA_InputV_changes_fitted, type ="l", col = "red", lty = 1)
points(AR_InputV_changes_fitted, type ="l", col = "blue", lty = 1)

# Forecasting after 60 minutes
MA_Voltage_forecast <- predict(MA_InputV_changes, n.ahead = 60)$pred
MA_Voltage_forecast_se <- predict(MA_InputV_changes, n.ahead = 60)$se

AR_Voltage_forecast <- predict(AR_InputV_changes, n.ahead = 60)$pred
AR_Voltage_forecast_se <- predict(AR_InputV_changes, n.ahead = 60)$se

points(MA_Voltage_forecast, type = "l", col = "red")
points(MA_Voltage_forecast - 2 * MA_Voltage_forecast_se, type = "l", col = "red", lty = 2)
points(MA_Voltage_forecast + 2 * MA_Voltage_forecast_se, type = "l", col = "red", lty = 2)

points(AR_Voltage_forecast, type = "l", col = "blue")
points(AR_Voltage_forecast - 2 * AR_Voltage_forecast_se, type = "l", col = "blue", lty = 2)
points(AR_Voltage_forecast + 2 * AR_Voltage_forecast_se, type = "l", col = "blue", lty = 2)

AIC(MA_InputV_changes)
AIC(AR_InputV_changes)

BIC(MA_InputV_changes)
BIC(AR_InputV_changes)
# MA model:
# order = C(0, 0, 1)
# 
# AR model:
# order = c(1, 0, 0)