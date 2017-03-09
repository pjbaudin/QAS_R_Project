test_df <- OTList

test_df <- test_df %>%
        select(Date, variable, OTRate, value) %>%
        na.omit() %>%
        group_by(Date, variable, OTRate) %>%
        summarise(OThours = sum(value, na.rm = TRUE)) %>%
        spread(OTRate, OThours, fill = 0) %>%
        group_by(Date) %>%
        summarise(TotalOT = sum(OT),
                  TotalWKD = sum(WKD))


OT <- as.ts(test_df$TotalOT)
WKD <- as.ts(test_df$TotalWKD)

ts.plot(diff(OT))
ts.plot(diff(WKD))

acf(diff(OT), lag.max = 20)
acf(diff(WKD), lag.max = 20)

# Fitting a first order MA model
MA_OT_model <- arima(OT, order = c(0, 0, 1))
AR_OT_model <- arima(OT, order = c(1, 0, 0))


# Ploting the fitted model
ts.plot(OT)
MA_OT_fitted <- OT - residuals(MA_OT_model)
AR_OT_fitted <- OT - residuals(AR_OT_model)

points(MA_OT_fitted, type ="l", col = "red", lty = 1)
points(AR_OT_fitted, type ="l", col = "blue", lty = 1)

# Forecasting for 15 days

MA_OT_forecast <- predict(MA_OT_model, n.ahead = 30)$pred
MA_OT_forecast_se <- predict(MA_OT_model, n.ahead = 30)$se

AR_OT_forecast <- predict(AR_OT_model, n.ahead = 30)$pred
AR_OT_forecast_se <- predict(AR_OT_model, n.ahead = 30)$se

points(MA_OT_forecast, type = "l", col = "red")
points(MA_OT_forecast - 2 * MA_OT_forecast_se, type = "l", col = "red", lty = 2)
points(MA_OT_forecast + 2 * MA_OT_forecast_se, type = "l", col = "red", lty = 2)

points(AR_OT_forecast, type = "l", col = "blue")
points(AR_OT_forecast - 2 * AR_OT_forecast_se, type = "l", col = "blue", lty = 2)
points(AR_OT_forecast + 2 * AR_OT_forecast_se, type = "l", col = "blue", lty = 2)

AIC(MA_OT_model)
AIC(AR_OT_model)

BIC(MA_OT_model)
BIC(AR_OT_model)
# MA model:
# order = C(0, 0, 1)
# 
# AR model:
# order = c(1, 0, 0)