library(purrr)      # to use map function
library(dplyr)      # to help with data manipulation
library(tidyr)      # to tidy data
library(stringr)    # to clean data
library(lubridate)  # to tidy the date and time
library(magrittr)   # to use pipe operator
library(tictoc)     # To measure computation time
library(ggvis)
library(shiny)

dat <- res

ind <- res$DocDate > Sys.Date() - 2
last2days <- res[ind,] 

summary(last2days)

last2days %>%
  arrange(desc(DocPrice)) %>%
  filter(DocPrice > 5000) %>%
  filter()
  ggvis(~DocPrice, ~CardCode, stroke := ~DocPrice, fill := "blue") %>%
  layer_points()
  