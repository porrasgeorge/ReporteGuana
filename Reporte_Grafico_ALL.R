rm(list = ls())
options(encoding = "UTF-8")

## Carga de Librerias
library(dplyr)
library(RODBC)
library(lubridate)
library(ggplot2)
library(tidyr)
library(openxlsx)

source(file = "Reporte_Grafico_ALL_Vphase.R")
source(file = "Reporte_Grafico_ALL_Vline.R")
source(file = "Reporte_Grafico_ALL_Vunbal.R")
source(file = "Reporte_Grafico_ALL_Iunbal.R")
source(file = "Reporte_Grafico_ALL_VTHD.R")

#########################################################################################################
## Lista de medidores
source_list <- c(47, 48, 49, 50, 51, 52, 53, 54, 56, 57, 58, 59, 60, 61, 71, 72, 73, 74)
#source_list <- c(78)
#########################################################################################################

channel <- odbcConnect("SQL_ION", uid = "R", pwd = "Con3$adm.")

sources <-
  sqlQuery(
    channel ,
    "select top 1000 ID, Name, DisplayName from Source"
  )
odbcCloseAll()

sources <- sources %>%
  separate(Name, c("Cooperative", "Meter"), "\\.") %>%
  filter(!Cooperative %in% c("LOGINSERTER", "QUERYSERVER", "VIP")) %>% 
  filter(ID %in% source_list)


#########################################################################################################
## Para reporte semanal
#initial_dateCR <- floor_date(now(), "week", week_start = 1) - weeks(1)
initial_dateCR <- floor_date(now(), "week", week_start = 1)
period_time <- weeks(1)

## Para reporte mensual
# initial_dateCR <- floor_date(now(), "month", week_start = 1) - months(1)
# period_time <- months(1)
#########################################################################################################

Vphase_Report(initial_dateCR, period_time, sources)
Vline_Report(initial_dateCR, period_time, sources)
Vunbalance_Report(initial_dateCR, period_time, sources)
Iunbalance_Report(initial_dateCR, period_time, sources)
VTHD_Report(initial_dateCR, period_time, sources)
