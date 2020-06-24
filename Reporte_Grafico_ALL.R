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
source_list <- c(47, 48, 49, 50, 51, 52, 53, 54, 56, 57, 58, 59, 60, 61)
#########################################################################################################

#########################################################################################################
## Para reporte semanal
initial_dateCR <- floor_date(now(), "week", week_start = 1) - weeks(1)
period_time <- weeks(1)

## Para reporte mensual
# initial_dateCR <- floor_date(now(), "month", week_start = 1) - months(1)
# period_time <- months(1)
#########################################################################################################

Vphase_Report(source_list, initial_dateCR, period_time)
Vline_Report(source_list, initial_dateCR, period_time)
Vunbalance_Report(source_list, initial_dateCR, period_time)
Iunbalance_Report(source_list, initial_dateCR, period_time)
VTHD_Report(source_list, initial_dateCR, period_time)
