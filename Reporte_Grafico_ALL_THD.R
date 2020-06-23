rm(list = ls())

## Carga de Librerias
library(dplyr)
library(RODBC)
library(lubridate)
library(ggplot2)
library(tidyr)
library(openxlsx)

VTHD_Report <- function(source_list, initial_dateCR, period_time) {
  
  porc_nom_HD <- 3
  porc_nom_THD <- 5
  
  ## Conexion a SQL Server
  channel <- odbcConnect("SQL_ION", uid = "R", pwd = "Con3$adm.")
  
  ## Carga de Tabla Source
  sources <-
    sqlQuery(
      channel ,
      "select top 100 ID, Name, DisplayName from Source where Name like 'Coopeguanacaste.%'"
    )
  sources$Name <- gsub("Coopeguanacaste.", '', sources$Name)
  
  ## Filtrado de Tabla Source
  sources <- sources %>% filter(ID %in% source_list)
  
  #########################################################################################################
  #############    Valores Promedios    ###################################################################
  #########################################################################################################
  
  
  ## NEW DATA
  #Carga de Tabla Quantity
  quantity <-
    sqlQuery(channel ,
             "select top 1500000 ID, Name from Quantity where Name like '%_Harm%'order by Name")
  odbcCloseAll()
  
  ## Filtrado de Tabla Quantity
  quantity <- quantity %>% 
    filter(grepl("^V[123]_Harm([0-9][0-9]|_Total)$", Name),
           !grepl("^V[123]_Harm01$", Name)) %>% 
    arrange(Name)
  ##  quantity$Name <- c('Van', 'Vbn', 'Vcn')
  
  ## Rango de fechas del reporte
  ## fecha inicial en local time
  initial_date <- with_tz(initial_dateCR, tzone = "UTC")
  final_date <- initial_date + period_time
  
  
  channel <- odbcConnect("SQL_ION", uid = "R", pwd = "Con3$adm.")
  ## Carga de Tabla DataLog2
  sources_ids <- paste0(sources$ID, collapse = ",")
  quantity_ids <- paste0(quantity$ID, collapse = ",")
  dataLog2 <-
    sqlQuery(
      channel ,
      paste0(
        "select top 5000000 * from DataLog2 where ",
        "SourceID in (",
        sources_ids,
        ")",
        " and QuantityID in (",
        quantity_ids,
        ")",
        " and TimestampUTC >= '",
        initial_date,
        "'",
        " and TimestampUTC < '",
        final_date,
        "'"
      )
    )
  
  odbcCloseAll()
  
  ## Transformacion de columnas
  dataLog2$TimestampUTC <- as_datetime(dataLog2$TimestampUTC)
  dataLog2$TimestampCR <-
    with_tz(dataLog2$TimestampUTC, tzone = "America/Costa_Rica")
  dataLog2$TimestampUTC <- NULL
  dataLog2$ID <- NULL
  
  ## Union de tablas, borrado de columnas no importantes y Categorizacion de valores
  dataLog <-
    dataLog2 %>% left_join(quantity, by = c('QuantityID' = "ID")) %>%
    left_join(sources, by = c('SourceID' = "ID"))
  #rm(dataLog2)
  names(dataLog)[names(dataLog) == "Name.x"] <- "Quantity"
  names(dataLog)[names(dataLog) == "Name.y"] <- "Meter"
  dataLog$SourceID <- NULL
  dataLog$QuantityID <- NULL
  dataLog$DisplayName <- NULL
  
  datalog_THD <- dataLog %>% 
    filter(grepl("^V[123]_Harm_Total$", Quantity)) %>%
    arrange(Meter, Quantity, TimestampCR) %>%
    mutate(LT_5percent = ifelse(Value < porc_nom_THD, TRUE, FALSE)) %>%
    group_by(Meter, Quantity) %>%
    summarise(cantidad = n(),
              LT_5 = sum(LT_5percent),
              GE_5 = sum(!LT_5percent)) %>%
    mutate(LT_5p = LT_5/cantidad, 
           GE_5p = GE_5/cantidad)

  
  datalog_HD <- dataLog %>% 
    filter(grepl("^V[123]_Harm[0-9][0-9]$", Quantity),
           !grepl("^V[123]_Harm01$", Quantity)) %>%
    arrange(Meter, Quantity, TimestampCR) %>%
    mutate(LT_3percent = ifelse(Value < porc_nom_HD, TRUE, FALSE)) %>%
    group_by(Meter, Quantity) %>%
    summarise(cantidad = n(),
              LT_3 = sum(LT_3percent),
              GE_3 = sum(!LT_3percent)) %>%
    mutate(LT_3p = LT_3/cantidad, 
           GE_3p = GE_3/cantidad)
  
  
  
  ## Creacion del histograma
  wb <- createWorkbook()
  
  titleStyle <- createStyle(
    fontSize = 18, fontColour = "#FFFFFF", halign = "center",
    fgFill = "#4F81BD", border = "TopBottom", borderColour = "#4F81BD")
  
  headerStyle <- createStyle(
      fontSize = 12, halign = "center", textDecoration = "Bold")
  
  
  # Recorrer cada medidor
  for (meter in sources$Name) {
    ## meter <- "Casa_Chameleon"
    print(paste0("THD ", meter))
    
    # datos del medidor
    data_THD <- datalog_THD %>% filter(Meter == meter)
    data_HD <- datalog_HD %>% filter(Meter == meter)
    
    data_THD <- as.data.frame(data_THD[,c(2:7)])
    data_HD <- as.data.frame(data_HD[,c(2:7)])
    
    
    if (nrow(data_THD) > 0) {
      meterFileName = paste0(meter, ".png")
      
      # Crear una hoja
      addWorksheet(wb, meter)

      class(data_HD$LT_3p) <- "percentage"
      class(data_HD$GE_3p) <- "percentage" 
      class(data_THD$LT_5p) <- "percentage"
      class(data_THD$GE_5p) <- "percentage" 
      
      colnames(data_HD) <- c("Variable", "Cantidad Total", "Cantidad <3%", "Cantidad >=3%", "Porc <3%", "Porc >=3%")
      colnames(data_THD) <- c("Variable", "Cantidad Total", "Cantidad <5%", "Cantidad >=5%", "Porc <5%", "Porc >=5%")
      setColWidths(wb, meter, cols = c(1:10), widths = c(rep(15, 6)))
      
      writeData(wb, meter, x = c("Distorsión Armónica Total"), startRow = 1)
      mergeCells(wb, meter, cols = 1:6, rows = 1)
      addStyle(wb, sheet = meter, titleStyle, rows = 1, cols = 1)
      addStyle(wb, sheet = meter, headerStyle, rows = 2 , cols = 1:6)
      writeDataTable(wb, meter, x = data_THD, startRow = 2, rowNames = F, tableStyle = "TableStyleMedium1", withFilter = F)
      
      for (i in 2:20) {
        data_row <- data_HD %>% filter(grepl(paste0("^V[123]_Harm", formatC(i, digits = 1, flag = "0"),"$"), Variable))
        class(data_row$"Porc <3%") <- "percentage"
        class(data_row$"Porc >=3%") <- "percentage"
        
        actual_row <- ((i-2)* 7) + 9
        writeData(wb, meter, x = c(paste0(i, "° Componente de Distorsión Armónica ")), startRow = actual_row)
        mergeCells(wb, meter, cols = 1:6, rows = actual_row)
        addStyle(wb, sheet = meter, titleStyle, rows = actual_row, cols = 1)
        addStyle(wb, sheet = meter, headerStyle, rows = actual_row +1 , cols = 1:6)
        writeDataTable(wb, meter, x = data_row, startRow = actual_row +1 , rowNames = F, tableStyle = "TableStyleMedium1", withFilter = F)
      }
    }
  }
  
  ## nombre de archivo
  fileName <-
    paste0(
      "C:/Data Science/ArhivosGenerados/Coopeguanacaste/E_THD (",
      with_tz(initial_date, tzone = "America/Costa_Rica"),
      ") - (",
      with_tz(final_date, tzone = "America/Costa_Rica"),
      ").xlsx"
    )
  saveWorkbook(wb, fileName, overwrite = TRUE)
  
}


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

VTHD_Report(source_list, initial_dateCR, period_time)


