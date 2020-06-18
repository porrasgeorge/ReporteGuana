rm(list = ls())

## Carga de Librerias
library(dplyr)
library(RODBC)
library(lubridate)
library(ggplot2)
library(tidyr)
library(openxlsx)

## Constantes
# porcentaje nomimal
porc_nom <- 0.5

## Conexion a SQL Server
channel <- odbcConnect("SQL_ION", uid="R", pwd="Con3$adm.")

## Carga de Tabla Source
sources <- sqlQuery(channel , "select top 100 ID, Name, DisplayName from Source where Name like 'Coopeguanacaste.%'")
sources$Name <- gsub("Coopeguanacaste.", '', sources$Name)

## Filtrado de Tabla Source
sources <- sources %>% filter(ID %in% c(47, 48, 49, 50, 51, 52, 53, 54, 56, 57, 58, 59, 60, 61))

#########################################################################################################
#############    Valores Promedios    ###################################################################
#########################################################################################################

#Carga de Tabla Quantity
quantity <- sqlQuery(channel , "select top 1500000 ID, Name from Quantity where Name like 'V_unbalance%'")
odbcCloseAll()

## Filtrado de Tabla Quantity
quantity$Name <- c('V_unbal')

## Rango de fechas del reporte
# fecha inicial en local time

initial_dateCR <- floor_date(now(), "week", week_start = 1) - weeks(1)
initial_date <- with_tz(initial_dateCR, tzone = "UTC") 
final_date <- initial_date + weeks(1)

channel <- odbcConnect("SQL_ION", uid="R", pwd="Con3$adm.")
## Carga de Tabla DataLog2
sources_ids <- paste0(sources$ID, collapse = ",")
quantity_ids <- paste0(quantity$ID, collapse = ",")
dataLog2 <- sqlQuery(channel , paste0("select top 500000 * from DataLog2 where ",
                                      "SourceID in (", sources_ids, ")",
                                      " and QuantityID in (", quantity_ids, ")",
                                      " and TimestampUTC >= '", initial_date, "'",
                                      " and TimestampUTC < '", final_date, "'"))

odbcCloseAll()

## Transformacion de columnas
dataLog2$TimestampUTC <- as_datetime(dataLog2$TimestampUTC)
dataLog2$TimestampCR <- with_tz(dataLog2$TimestampUTC, tzone = "America/Costa_Rica") 
dataLog2$TimestampUTC <- NULL
dataLog2$ID <- NULL
dataLog2$year <- year(dataLog2$TimestampCR)
dataLog2$month <- month(dataLog2$TimestampCR)



## Union de tablas, borrado de columnas no importantes y Categorizacion de valores
dataLog <- dataLog2 %>% left_join(quantity, by = c('QuantityID' = "ID")) %>%
  left_join(sources, by = c('SourceID' = "ID"))
#rm(dataLog2)
names(dataLog)[names(dataLog) == "Name.x"] <- "Quantity"
names(dataLog)[names(dataLog) == "Name.y"] <- "Meter"
dataLog$SourceID <- NULL
dataLog$QuantityID <- NULL
dataLog$DisplayName <- NULL

##dataLog_spread <- dataLog

dataLog$LT_3 <- if_else(dataLog$Value < porc_nom, TRUE, FALSE)
dataLog$GT_3 <- if_else(dataLog$Value >= porc_nom, TRUE, FALSE)

## Sumarizacion de variables
datalog_byMonth <- dataLog %>% 
  group_by(Meter, Quantity) %>%
  summarise(cantidad = n(), 
            LT_3 = sum(LT_3), 
            GT_3 = sum(GT_3))

## Creacion del histograma
wb <- createWorkbook()

# Recorrer cada medidor
for(meter in sources$Name){
  
  meter <- "Casa_Chameleon"
  
  # datos del medidor
  data <- datalog_byMonth %>% filter(Meter == meter)
  data <- as.data.frame(data[,c(2:5)])
  if (nrow(data) > 0) {
    meterFileName = paste0(meter,".png");
    # Crear una hoja 
    addWorksheet(wb, meter)
    
    ## se agregan columnas con porcentuales
    data$LT_3p <- data$LT_3/data$cantidad
    data$GT_3p <- data$GT_3/data$cantidad

    
##    setColWidths(wb, meter, cols = c(1:10), widths = c(20, rep(15, 9) ))
    writeDataTable(wb, meter, x = data, startRow = 2, rowNames = F, tableStyle = "TableStyleMedium1")
    
    tabla_resumen <- as.data.frame(t(data), stringsAsFactors = F)
    colnames(tabla_resumen) <- as.character(unlist(tabla_resumen[1,]))
    tabla_resumen <- tabla_resumen[3:nrow(tabla_resumen),]
    tabla_resumen <- as.data.frame(as.numeric(tabla_resumen))
    t1 <- tibble::rownames_to_column(tabla_resumen, "Medicion")
    class(data$LT_3p) <- "percentage"
    class(data$GT_3p) <- "percentage"
    

    p1 <- ggplot(data, aes(x = Quantity, y = Value, fill = Quantity, label = Value)) +
      geom_col(position = "dodge", width = 0.7) +
      scale_y_continuous(labels = function(x) paste0(100*x, "%"), limits = c(0, 1.2)) +
      geom_text(aes(label=sprintf("%0.2f%%", 100*Value)), 
                position = position_dodge(0.9), 
                angle = 90, size = 2) + 
      theme(axis.text.x = element_text(angle = 90, hjust = 1, size = 8)) +
      ggtitle(gsub("_", " ", meter)) +
      xlab("Medicion") + 
      ylab("Porcentaje")
    print(p1) #plot needs to be showing
    insertPlot(wb, meter, xy = c("A", 16), width = 12, height = 6, fileType = "png", units = "in")
  }
}

## nombre de archivo
fileName <- paste0("C:/Data Science/ArhivosGenerados/Coopeguanacaste/Vunbalance_avg (",
                   with_tz(initial_date, tzone = "America/Costa_Rica"),
                   ") - (",
                   with_tz(final_date, tzone = "America/Costa_Rica"),
                   ").xlsx")
saveWorkbook(wb, fileName, overwrite = TRUE)
