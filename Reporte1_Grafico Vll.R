rm(list = ls())

## Carga de Librerias
library(dplyr)
library(RODBC)
library(lubridate)
library(ggplot2)
library(tidyr)
library(openxlsx)

## Constantes
# mes anho y numero de meses del reporte
year <- 2020
month <- 5

# tension nomimal
tn <- 24900
limit087 <- 0.87*tn
limit091 <- 0.91*tn
limit093 <- 0.93*tn
limit095 <- 0.95*tn
limit105 <- 1.05*tn
limit107 <- 1.07*tn
limit109 <- 1.09*tn
limit113 <- 1.13*tn
# Bloqueo de las variables
lockBinding(c("tn", "limit087", "limit091", "limit093", "limit095", "limit105", "limit107", "limit109", "limit113"), globalenv())

## Conexion a SQL Server
channel <- odbcConnect("SQL_ION", uid="R", pwd="Con3$adm.")

## Carga de Tabla Source
sources <- sqlQuery(channel , "select top 100 ID, Name, DisplayName from Source where Name like 'Coopeguanacaste.%'")
sources$Name <- gsub("Coopeguanacaste.", '', sources$Name)

## Filtrado de Tabla Source
sources <- sources %>% filter(ID %in% c(47, 48, 49, 50, 51, 52, 53, 54, 56, 57, 58, 59, 60, 61))
##sources <- sources %>% filter(ID %in% c(57))



#########################################################################################################
#############    Valores Promedios    ###################################################################
#########################################################################################################


#Carga de Tabla Quantity
quantity <- sqlQuery(channel , "select top 1500000 ID, Name from Quantity where Name like 'Voltage%'")
odbcCloseAll()

## Filtrado de Tabla Quantity
quantity <- quantity %>% filter(grepl("^Voltage Phases [ABC][ABC] Mean$", Name))
quantity$Name <- c('Vab', 'Vbc', 'Vca')

## Rango de fechas del reporte
# fecha inicial en local time
initial_dateCR <- ymd_hm(paste0(year,"/",month,"/01 00:00"), tz = "America/Costa_Rica")
initial_dateCR <- floor_date(initial_dateCR, "week") + days(1)
end_final_dateCR <- now()

initial_date <- with_tz(initial_dateCR, tzone = "UTC") 
final_date <- initial_date + weeks(1)


while(final_date < end_final_dateCR){
  
  ## Conexion a SQL Server
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
  
  dataLog_spread <- dataLog
  
  dataLog$TN087 <- if_else(dataLog$Value <= limit087, TRUE, FALSE)
  dataLog$TN087_091 <- if_else((dataLog$Value > limit087) & (dataLog$Value <= limit091), TRUE, FALSE)
  dataLog$TN091_093 <- if_else((dataLog$Value > limit091) & (dataLog$Value <= limit093), TRUE, FALSE)
  dataLog$TN093_095 <- if_else((dataLog$Value > limit093) & (dataLog$Value <= limit095), TRUE, FALSE)
  dataLog$TN095_105 <- if_else((dataLog$Value > limit095) & (dataLog$Value <= limit105), TRUE, FALSE)
  dataLog$TN105_107 <- if_else((dataLog$Value > limit105) & (dataLog$Value <= limit107), TRUE, FALSE)
  dataLog$TN107_109 <- if_else((dataLog$Value > limit107) & (dataLog$Value <= limit109), TRUE, FALSE)
  dataLog$TN109_113 <- if_else((dataLog$Value > limit109) & (dataLog$Value <= limit113), TRUE, FALSE)
  dataLog$TN113 <- if_else(dataLog$Value > limit113, TRUE, FALSE)
  
  ## Sumarizacion de variables
  datalog_byMonth <- dataLog %>% 
##    group_by(year, month, Meter, Quantity) %>%
    group_by(Meter, Quantity) %>%
    summarise(cantidad = n(), 
              TN087 = sum(TN087), 
              TN087_091 = sum(TN087_091), 
              TN091_093 = sum(TN091_093), 
              TN093_095 = sum(TN093_095), 
              TN095_105 = sum(TN095_105), 
              TN105_107 = sum(TN105_107), 
              TN107_109 = sum(TN107_109), 
              TN109_113 = sum(TN109_113), 
              TN113 = sum(TN113))
  
  ## Creacion del histograma
  wb <- createWorkbook()
  meter <- sources[1,2]
  # Recorrer cada medidor
  for(meter in sources$Name){
    
    # datos del medidor
    data <- datalog_byMonth %>% filter(Meter == meter)
    data <- as.data.frame(data[,c(2:12)])
    if (nrow(data) > 0) {
      meterFileName = paste0(meter,".png");
      # Crear una hoja 
      addWorksheet(wb, meter)
      
      ## se agregan columnas con porcentuales
      data$TN087p <- data$TN087/data$cantidad
      data$TN087_091p <- data$TN087_091/data$cantidad
      data$TN091_093p <- data$TN091_093/data$cantidad
      data$TN093_095p <- data$TN093_095/data$cantidad
      data$TN095_105p <- data$TN095_105/data$cantidad
      data$TN105_107p <- data$TN105_107/data$cantidad
      data$TN107_109p <- data$TN107_109/data$cantidad
      data$TN109_113p <- data$TN109_113/data$cantidad
      data$TN113p <- data$TN113/data$cantidad
      
      tabla_resumen <- as.data.frame(t(data), stringsAsFactors = F)
      colnames(tabla_resumen) <- as.character(unlist(tabla_resumen[1,]))
      tabla_resumen <- tabla_resumen[3:nrow(tabla_resumen),]
      t1 <- tabla_resumen[1:9,]
      t1[,4:6] <- tabla_resumen[10:18,]
      colnames(t1) <- c("Cantidad_Vab","Cantidad_Vbc","Cantidad_Vca", "Porcent_Vab","Porcent_Vbc","Porcent_Vca")
      t1$Lim_Inferior <- c(0, limit087, limit091, limit093, limit095, limit105, limit107, limit109, limit113)
      t1$Lim_Superior <- c(limit087, limit091, limit093, limit095, limit105, limit107, limit109, limit113, 100000)
      t1 <- t1[,c(7, 8, 1, 4, 2 , 5, 3, 6)]
      
      for (i in 1:ncol(t1)){
        t1[,i] <- as.numeric(t1[,i]) 
      }
      rm(tabla_resumen)
      
      t1 <- tibble::rownames_to_column(t1, "Medicion")
      class(t1$Porcent_Vab) <- "percentage"
      class(t1$Porcent_Vbc) <- "percentage"
      class(t1$Porcent_Vca) <- "percentage"
      
      setColWidths(wb, meter, cols = c(1:10), widths = c(20, rep(15, 9) ))
      writeDataTable(wb, meter, x = t1, startRow = 2, rowNames = F, tableStyle = "TableStyleMedium1")
      
      data_gather <- data[,c(1,2,12:20)]
      data_gather <- data_gather %>% gather("Variable", "Value", -c("Quantity", "cantidad"))
      data_gather$Variable <- factor(data_gather$Variable, levels = c("TN087p", 
                                                                      "TN087_091p", 
                                                                      "TN091_093p", 
                                                                      "TN093_095p", 
                                                                      "TN095_105p", 
                                                                      "TN105_107p", 
                                                                      "TN107_109p", 
                                                                      "TN109_113p", 
                                                                      "TN113p"))
      
      p1 <- ggplot(data_gather, aes(x = Variable, y = Value, fill = Quantity, label = Value)) +
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
  fileName <- paste0("C:/Data Science/ArhivosGenerados/Coopeguanacaste/Vlinea_avg (",
                     with_tz(initial_date, tzone = "America/Costa_Rica"),
                     ") - (",
                     with_tz(final_date, tzone = "America/Costa_Rica"),
                     ").xlsx")
  saveWorkbook(wb, fileName, overwrite = TRUE)
  
  initial_date <- final_date
  final_date <- final_date + weeks(1)
}



