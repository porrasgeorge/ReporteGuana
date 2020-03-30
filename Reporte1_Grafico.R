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
month <- 3

# tension nomimal
tn <- 14376.02
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
channel <- odbcConnect("SQL_ION", uid="sa", pwd="Con3$adm.")

## Carga de Tabla Source
sources <- sqlQuery(channel , "select top 100 ID, Name, DisplayName from Source where Name like 'Coopeguanacaste.%'")
sources$Name <- gsub("Coopeguanacaste.", '', sources$Name)

## Filtrado de Tabla Source
sources <- sources %>% filter(ID %in% c(47, 48, 49, 50, 51, 52, 53, 54))
##sources <- sources %>% filter(ID %in% c(47, 48, 49))



#########################################################################################################
#############    Valores Promedios    ###################################################################
#########################################################################################################


#Carga de Tabla Quantity
quantity <- sqlQuery(channel , "select top 1500000 ID, Name from Quantity where Name like 'Voltage%'")

## Filtrado de Tabla Quantity
quantity <- quantity %>% filter(grepl("^Voltage on Input V[123] Mean - Power Quality Monitoring$", Name))
quantity$Name <- c('Van', 'Vbn', 'Vcn')

## Rango de fechas del reporte
# fecha inicial en local time
initial_dateCR <- ymd_hm(paste0(year,"/",month,"/01 00:00"), tz = "America/Costa_Rica")
# conversion a utc y creacion de fecha final
initial_date <- with_tz(initial_dateCR, tzone = "UTC") 
# fecha final en utc
final_date <- initial_date + months(1)


## Carga de Tabla DataLog2
sources_ids <- paste0(sources$ID, collapse = ",")
quantity_ids <- paste0(quantity$ID, collapse = ",")
dataLog2 <- sqlQuery(channel , paste0("select top 500000 * from DataLog2 where ",
                                      "SourceID in (", sources_ids, ")",
                                      " and QuantityID in (", quantity_ids, ")",
                                      " and TimestampUTC >= '", initial_date, "'",
                                      " and TimestampUTC < '", final_date, "'"))
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
  group_by(year, month, Meter, Quantity) %>%
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

# Recorrer cada medidor
for(meter in sources$Name){

  # datos del medidor
  data <- datalog_byMonth %>% filter(Meter == meter)
  data <- as.data.frame(data[,c(4:14)])
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
    colnames(t1) <- c("Cantidad_Van","Cantidad_Vbn","Cantidad_Vcn", "Porcent_Van","Porcent_Vbn","Porcent_Vcn")
    t1$Lim_Inferior <- c(0, limit087, limit091, limit093, limit095, limit105, limit107, limit109, limit113)
    t1$Lim_Superior <- c(limit087, limit091, limit093, limit095, limit105, limit107, limit109, limit113, 100000)
    t1 <- t1[,c(7, 8, 1, 4, 2 , 5, 3, 6)]
    
    for (i in 1:ncol(t1)){
      t1[,i] <- as.numeric(t1[,i]) 
    }
    rm(tabla_resumen)
    
    t1 <- tibble::rownames_to_column(t1, "Medicion")
    class(t1$Porcent_Van) <- "percentage"
    class(t1$Porcent_Vbn) <- "percentage"
    class(t1$Porcent_Vcn) <- "percentage"
    
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
fileName <- paste0("C:/Data Science/ArhivosGenerados/Coopeguanacaste/Coopeguanacaste_Reporte1_Mean_",
                   initial_dateCR,
                   ".xlsx")
saveWorkbook(wb, fileName, overwrite = TRUE)



#########################################################################################################
#############    Desbalances de Tension      ############################################################
#########################################################################################################

asimetrias  <- dataLog_spread %>% 
  spread(key = Quantity, value = Value)%>%
  rowwise() %>%
  mutate(AVG = mean(c(Van, Vbn, Vcn))) %>%
  mutate(Vperc = max(abs(Van - AVG),abs(Vbn - AVG),abs(Vcn - AVG)) / AVG)

asimetrias <- asimetrias[,-c(5, 6, 7)]
asimetrias$P05 <- if_else(asimetrias$Vperc <= 0.005, TRUE, FALSE)
asimetrias$P05_10 <- if_else((asimetrias$Vperc > 0.005) & (asimetrias$Vperc <= 0.010), TRUE, FALSE)
asimetrias$P10_15 <- if_else((asimetrias$Vperc > 0.010) & (asimetrias$Vperc <= 0.015), TRUE, FALSE)
asimetrias$P15_20 <- if_else((asimetrias$Vperc > 0.015) & (asimetrias$Vperc <= 0.020), TRUE, FALSE)
asimetrias$P20_25 <- if_else((asimetrias$Vperc > 0.020) & (asimetrias$Vperc <= 0.025), TRUE, FALSE)
asimetrias$P25_30 <- if_else((asimetrias$Vperc > 0.025) & (asimetrias$Vperc <= 0.030), TRUE, FALSE)
asimetrias$P30 <- if_else((asimetrias$Vperc > 0.030), TRUE, FALSE)
asimetrias <- asimetrias %>% ungroup()

asimetrias_byMonth <- asimetrias %>% 
group_by(year, month, Meter) %>%
  summarise(cantidad = n(), 
            P05 = sum(P05), 
            P05_10 = sum(P05_10), 
            P10_15 = sum(P10_15), 
            P15_20 = sum(P15_20), 
            P20_25 = sum(P20_25), 
            P25_30 = sum(P25_30), 
            P30 = sum(P30))


setColWidths(wb2, ws2, cols = c(3:13), widths = c(20, rep(12,10)))
writeDataTable(wb2, ws2, x = datalog_byMonth_perc, startRow = 2, rowNames = F, tableStyle = "TableStyleMedium1")
# addDataFrame(as.data.frame(datalog_byMonth_perc), sheet2, col.names=TRUE, row.names=F,
#               startRow=1, startColumn=1, rownamesStyle=TITLE_STYLE)

## nombre de archivo
fileName <- paste0("C:/Data Science/ArhivosGenerados/Coopeguanacaste/Coopeguanacaste_Reporte2_",
                   initial_dateCR,
                   ".xlsx")
saveWorkbook(wb2, fileName, overwrite = TRUE)

#rm(list = ls())




#########################################################################################################
#############    Valores Maximos      ###################################################################
#########################################################################################################


#Carga de Tabla Quantity
quantity <- sqlQuery(channel , "select top 1500000 ID, Name from Quantity where Name like 'Voltage%'")

## Filtrado de Tabla Quantity
quantity <- quantity %>% filter(grepl("^Voltage on Input V[123] High - Power Quality Monitoring$", Name))
quantity$Name <- c('Van', 'Vbn', 'Vcn')

## Rango de fechas del reporte
# fecha inicial en local time
initial_dateCR <- ymd_hm(paste0(year,"/",month,"/01 00:00"), tz = "America/Costa_Rica")
# conversion a utc y creacion de fecha final
initial_date <- with_tz(initial_dateCR, tzone = "UTC") 
# fecha final en utc
final_date <- initial_date + months(1)


## Carga de Tabla DataLog2
sources_ids <- paste0(sources$ID, collapse = ",")
quantity_ids <- paste0(quantity$ID, collapse = ",")
dataLog2 <- sqlQuery(channel , paste0("select top 500000 * from DataLog2 where ",
                                      "SourceID in (", sources_ids, ")",
                                      " and QuantityID in (", quantity_ids, ")",
                                      " and TimestampUTC >= '", initial_date, "'",
                                      " and TimestampUTC < '", final_date, "'"))
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
  group_by(year, month, Meter, Quantity) %>%
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

# Recorrer cada medidor
for(meter in sources$Name){
  
  # datos del medidor
  data <- datalog_byMonth %>% filter(Meter == meter)
  data <- as.data.frame(data[,c(4:14)])
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
    colnames(t1) <- c("Cantidad_Van","Cantidad_Vbn","Cantidad_Vcn", "Porcent_Van","Porcent_Vbn","Porcent_Vcn")
    t1$Lim_Inferior <- c(0, limit087, limit091, limit093, limit095, limit105, limit107, limit109, limit113)
    t1$Lim_Superior <- c(limit087, limit091, limit093, limit095, limit105, limit107, limit109, limit113, 100000)
    t1 <- t1[,c(7, 8, 1, 4, 2 , 5, 3, 6)]
    
    for (i in 1:ncol(t1)){
      t1[,i] <- as.numeric(t1[,i]) 
    }
    rm(tabla_resumen)
    
    t1 <- tibble::rownames_to_column(t1, "Medicion")
    class(t1$Porcent_Van) <- "percentage"
    class(t1$Porcent_Vbn) <- "percentage"
    class(t1$Porcent_Vcn) <- "percentage"
    
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
fileName <- paste0("C:/Data Science/ArhivosGenerados/Coopeguanacaste/Coopeguanacaste_Reporte1_Max_",
                   initial_dateCR,
                   ".xlsx")
saveWorkbook(wb, fileName, overwrite = TRUE)


#########################################################################################################
#############    Valores Minimos      ###################################################################
#########################################################################################################


#Carga de Tabla Quantity
quantity <- sqlQuery(channel , "select top 1500000 ID, Name from Quantity where Name like 'Voltage%'")

## Filtrado de Tabla Quantity
quantity <- quantity %>% filter(grepl("^Voltage on Input V[123] Low - Power Quality Monitoring$", Name))
quantity$Name <- c('Van', 'Vbn', 'Vcn')

## Rango de fechas del reporte
# fecha inicial en local time
initial_dateCR <- ymd_hm(paste0(year,"/",month,"/01 00:00"), tz = "America/Costa_Rica")
# conversion a utc y creacion de fecha final
initial_date <- with_tz(initial_dateCR, tzone = "UTC") 
# fecha final en utc
final_date <- initial_date + months(1)


## Carga de Tabla DataLog2
sources_ids <- paste0(sources$ID, collapse = ",")
quantity_ids <- paste0(quantity$ID, collapse = ",")
dataLog2 <- sqlQuery(channel , paste0("select top 500000 * from DataLog2 where ",
                                      "SourceID in (", sources_ids, ")",
                                      " and QuantityID in (", quantity_ids, ")",
                                      " and TimestampUTC >= '", initial_date, "'",
                                      " and TimestampUTC < '", final_date, "'"))
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
  group_by(year, month, Meter, Quantity) %>%
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

# Recorrer cada medidor
for(meter in sources$Name){
  
  # datos del medidor
  data <- datalog_byMonth %>% filter(Meter == meter)
  data <- as.data.frame(data[,c(4:14)])
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
    colnames(t1) <- c("Cantidad_Van","Cantidad_Vbn","Cantidad_Vcn", "Porcent_Van","Porcent_Vbn","Porcent_Vcn")
    t1$Lim_Inferior <- c(0, limit087, limit091, limit093, limit095, limit105, limit107, limit109, limit113)
    t1$Lim_Superior <- c(limit087, limit091, limit093, limit095, limit105, limit107, limit109, limit113, 100000)
    t1 <- t1[,c(7, 8, 1, 4, 2 , 5, 3, 6)]
    
    for (i in 1:ncol(t1)){
      t1[,i] <- as.numeric(t1[,i]) 
    }
    rm(tabla_resumen)
    
    t1 <- tibble::rownames_to_column(t1, "Medicion")
    class(t1$Porcent_Van) <- "percentage"
    class(t1$Porcent_Vbn) <- "percentage"
    class(t1$Porcent_Vcn) <- "percentage"
    
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
fileName <- paste0("C:/Data Science/ArhivosGenerados/Coopeguanacaste/Coopeguanacaste_Reporte1_Min_",
                   initial_dateCR,
                   ".xlsx")
saveWorkbook(wb, fileName, overwrite = TRUE)

odbcCloseAll()

##rm(list = ls())


