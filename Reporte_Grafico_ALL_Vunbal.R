Vunbalance_Report <- function(source_list, initial_dateCR, period_time) {
  
  ## Constantes
  # porcentaje nomimal
  porc_nom <- 3
  
  ## Conexion a SQL Server
  channel <- odbcConnect("SQL_ION", uid="R", pwd="Con3$adm.")
  
  ## Carga de Tabla Source
  sources <- sqlQuery(channel , "select top 100 ID, Name, DisplayName from Source where Name like 'Coopeguanacaste.%'")
  sources$Name <- gsub("Coopeguanacaste.", '', sources$Name)
  
  ## Filtrado de Tabla Source
  sources <- sources %>% filter(ID %in% source_list)
  
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
  initial_date <- with_tz(initial_dateCR, tzone = "UTC")
  final_date <- initial_date + period_time
  
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
  
  ## Sumarizacion de variables
  datalog_byMonth <- dataLog %>% 
    group_by(Meter, Quantity) %>%
    summarise(cantidad = n(), 
              LT_3 = sum(LT_3))
  
  ## Creacion del histograma
  wb <- createWorkbook()
  
  titleStyle <- createStyle(
    fontSize = 18,
    fontColour = "#FFFFFF",
    halign = "center",
    fgFill = "#4F81BD",
    border = "TopBottom",
    borderColour = "#4F81BD"
  )
  
  redStyle <- createStyle(
    fontSize = 14,
    fontColour = "#FFFFFF",
    fgFill = "#FF1111",
    halign = "center",
    textDecoration = "Bold",
  )
  
  mainTableStyle <- createStyle(
    fontSize = 12,
    halign = "center",
    textDecoration = "Bold",
    border = "TopBottomLeftRight",
    borderColour = "#4F81BD"
  )
  
  mainTableStyle2 <- createStyle(
    fontSize = 12,
    halign = "center",
    border = "TopBottomLeftRight",
    borderColour = "#4F81BD"
  )
  
  
  # Recorrer cada medidor
  for(meter in sources$Name){
    print(paste0("Vunbalance ", meter))
    ## meter <- "Casa_Chameleon"
    
    # datos del medidor
    data <- datalog_byMonth %>% filter(Meter == meter)
    data <- as.data.frame(data[,c(2:4)])
    if (nrow(data) > 0) {
      meterFileName = paste0(meter,".png");
      # Crear una hoja 
      addWorksheet(wb, meter)
      
      head_table <- data.frame("c1" = c("Medidor", "Variable", "Cantidad de filas"),
                               "c2" = c(gsub("_", " ", meter), 
                                        "Desbalance de tensión",
                                        sum(data$cantidad)))
      
      head_table2 <- data.frame("c1" = c("Fecha Inicial", "Fecha Final"),
                                "c2" = c(format(initial_date, '%d/%m/%Y', tz = "America/Costa_Rica"),
                                         format(final_date, '%d/%m/%Y', tz = "America/Costa_Rica")))
      
      mergeCells(wb, meter, cols = 2:3, rows = 1)
      mergeCells(wb, meter, cols = 2:3, rows = 2)
      mergeCells(wb, meter, cols = 2:3, rows = 3)
      mergeCells(wb, meter, cols = 6:7, rows = 1)
      mergeCells(wb, meter, cols = 6:7, rows = 2)
      
      writeData(
        wb,
        meter,
        x = head_table,
        startRow = 1,
        rowNames = F,
        colNames = F,
        withFilter = F
      )
      
      writeData(
        wb,
        meter,
        x = head_table2,
        startRow = 1,
        startCol = 5,
        rowNames = F,
        colNames = F,
        withFilter = F
      )
      
      addStyle(wb, sheet = meter, mainTableStyle, rows = 1:3, cols = 1)
      addStyle(wb, sheet = meter, mainTableStyle, rows = 1:2, cols = 5)
      
      addStyle(wb, sheet = meter, mainTableStyle2, rows = 1:3, cols = 2:3, gridExpand = T)
      addStyle(wb, sheet = meter, mainTableStyle2, rows = 1:2, cols = 6:7, gridExpand = T)
      
      if(sum(data$cantidad) < 1008){
        addStyle(wb, sheet = meter, redStyle, rows = 3, cols = 1:3)
      }
      
      
      ## se agregan columnas con porcentuales
      data$LT_3p <- data$LT_3/data$cantidad
      data$GE_3 <- data$cantidad - data$LT_3
      data$GE_3p <- data$GE_3/data$cantidad

      class(data$LT_3p) <- "percentage"
      class(data$GE_3p) <- "percentage"
      colnames(data) <- c("Medida", "Cantidad Total", "Cantidad <3%", "Porc <3%", "Cantidad >=3%", "Porc >3%"  )

      setColWidths(wb, meter, cols = c(1:6), widths = c(20, rep(15, 5)))
      writeDataTable(wb, meter, x = data, startRow = 6, rowNames = F, tableStyle = "TableStyleMedium1")
      
      t1 <- as.data.frame(matrix(data = 0, nrow = 2, ncol = 5))
      colnames(t1) <- c("Quantity", "Type", "Total", "Value", "Perc_Value")
      t1[1,1] <- data[1,1]
      t1[2,1] <- data[1,1]
      t1[1,2] <- "<3%"
      t1[2,2] <- ">=3%"
      t1[1,3] <- data[1,2]
      t1[2,3] <- data[1,2]
      t1[1,4] <- data[1,3]
      t1[2,4] <- data[1,5]
      t1[1,5] <- data[1,4]
      t1[2,5] <- data[1,6]

      p1 <- ggplot(t1, aes(x = Type, y = Perc_Value, fill = Type, label = Type)) +
        geom_col(position = "dodge", width = 0.7) +
        scale_y_continuous(labels = function(x) paste0(100*x, "%"), limits = c(0, 1.2)) +
        geom_text(aes(label=sprintf("%0.2f%%", 100*Perc_Value)), 
                  position = position_dodge(0.9), 
                  angle = 90, size = 2) + 
        theme(axis.text.x = element_text(angle = 90, hjust = 1, size = 8)) +
        ggtitle(gsub("_", " ", meter)) +
        xlab("Distribucion") + 
        ylab("Porcentaje")
      print(p1) #plot needs to be showing
      
      insertPlot(wb, meter, xy = c("A", 9), width = 8, height = 4, fileType = "png", units = "in")
    }
  }
  
  ## nombre de archivo
  fileName <- paste0("C:/Data Science/ArhivosGenerados/Coopeguanacaste/C_Desbalance de tension (",
                     with_tz(initial_date, tzone = "America/Costa_Rica"),
                     ") - (",
                     with_tz(final_date, tzone = "America/Costa_Rica"),
                     ").xlsx")
  saveWorkbook(wb, fileName, overwrite = TRUE)
}