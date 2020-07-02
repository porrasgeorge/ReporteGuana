
VTHD_Report <- function(initial_dateCR, period_time, sources) {
  
  porc_nom_HD <- 3
  porc_nom_THD <- 5
  
  ## Conexion a SQL Server
  channel <- odbcConnect("SQL_ION", uid = "R", pwd = "Con3$adm.")
  
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
  names(dataLog)[names(dataLog) == "Name"] <- "Quantity"
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
  for (meter in sources$Meter) {
    ## meter <- "Casa_Chameleon"
    print(paste0("THD ", meter))
    
    # datos del medidor
    data_THD_all <- datalog_THD %>% filter(Meter == meter)
    data_HD_all <- datalog_HD %>% filter(Meter == meter)
    
    data_THD <- as.data.frame(data_THD_all[,c(2:7)])
    data_HD <- as.data.frame(data_HD_all[,c(2:7)])
    
    
    if (nrow(data_THD) > 0) {
      meterFileName = paste0(meter, ".png")
      
      # Crear una hoja
      addWorksheet(wb, meter)
      
      head_table <- data.frame("c1" = c("Medidor", "Variable", "# de filas"),
                               "c2" = c(gsub("_", " ", meter), 
                                        "Distorsión Armónica",
                                        sum(data_THD$cantidad)))
      
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
      
      if(sum(data_THD$cantidad) < 3024){
        addStyle(wb, sheet = meter, redStyle, rows = 3, cols = 1:3)
      }
      
      class(data_HD$LT_3p) <- "percentage"
      class(data_HD$GE_3p) <- "percentage" 
      class(data_THD$LT_5p) <- "percentage"
      class(data_THD$GE_5p) <- "percentage" 
      
      colnames(data_HD) <- c("Variable", "Cantidad Total", "Cantidad <3%", "Cantidad >=3%", "Porc <3%", "Porc >=3%")
      colnames(data_THD) <- c("Variable", "Cantidad Total", "Cantidad <5%", "Cantidad >=5%", "Porc <5%", "Porc >=5%")
      setColWidths(wb, meter, cols = c(1:10), widths = c(rep(15, 6)))
      
      writeData(wb, meter, x = c("Distorsión Armónica Total"), startRow = 6)
      mergeCells(wb, meter, cols = 1:6, rows = 6)
      addStyle(wb, sheet = meter, titleStyle, rows = 6, cols = 1)
      ##addStyle(wb, sheet = meter, headerStyle, rows = 6 , cols = 1:6)
      
      writeDataTable(wb, meter, x = data_THD, startRow = 7, rowNames = F, tableStyle = "TableStyleMedium1", withFilter = F)
      
      plot_data <- dataLog %>% filter(Meter == meter, grepl("^V[123]_Harm_Total$", Quantity))
      
      plot_thd <- ggplot(data=plot_data, aes(x=TimestampCR, y=Value, group=Quantity)) +
        ylim(0, NA) +
        ggtitle("Distorsión Armónica Total") +
        geom_step(aes(color=Quantity), direction = 'vh') +
        geom_hline(yintercept=5, 
                   linetype="dashed", 
                   color = "chocolate1", 
                   size=1.5)+ 
        xlab("Fecha") +
        ylab("Porcentaje (%)") 
      
      print(plot_thd) #plot needs to be showing
      
      insertPlot(
        wb,
        meter,
        xy = c("A", 12),
        width = 16,
        height = 4,
        fileType = "png",
        units = "in"
      )
      
      for (i in 2:20) {
        data_row <- data_HD %>% filter(grepl(paste0("^V[123]_Harm", formatC(i, digits = 1, flag = "0"),"$"), Variable))
        class(data_row$"Porc <3%") <- "percentage"
        class(data_row$"Porc >=3%") <- "percentage"
        
        actual_row <- ((i-2)* 26) + 35
        writeData(wb, meter, x = c(paste0(i, ".° Componente de Distorsión Armónica ")), startRow = actual_row)
        mergeCells(wb, meter, cols = 1:6, rows = actual_row)
        addStyle(wb, sheet = meter, titleStyle, rows = actual_row, cols = 1)
        addStyle(wb, sheet = meter, headerStyle, rows = actual_row +1 , cols = 1:6)
        writeDataTable(wb, meter, x = data_row, startRow = actual_row +1 , rowNames = F, tableStyle = "TableStyleMedium1", withFilter = F)
        
        plot_data <- dataLog %>% filter(Meter == meter, grepl(paste0("^V[123]_Harm", formatC(i, digits = 1, flag = "0"),"$"), Quantity))
        
        plot_thd <- ggplot(data=plot_data, aes(x=TimestampCR, y=Value, group=Quantity)) +
          ylim(0, NA) +
          ggtitle(paste0(i, ".° Componente de Distorsión Armónica ")) +
          geom_step(aes(color=Quantity), direction = 'vh') +
          geom_hline(yintercept=3, 
                     linetype="dashed", 
                     color = "chocolate1", 
                     size=1.1) + 
          xlab("Fecha") +
          ylab("Porcentaje (%)") 
        
        print(plot_thd) #plot needs to be showing
        insertPlot(
          wb,
          meter,
          xy = c("A", actual_row + 7),
          width = 16,
          height = 3,
          fileType = "png",
          units = "in"
        )
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


