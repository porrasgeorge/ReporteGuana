#source("sheet_styles.R")

####################################################################################################
Vline_Report <- function(initial_dateCR, period_time, sources, nom_Volt) {

  ## Rango de fechas del reporte
  initial_date <- with_tz(initial_dateCR, tzone = "UTC")
  date_range <- c("initial_date" = initial_date, 
                  "final_date" = initial_date + period_time)
  rm(initial_date)

  
  # Carga de tablas Quantity y Datalog
  quantity <-DB_get_Vline_quantities()
  sources <- add_nomVoltToSources(sources, sqrt(3))
  dataLog <- DB_get_dataLog(sources, quantity, date_range)
  dataLog_grouped <- classifyGroup_datalogVolts(dataLog)

  # Creacion del libro de excel
  wb <- createWorkbook()
  
  # Recorrer cada medidor
  for (meter in sources$Meter) {
    ## meter <- "Casa_Chameleon"

    print(paste0("Vline ", meter))
    # datos del medidor
    data <- dataLog_grouped %>% filter(Meter == meter)
    data <- as.data.frame(data[, c(2:12)])
    if (nrow(data) > 0) {
      meterFileName = paste0(meter, ".png")
      
      # Crear una hoja
      addWorksheet(wb, meter)
      
      head_table <- data.frame("c1" = c("Medidor", "Variable", "Cantidad de filas"),
                               "c2" = c(gsub("_", " ", meter), 
                                        "Tension de línea",
                                        sum(data$cantidad)))
      
      head_table2 <- data.frame("c1" = c("Fecha Inicial", "Fecha Final"),
                                "c2" = c(format(date_range["initial_date"], '%d/%m/%Y', tz = "America/Costa_Rica"),
                                         format(date_range["final_date"], '%d/%m/%Y', tz = "America/Costa_Rica")))
      
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
      
      if(sum(data$cantidad) < 3024){
        addStyle(wb, sheet = meter, redStyle, rows = 3, cols = 1:3)
      }
      
      ## se agregan columnas con porcentuales
      data$TN087p <- data$TN087 / data$cantidad
      data$TN087_091p <- data$TN087_091 / data$cantidad
      data$TN091_093p <- data$TN091_093 / data$cantidad
      data$TN093_095p <- data$TN093_095 / data$cantidad
      data$TN095_105p <- data$TN095_105 / data$cantidad
      data$TN105_107p <- data$TN105_107 / data$cantidad
      data$TN107_109p <- data$TN107_109 / data$cantidad
      data$TN109_113p <- data$TN109_113 / data$cantidad
      data$TN113p <- data$TN113 / data$cantidad
      
      tabla_resumen <- as.data.frame(t(data), stringsAsFactors = F)
      colnames(tabla_resumen) <-
        as.character(unlist(tabla_resumen[1, ]))
      tabla_resumen <- tabla_resumen[3:nrow(tabla_resumen), ]
      t1 <- tabla_resumen[1:9, ]
      t1[, 4:6] <- tabla_resumen[10:18, ]
      colnames(t1) <-
        c(
          "Cantidad_Vab",
          "Cantidad_Vbc",
          "Cantidad_Vca",
          "Porcent_Vab",
          "Porcent_Vbc",
          "Porcent_Vca"
        )
      rownames(t1)  <-
        c(
          '<87%',
          '87% <=x< 91%',
          '91% <=x< 93%',
          '93% <=x< 95%',
          '95% <=x< 105%',
          '105% <=x< 107%',
          '107% <=x< 109%',
          '109% <=x< 113%',
          '>= 113%'
        )
      t1$Lim_Inferior <-
        c(
          0,
          limit087,
          limit091,
          limit093,
          limit095,
          limit105,
          limit107,
          limit109,
          limit113
        )
      t1$Lim_Superior <-
        c(
          limit087,
          limit091,
          limit093,
          limit095,
          limit105,
          limit107,
          limit109,
          limit113,
          100000
        )
      t1$Lim_Inferior <- round(t1$Lim_Inferior, 0)
      t1$Lim_Superior <- round(t1$Lim_Superior, 0)
      t1 <- t1[, c(7, 8, 1, 4, 2 , 5, 3, 6)]
      
      for (i in 1:ncol(t1)) {
        t1[, i] <- as.numeric(t1[, i])
      }
      rm(tabla_resumen)
      
      t1 <- tibble::rownames_to_column(t1, "Variable")
      class(t1$Porcent_Vab) <- "percentage"
      class(t1$Porcent_Vbc) <- "percentage"
      class(t1$Porcent_Vca) <- "percentage"
      
      setColWidths(wb,
                   meter,
                   cols = c(1:10),
                   widths = c(20, rep(15, 9)))
      writeDataTable(
        wb,
        meter,
        x = t1,
        startRow = 6,
        rowNames = F,
        tableStyle = "TableStyleMedium1",
        withFilter = F
      )
      
      data_gather <- data[, c(1, 2, 12:20)]
      colnames(data_gather) <-
        c('Medicion',
          'Cantidad Total',
          '<87%',
          '87% <=x< 91%',
          '91% <=x< 93%',
          '93% <=x< 95%',
          '95% <=x< 105%',
          '105% <=x< 107%',
          '107% <=x< 109%',
          '109% <=x< 113%',
          '>= 113%'
        )
      data_gather <-
        data_gather %>% gather("Variable", "Value", -c("Medicion", "Cantidad Total"))
      
      data_gather$Variable <-
        factor(
          data_gather$Variable,
          levels = c('<87%',
                     '87% <=x< 91%',
                     '91% <=x< 93%',
                     '93% <=x< 95%',
                     '95% <=x< 105%',
                     '105% <=x< 107%',
                     '107% <=x< 109%',
                     '109% <=x< 113%',
                     '>= 113%'
          )
        )
      
      p1 <-
        ggplot(data_gather,
               aes(
                 x = Variable,
                 y = Value,
                 fill = Medicion,
                 label = Value
               )) +
        geom_col(position = "dodge", width = 0.7) +
        scale_y_continuous(
          labels = function(x)
            paste0(100 * x, "%"),
          limits = c(0, 1.2)
        ) +
        geom_text(
          aes(label = sprintf("%0.2f%%", 100 * Value)),
          position = position_dodge(0.9),
          angle = 90,
          size = 2
        ) +
        theme(axis.text.x = element_text(
          angle = 90,
          hjust = 1,
          size = 8
        )) +
        ggtitle(gsub("_", " ", meter)) +
        xlab("Variable") +
        ylab("Porcentaje")
      print(p1) #plot needs to be showing
      
      insertPlot(
        wb,
        meter,
        xy = c("A", 17),
        width = 12,
        height = 6,
        fileType = "png",
        units = "in"
      )
    }
  }
  
  ## nombre de archivo
  fileName <-
    paste0(
      "C:/Data Science/ArhivosGenerados/Coopeguanacaste/B_Tension de linea (",
      with_tz(date_range["initial_date"], tzone = "America/Costa_Rica"),
      ") - (",
      with_tz(date_range["final_date"], tzone = "America/Costa_Rica"),
      ").xlsx"
    )
  saveWorkbook(wb, fileName, overwrite = TRUE)
  
}
