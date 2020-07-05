library(RODBC)

####################################################################################################
DB_get_sources <- function(source_list) {

  channel <- odbcConnect("SQL_ION", uid = "R", pwd = "Con3$adm.")
  sources <-
    sqlQuery(channel ,
             "select top 1000 ID, Name, DisplayName from Source")
  odbcCloseAll()
  
  nom_Volt <- DB_get_NominalVoltage()
  
  sources <- sources %>%
    separate(Name, c("Cooperative", "Meter"), "\\.") %>%
    filter(!Cooperative %in% c("LOGINSERTER", "QUERYSERVER", "VIP")) %>%
    filter(ID %in% source_list) %>%
    left_join(nom_Volt, by = c("ID" = "SourceID"))
  
  return(sources)
}


####################################################################################################
DB_get_NominalVoltage <- function(source_list) {
  channel <- odbcConnect("SQL_ION", uid = "R", pwd = "Con3$adm.")
  nominalVoltages <-
    sqlQuery(
      channel ,
      "select t1.SourceID, t1.Value as NomVoltage from [ION_Data].[dbo].[DeviceSettings] t1 
      right join (SELECT [SourceID] ,max([LastVerifiedUTC]) as max_date
                FROM [ION_Data].[dbo].[DeviceSettings]
                where QuantityID = 262
                group by SourceID) as t2
      ON (t2.SourceID = t1.SourceID and t2.max_date = t1.LastVerifiedUTC)
      where QuantityID = 262
    	order by t1.SourceID"
    )
  odbcCloseAll()
  
  nominalVoltages$NomVoltage <- as.numeric(
    case_when(
      ## no disponibles en 34500
      nominalVoltages$NomVoltage == "Not Available" & nominalVoltages$SourceID %in% c(4, 5, 7, 9:16, 25:29, 46, 78:80, 82, 96:99, 112:114) ~ "19920",
      ## no disponibles en 24900
      nominalVoltages$NomVoltage == "Not Available" ~ "14400",
      ## Los demas
      TRUE ~ nominalVoltages$NomVoltage
      ))

  return(nominalVoltages)
}

  
####################################################################################################
DB_get_Vline_quantities <- function() {
  channel <- odbcConnect("SQL_ION", uid = "R", pwd = "Con3$adm.")
  
  quantities <-
    sqlQuery(channel ,
             "select top 1500000 ID, Name from Quantity where Name like 'Vline%'")
  odbcCloseAll()
  
  ## Filtrado de Tabla Quantity
  quantities <-
    quantities %>% filter(grepl("^Vline [abc][abc]$", Name))
  quantities$Name <- c('Vab', 'Vbc', 'Vca')
  return(quantities)
}


####################################################################################################
DB_get_dataLog <- function(sources, quantity, date_range) {
  channel <- odbcConnect("SQL_ION", uid = "R", pwd = "Con3$adm.")
  ## Carga de Tabla DataLog2
  sources_ids <- paste0(sources$ID, collapse = ",")
  quantity_ids <- paste0(quantity$ID, collapse = ",")
  dataLog2 <-
    sqlQuery(
      channel ,
      paste0(
        "select top 500000 * from DataLog2 where ",
        "SourceID in (",
        sources_ids,
        ")",
        " and QuantityID in (",
        quantity_ids,
        ")",
        " and TimestampUTC >= '",
        date_range["initial_date"],
        "'",
        " and TimestampUTC < '",
        date_range["final_date"],
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
    dataLog2 %>% 
    left_join(quantity, by = c('QuantityID' = "ID")) %>%
    left_join(sources, by = c('SourceID' = "ID"))
  names(dataLog)[names(dataLog) == "Name"] <- "Quantity"

  dataLog$SourceID <- NULL
  dataLog$QuantityID <- NULL
  dataLog$DisplayName <- NULL

    return(dataLog)
}


####################################################################################################
add_nomVoltToSources <- function(sources, mult_factor) {
  sources$limit087 <- 0.87 * sources$NomVoltage*mult_factor
  sources$limit091 <- 0.91 * sources$NomVoltage*mult_factor
  sources$limit093 <- 0.93 * sources$NomVoltage*mult_factor
  sources$limit095 <- 0.95 * sources$NomVoltage*mult_factor
  sources$limit105 <- 1.05 * sources$NomVoltage*mult_factor
  sources$limit107 <- 1.07 * sources$NomVoltage*mult_factor
  sources$limit109 <- 1.09 * sources$NomVoltage*mult_factor
  sources$limit113 <- 1.13 * sources$NomVoltage*mult_factor
  
  return(sources)
}


####################################################################################################
classifyGroup_datalogVolts <- function(dataLog) {
  dataLog$TN087 <- if_else(dataLog$Value <= dataLog$limit087, TRUE, FALSE)
  dataLog$TN087_091 <-
    if_else((dataLog$Value > dataLog$limit087) &
              (dataLog$Value <= dataLog$limit091),
            TRUE,
            FALSE)
  dataLog$TN091_093 <-
    if_else((dataLog$Value > dataLog$limit091) &
              (dataLog$Value <= dataLog$limit093),
            TRUE,
            FALSE)
  dataLog$TN093_095 <-
    if_else((dataLog$Value > dataLog$limit093) &
              (dataLog$Value <= dataLog$limit095),
            TRUE,
            FALSE)
  dataLog$TN095_105 <-
    if_else((dataLog$Value > dataLog$limit095) &
              (dataLog$Value <= dataLog$limit105),
            TRUE,
            FALSE)
  dataLog$TN105_107 <-
    if_else((dataLog$Value > dataLog$limit105) &
              (dataLog$Value <= dataLog$limit107),
            TRUE,
            FALSE)
  dataLog$TN107_109 <-
    if_else((dataLog$Value > dataLog$limit107) &
              (dataLog$Value <= dataLog$limit109),
            TRUE,
            FALSE)
  dataLog$TN109_113 <-
    if_else((dataLog$Value > dataLog$limit109) &
              (dataLog$Value <= dataLog$limit113),
            TRUE,
            FALSE)
  dataLog$TN113 <- if_else(dataLog$Value > dataLog$limit113, TRUE, FALSE)
  
  ## Sumarizacion de variables
  datalog_grouped <- dataLog %>%
    group_by(Meter, Quantity) %>%
    summarise(
      cantidad = n(),
      TN087 = sum(TN087),
      TN087_091 = sum(TN087_091),
      TN091_093 = sum(TN091_093),
      TN093_095 = sum(TN093_095),
      TN095_105 = sum(TN095_105),
      TN105_107 = sum(TN105_107),
      TN107_109 = sum(TN107_109),
      TN109_113 = sum(TN109_113),
      TN113 = sum(TN113)
    )
  return(datalog_grouped)
}

