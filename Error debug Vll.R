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

end_dateCR <- floor_date(now(), "week")
start_dateCR <- end_dateCR - weeks(5)+ days(1)

initial_date <- with_tz(start_dateCR, tzone = "UTC") 
final_date <- with_tz(end_dateCR, tzone = "UTC") 

# tension nomimal
tn <- 24900
tensiones <- c(tn, 0.87*tn, 0.91*tn,0.93*tn,0.95*tn,1.05*tn,1.07*tn,1.09*tn,1.13*tn)
names(tensiones) <- c("Nom", "limit087" ,"limit091","limit093","limit095" ,"limit105" ,"limit107" ,"limit109" ,"limit113")


## Conexion a SQL Server
channel <- odbcConnect("SQL_ION", uid="sa", pwd="Con3$adm.")

## Carga de Tabla Source
sources <- sqlQuery(channel , "select top 100 ID, Name, DisplayName from Source where Name like 'Coopeguanacaste.%'")
sources$Name <- gsub("Coopeguanacaste.", '', sources$Name)
sources <- sources %>% filter(ID %in% c(57))

#Carga de Tabla Quantity
quantity <- sqlQuery(channel , "select top 1500000 ID, Name from Quantity where Name like 'Voltage%'")
quantity <- quantity %>% filter(grepl("^Voltage Phases [ABC][ABC] Mean$", Name))
quantity$Name <- c('Vab', 'Vbc', 'Vca')

## Carga de Tabla dataLog
sources_ids <- paste0(sources$ID, collapse = ",")
quantity_ids <- paste0(quantity$ID, collapse = ",")

dataLog <- sqlQuery(channel , paste0("select top 500000 * from dataLog2 where ",
                                      "SourceID in (", sources_ids, ")",
                                      " and QuantityID in (", quantity_ids, ")",
                                      " and TimestampUTC >= '", initial_date, "'",
                                      " and TimestampUTC < '", final_date, "'"))

odbcCloseAll()

## Transformacion de columnas
dataLog$TimestampUTC <- as_datetime(dataLog$TimestampUTC)
dataLog$TimestampCR <- with_tz(dataLog$TimestampUTC, tzone = "America/Costa_Rica") 
dataLog$TimestampUTC <- NULL
dataLog$ID <- NULL
#dataLog$year <- year(dataLog$TimestampCR)
#dataLog$month <- month(dataLog$TimestampCR)



## Union de tablas, borrado de columnas no importantes y Categorizacion de valores
dataLog <- dataLog %>% left_join(quantity, by = c('QuantityID' = "ID")) %>%
  left_join(sources, by = c('SourceID' = "ID"))
#rm(dataLog)
names(dataLog)[names(dataLog) == "Name.x"] <- "Quantity"
names(dataLog)[names(dataLog) == "Name.y"] <- "Meter"
dataLog$SourceID <- NULL
dataLog$QuantityID <- NULL
dataLog$DisplayName <- NULL

print(tensiones)


dataLog <- dataLog %>% mutate(classif = case_when(Value < tensiones["limit087"] ~ "TN087",
                                                  Value < tensiones["limit091"] ~ "TN087_091",
                                                  Value < tensiones["limit093"] ~ "TN091_093",
                                                  Value < tensiones["limit095"] ~ "TN093_095",
                                                  Value < tensiones["limit105"] ~ "TN095_105",
                                                  Value < tensiones["limit107"] ~ "TN105_107",
                                                  Value < tensiones["limit109"] ~ "TN107_109",
                                                  Value < tensiones["limit113"] ~ "TN109_113",
                                                  TRUE ~ "TN113"
                                                  ))
  

dataLog$classif <- factor(dataLog$classif, levels = list("TN087", "TN087_091", "TN091_093", "TN093_095", "TN095_105", "TN105_107", "TN107_109", "TN109_113", "TN113"))
dataLog_table <- as.data.frame(table(dataLog$classif, dataLog$Quantity, dnn = c("Classif", "Quantity")))
sum_total <- sum(dataLog_table$Freq)/3
dataLog_table$Perc <- round(100*(dataLog_table$Freq/sum_total), 2)





t1 <- dataLog_table %>% select(Classif, Quantity, Freq) %>% spread(Quantity, value = Freq, fill = 0)
t2 <- dataLog_table %>% select(Classif, Quantity, Perc) %>% spread(Quantity, value = Perc, fill = 0)
t3 <- t1 %>% left_join(t2, by = "Classif")
rm(t1, t2)
t3$Lim_Inferior <- c(0, limit087, limit091, limit093, limit095, limit105, limit107, limit109, limit113)
t3$Lim_Superior <- c(limit087, limit091, limit093, limit095, limit105, limit107, limit109, limit113, 100000)
t3 <- t3[,c(1, 8, 9, 2, 5, 3 , 6, 4, 7)]
colnames(t3) <- c("Clasificacion", "Lim_Inferior", "Lim_Superior", "Cantidad_Vab", "Porcent_Vab", "Cantidad_Vbc", "Porcent_Vbc", "Cantidad_Vca", "Porcent_Vca")  
  






glimpse(dataLog_table)





a <- c(1, 2 ,3 5, 6)
b <- paste0(a, "%")}
sprintf('%', a)






p1 <- ggplot(dataLog_table, aes(x = Classif, y = Perc, label = Perc, fill = Quantity )) +
  geom_col(position = "dodge", width = 0.7) +
  scale_y_continuous(labels = function(x) paste0(100*x, "%"), limits = c(0, 1.2)) +
  geom_text(aes(label=sprintf("%0.2f%%", 100*Perc)), 
            position = position_dodge(0.9), 
            angle = 90, size = 2) + 
  theme(axis.text.x = element_text(angle = 90, hjust = 1, size = 8)) +
  ggtitle("meter") +
  xlab("Medicion") + 
  ylab("Porcentaje")

p1



wb <- createWorkbook()
addWorksheet(wb, "Sheet 1")

c1 <- createComment(comment = "this is comment")
writeComment(wb, 1, col = "B", row = 10, comment = c1)

s1 <- createStyle(fontSize = 12, fontColour = "red", textDecoration = c("BOLD"))
s2 <- createStyle(fontSize = 9, fontColour = "black")

c2 <- createComment(comment = c("This Part Bold red\n\n", "This part black"), style = c(s1, s2))
c2

writeComment(wb, 1, col = 6 , row = 3, comment = c2)
writeData(wb, 1, x = c("Vanidad"), startRow = 5)
mergeCells(wb, 1, cols = 1:6, rows = 1)

saveWorkbook(wb, file = "writeCommentExample.xlsx", overwrite = TRUE)


for (i in 2:20) {
print(paste0("v", formatC(i, digits = 1, flag = "0")))
}

format(1:10, digits = 5)

x <- data.frame("SN" = 1:2, "Age" = c(21,15), "Name" = c("John","Dora"))

head_table <- data.frame("c1" = c("Medidor", "Variable","Fecha Inicial", "Fecha Final"),
                         "c2" = c("m1", "V1","f2", "f2"))


format(initial_date, '%d/%m/%Y', tz = "America/Costa_Rica")



