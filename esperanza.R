library(dplyr)
library(lubridate)



a <- read.csv("Esperanza.CSV", stringsAsFactors = F)
# n <- colnames(a)
# n[1] = "Time"
# colnames(a) <- n

glimpse(a)

a$Time <- dmy_hm(a$Time)
# a$Esp.U1 <- as.numeric(a$Esp.U1)
# a$Esp.U2 <- as.numeric(a$Esp.U2)

a <- a %>% mutate(anho = year(Time), mes = month(Time), dia = day(Time), hora = hour(Time), minuto = minute(Time), minute2 = 15 * as.integer(minute(Time)/15))


b <- a %>% group_by(anho, mes, dia, hora, minute2) %>% 
summarise(U1 = mean(Esp.U1), U2 = mean(Esp.U2))
  
c <- b
c$U1 <- lag(c$U1)
c$U2 <- lag(c$U2)

write.csv(x = c, file = "Esperanza_resumen.csv")
