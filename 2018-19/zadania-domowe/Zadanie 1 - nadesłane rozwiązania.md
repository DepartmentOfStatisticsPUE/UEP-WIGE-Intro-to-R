# Proponowane rozwiązania

# Rozwiązanie 1 - Marta U. 

```{r}
library(openxlsx)

Dane2013 <- read.xlsx(xlsxFile = "Zad.1/Data-row/UdSC_2013.xlsx",
                      sheet = "KARTY POBYTU",
                      startRow = 5, 
                      colNames = FALSE)

Dane2014 <- read.xlsx(xlsxFile = "Zad.1/Data-row/UdSC_2014.xlsx",
                      sheet = "DOKUMENTY",
                      startRow = 5, 
                      colNames = FALSE )

Dane2015 <- read.xlsx(xlsxFile = "Zad.1/Data-row/UdSC_2015.xlsx",
                      sheet = "KARTY POBYTU",
                      startRow = 5, 
                      colNames = FALSE )

Dane2016 <- read.xlsx(xlsxFile = "Zad.1/Data-row/UdSC_2016.xlsx",
                      sheet = "KARTY POBYTU",
                      startRow = 5, 
                      colNames = FALSE )

Zadanie1 <- createWorkbook(creator = "Marta U", title= "Statystyki dot. kart pobytu")

addWorksheet(Zadanie1,"karty2013",tabColour = "green")
addWorksheet(Zadanie1,"karty2014",tabColour = "blue")
addWorksheet(Zadanie1,"karty2015",tabColour = "yellow")
addWorksheet(Zadanie1,"karty2016",tabColour = "grey")

style <- createStyle(textDecoration = "bold",textRotation = 45)

writeData(Zadanie1,"karty2013",Dane2013,withFilter = TRUE,headerStyle=style,
          borders = "surrounding",
          borderColour="#0000FF",
          borderStyle="dotted")

writeData(Zadanie1,"karty2014",Dane2014,withFilter = TRUE,
          headerStyle=style,
          borders = "surrounding",
          borderColour="#0000FF",
          borderStyle="dotted")


writeData(Zadanie1,"karty2015",Dane2015,withFilter = TRUE,
          headerStyle=style,
          borders = "surrounding",
          borderColour="#0000FF",
          borderStyle="dotted")

writeData(Zadanie1,"karty2016",Dane2016,withFilter = TRUE,
          headerStyle=style,
          borders = "surrounding",
          borderColour="#0000FF",
          borderStyle="dotted")

freezePane(Zadanie1, sheet = "karty2013", firstRow = TRUE, firstCol = TRUE)
freezePane(Zadanie1, sheet = "karty2014", firstRow = TRUE, firstCol = TRUE)
freezePane(Zadanie1, sheet = "karty2015", firstRow = TRUE, firstCol = TRUE)

freezePane(Zadanie1, sheet = "karty2016", firstRow = TRUE, firstCol = TRUE)

saveWorkbook(Zadanie1,'Zad.1/Data/zadanie.1.xlsx',overwrite = TRUE)
```

# Rozwiązanie 2 - Kornelia R. 

```{r}
library(openxlsx)
## Wczytuję dane z plików nt. Kart Pobytu i przypisuje je do zmiennych

karty_pobytu_2013 <- read.xlsx(xlsxFile = "data-raw/UdSC_2013.xlsx", 
                       sheet = "KARTY POBYTU", 
                       startRow = 5, colNames = FALSE)

karty_pobytu_2014 <- read.xlsx(xlsxFile = "data-raw/UdSC_2014.xlsx", 
                       sheet = "DOKUMENTY", 
                       startRow = 5, colNames = FALSE)

karty_pobytu_2015 <- read.xlsx(xlsxFile = "data-raw/UdSC_2015.xlsx", 
                       sheet = "KARTY POBYTU", 
                       startRow = 5, colNames = FALSE)

karty_pobytu_2016 <- read.xlsx(xlsxFile = "data-raw/UdSC_2016.xlsx", 
                       sheet = "KARTY POBYTU", 
                       startRow = 5, colNames = FALSE)
## Tworzę plik
plik_karty_pobytu <- createWorkbook(creator = "Kornelia R", title = "Statystyki dot. kart pobytu")

## Dodaję puste arkusze do pliku, przypisuję je do zmiennych i nadaję arkuszom kolory

s2013 <- addWorksheet(wb = plik_karty_pobytu, 
                      sheetName = "karty2013", 
                      tabColour = "deepskyblue")

s2014 <- addWorksheet(wb = plik_karty_pobytu, 
                      sheetName = "karty2014", 
                      tabColour = "brown1")

s2015 <- addWorksheet(wb = plik_karty_pobytu, 
                      sheetName = "karty2015", 
                      tabColour = "aquamarine")

s2016 <- addWorksheet(wb = plik_karty_pobytu, 
                      sheetName = "karty2016", 
                      tabColour = "deeppink")

## Definiuję styl nagłówka - pogrubienie i rotacja o 45 stopni

styl_naglowka <- createStyle(textDecoration = "bold", textRotation = 45) 

## Zapisuję dane do odpowiednich arkuszy, dodaję styl nagłówka i obramowania

writeData(wb = plik_karty_pobytu, sheet = s2013, karty_pobytu_2013, 
          headerStyle = styl_naglowka, borders = "all", 
          borderColour = "blue", borderStyle = "dotted")

writeData(wb = plik_karty_pobytu, sheet = s2014, karty_pobytu_2014,
          headerStyle = styl_naglowka, borders = "all", 
          borderColour = "blue", borderStyle = "dotted")

writeData(wb = plik_karty_pobytu, sheet = s2015, karty_pobytu_2015,
          headerStyle = styl_naglowka, borders = "all",
          borderColour = "blue", borderStyle = "dotted")

writeData(wb = plik_karty_pobytu, sheet = s2016, karty_pobytu_2016, 
          headerStyle = styl_naglowka, borders = "all", 
          borderColour = "blue", borderStyle = "dotted") 

## Dodaję filtry do kolumn nagłówka w każdym arkuszu

addFilter(wb = plik_karty_pobytu, sheet = s2013, 
          rows = 1, cols = 1:ncol(karty_pobytu_2013))

addFilter(wb = plik_karty_pobytu, sheet = s2014, 
          rows = 1, cols = 1:ncol(karty_pobytu_2014))

addFilter(wb = plik_karty_pobytu, sheet = s2015, 
          rows = 1, cols = 1:ncol(karty_pobytu_2015))

addFilter(wb = plik_karty_pobytu, sheet = s2016, 
          rows = 1, cols = 1:ncol(karty_pobytu_2016))

## Blokuję pierwszy wiersz i kolumnę każdego arkusza

freezePane(wb = plik_karty_pobytu, sheet = s2013, 
           firstRow = TRUE, firstCol = TRUE)

freezePane(wb = plik_karty_pobytu, sheet = s2014, 
           firstRow = TRUE, firstCol = TRUE)

freezePane(wb = plik_karty_pobytu, sheet = s2015, 
           firstRow = TRUE, firstCol = TRUE)

freezePane(wb = plik_karty_pobytu, sheet = s2016, 
           firstRow = TRUE, firstCol = TRUE)

## Zapisuję plik

saveWorkbook(wb = plik_karty_pobytu,file = "data/zadanie1.xlsx")
```


# Rozwiązanie 3 - Weronika Sz. 

```{r}
## Wczytanie pakietu 

library(openxlsx)

## Wczytanie danych

udsc2013 <- read.xlsx(xlsxFile = "data-row/UdSC_2013.xlsx", 
                      sheet = "KARTY POBYTU", 
                      startRow = 5, 
                      colNames = FALSE)
udsc2014 <- read.xlsx(xlsxFile = "data-row/UdSC_2014.xlsx", 
                      sheet = "DOKUMENTY", 
                      startRow = 5, 
                      colNames = FALSE)
udsc2015 <- read.xlsx(xlsxFile = "data-row/UdSC_2015.xlsx", 
                      sheet = "KARTY POBYTU", 
                      startRow = 5, 
                      colNames = FALSE)
udsc2016 <- read.xlsx(xlsxFile = "data-row/UdSC_2016.xlsx", 
                      sheet = "KARTY POBYTU", 
                      startRow = 5, 
                      colNames = FALSE)

## Utworzenie skoroszytu

workbook <- createWorkbook(creator = "Weronika Sz.", 
                      title = "Statystyki dot. kart pobytu")


## Arkusz dla roku 2013

addWorksheet(wb = workbook, 
             sheetName = "karty2013", 
             tabColour = "#F6CEEC")
writeData(wb = workbook, 
          sheet = "karty2013", 
          x = udsc2013, 
          borders = "surrounding", 
          borderColour = "#2E2EFE", 
          borderStyle = "dotted", 
          withFilter = TRUE, 
          headerStyle = createStyle(textRotation = 45, textDecoration = "bold"))
freezePane(wb = workbook, 
           sheet = "karty2013", 
           firstRow = TRUE, 
           firstCol = TRUE)

## Arkusz dla roku 2014

addWorksheet(wb = workbook, 
             sheetName = "karty2014", 
             tabColour = "#A9E2F3")
writeData(wb = workbook, 
          sheet = "karty2014", 
          x = udsc2014, 
          borders = "surrounding", 
          borderColour = "#2E2EFE", 
          borderStyle = "dotted", 
          withFilter = TRUE, 
          headerStyle = createStyle(textRotation = 45, textDecoration = "bold"))
freezePane(wb = workbook, 
           sheet = "karty2014", 
           firstRow = TRUE, 
           firstCol = TRUE)

## Arkusz dla roku 2015

addWorksheet(wb = workbook, 
             sheetName = "karty2015", 
             tabColour = "#F2F5A9")
writeData(wb = workbook, 
          sheet = "karty2015", 
          x = udsc2015, 
          borders = "surrounding", 
          borderColour = "#2E2EFE", 
          borderStyle = "dotted", 
          withFilter = TRUE, 
          headerStyle = createStyle(textRotation = 45, textDecoration = "bold"))
freezePane(wb = workbook, 
           sheet = "karty2015", 
           firstRow = TRUE, 
           firstCol = TRUE)

## Arkusz dla roku 2016

addWorksheet(wb = workbook, 
             sheetName = "karty2016", 
             tabColour = "#F5DA81")
writeData(wb = workbook, 
          sheet = "karty2016", 
          x = udsc2016, 
          borders = "surrounding", 
          borderColour = "#2E2EFE", 
          borderStyle = "dotted", 
          withFilter = TRUE, 
          headerStyle = createStyle(textRotation = 45, textDecoration = "bold"))
freezePane(wb = workbook, 
           sheet = "karty2016", 
           firstRow = TRUE, 
           firstCol = TRUE)

## Zapisanie do skoroszytu

saveWorkbook(workbook, "zadanie1.xlsx")
```
