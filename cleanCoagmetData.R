# Coagmet data including GDD

#coagmetRaw <- read.delim('clipboard')
coagPath <- 'W:/ARDEC projects/'
coagFilename <- 'coagmet.xlsx'  # Manually added a year column with data
readFile <- paste(coagPath, coagFilename, sep = '')
library(xlsx)
coagmetRaw <- read.xlsx2(readFile, sheetIndex = 1,
                         colClasses = c(rep('numeric', 15)))

library(dplyr)
coagmetTidy <- filter(coagmetRaw, !is.na(month))
writeFile <- paste(coagPath, 'coagmetTidy.xlsx')
write.xlsx2(coagmetTidy, writeFile)

