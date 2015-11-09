################################################################################
#
# This function reads sheet #1 of an Excel file, using the column classes
# listed on sheet #2.
#
################################################################################
readWithClasses <- function(fullFilename) {
  library(xlsx)
  columnClasses <- read.xlsx2(fullFilename, sheetIndex = 2,
                              stringsAsFactors = FALSE)
  classVector <- columnClasses[, 2]
  names(classVector) <- columnClasses[, 1]
  read.xlsx2(fullFilename, sheetIndex = 1, colClasses = classVector,
             stringsAsFactors = FALSE)
}


################################################################################
#
# Function chooseExcelFile
# Spawns a file chooser for selecting appropriate Excel soil file and reads
# file, applying the classes that are specified in the file's second worksheet.
#
################################################################################
chooseExcelFile <- function() {
  
  path <- 'C:/Users/Robert.Dadamo/Google Drive/USDA/ARDEC LTAR projects/'
  fileExt <- '.xlsx'
  defaultFile <- paste(path, 'Soil_or_Plant', fileExt, sep = '')
  filename <- choose.files(default = defaultFile)
  
  # xlsx provides read/write functions for Excel files
  library(xlsx)
  columnClasses <- read.xlsx2(filename, sheetIndex = 2, stringsAsFactors = FALSE)
  classVector <- columnClasses[, 2]
  names(classVector) <- columnClasses[, 1]
  read.xlsx2(filename, sheetIndex = 1, colClasses = classVector,
             stringsAsFactors = FALSE)
}


################################################################################
#
# Function fatalError
# Handle fatal errors with message output
#
################################################################################
fatalError <- function (errorMessage) {
  cat('\n\n')
  fullMessage <- paste(errorMessage, '. Program halted.', sep = '')
  stop(fullMessage, call. = FALSE)
}


#--------------------------------------
#Compare N rates in soil/plant summary files to those in ARDEC N-rate master file
#
#Read master N-rate file
nRateFile <- 'W:/ARDEC projects/ARDEC N Rates.xlsx'
nRatesAll <- readWithClasses(nRateFile)

#Read ARDEC data file, applying column classes
ardec <- chooseExcelFile()
#Identify current study
currentStudy <- ardec$study[1]
#Replace each NA value with a blank or a zero
ardec$trtType2[is.na(ardec$trtType2)] <- ''
ardec$trtRateN2_kg_per_ha[is.na(ardec$trtRateN2_kg_per_ha)] <- 0

  
library(dplyr)
#Subset nRatesAll for study of interest
nRates <- filter(nRatesAll, study == currentStudy)
nRates$trtType2[is.na(nRates$trtType2)] <- ''
nRates$trtRateN2_kg_per_ha[is.na(nRates$trtRateN2_kg_per_ha)] <- 0
#Determine years common to both DFs
commonYears <- intersect(unique(nRates$year), unique(ardec$sampYear))

#Subset both DFs by year, plotSuffix and trtLevel
for(yr in commonYears) {  #For each common year
  #Subset each df for current year
  ardecYr <- filter(ardec, sampYear == yr)  #Subset by year
  nRatesYr <- filter(nRates, year == yr)  #Subset by year
  #Check for mismatch in treatment levels
  if(sort(unique(nRatesYr$trtLevel)) != sort(unique(ardecYr$trtLevel)))
    fatalError(cat('trtLevel mismatch in year', yr))
  #Create list of current trtLevels
  trtLevels <- unique(nRatesYr$trtLevel)
  
  for(tr in trtLevels) {  #For each trtLevel
    ardecYrTr <- filter(ardecYr, trtLevel == tr)  #Subset by trtLevel
    nRatesYrTr <- filter(nRatesYr, trtLevel == tr)  #Subset by trtLevel
    #Check for mismatch in trtType1
    if(sort(unique(nRatesYrTr$trtType1)) != sort(unique(ardecYrTr$trtType1)))
      fatalError(cat('trtType1l mismatch in year', yr, 'trtLevel', tr))
    #Create list of current trtType1s
    tp1s <- unique(nRatesYr$trtType1)
    
    for(tp1 in tp1s) {  #For each trtType1
      ardecYrTrT1 <- filter(ardecYrTr, trtType1 == tp1)  #Subset by trtType1
      nRatesYrTrT1 <- filter(nRatesYrTr, trtType1 == tp1)  #Subset by trtType1
      #Check for mismatch in trtType2
      if(sort(unique(nRatesYrTrT1$trtType2)) != sort(unique(ardecYrTrT1$trtType2)))
        fatalError(cat('trtType2 mismatch in year', yr, 'trtLevel', tr,
                       'trtType1', tp1))
      #Create list of current trtType2s
      tp2s <- unique(nRatesYrTrT1$trtType2)
      
      for(tp2 in tp2s) {  #For each trtType2
        ardecYrTrT1T2 <- filter(ardecYrTrT1, trtType2 == tp2)  #Subset by trtType2
        nRatesYrTrT1T2 <- filter(nRatesYrTrT1, trtType2 == tp2)  #Subset by trtType2
        #Check for mismatch in trtRateN1
        if(sort(unique(nRatesYrTrT1T2$trtRateN1_kg_per_ha)) !=
           sort(unique(ardecYrTrT1T2$trtRateN1_kg_per_ha)))
          fatalError(cat('trtRateN1 mismatch in year', yr, 'trtLevel', tr,
                         'trtType1', tp1, 'trtType2', tp2))
        #Create list of current trtRateN1s
        rt1s <- unique(nRatesYrTrT1T2$trtRateN1_kg_per_ha)
        
        for(rt1 in rt1s) {  #For each trtRateN1
          ardecYrTrT1T2R1 <- filter(ardecYrTrT1T2, trtRateN1_kg_per_ha == rt1)  #Subset by trtRate1
          nRatesYrTrT1T2R1 <- filter(nRatesYrTrT1T2, trtRateN1_kg_per_ha == rt1)  #Subset by trtRate1
          #Check for mismatch in trtRateN2
          if(sort(unique(nRatesYrTrT1T2R1$trtRateN2_kg_per_ha)) !=
             sort(unique(ardecYrTrT1T2R1$trtRateN2_kg_per_ha)))
            fatalError(cat('trtRateN2 mismatch in year', yr, 'trtLevel', tr,
                           'trtType1', tp1, 'trtType2', tp2, 'trtRate1', rt1))
        }
      }
    }
  }
}


