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


#-----------------------------------
#Compare N rates in soil/plant summary files to those in ARDEC N-rate master file
#
#Read master N-rate file
nRateFile <- 'C:/Users/Robert.Dadamo/Google Drive/USDA/ARDEC LTAR projects/ARDEC N Rates.xlsx'
nRatesAll <- readWithClasses(nRateFile)

#Read ARDEC data file, applying column classes
ardec <- chooseExcelFile()
#Identify current study
currentStudy <- ardec$study[1]
#Replace each NA value with a blank or a zero, depending on column class
# replaceNA(ardec$trtType2)
# replaceNA(ardec$trtRateN2_kg_per_ha)
# replaceNA(ardec$plotSuffix)
ardec[is.na(ardec)] <- 0

library(dplyr)
#Subset nRatesAll for study of interest
nRates <- filter(nRatesAll, study == currentStudy)
#Replace each NA value with a blank or a zero, depending on column class
# replaceNA(nRates$trtType2)
# replaceNA(nRates$trtRateN2_kg_per_ha)
# replaceNA(nRates$plotSuffix)
nRates[is.na(nRates)] <- 0

#Determine years common to both DFs
commonYears <- intersect(unique(nRates$year), unique(ardec$sampYear))

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
      fatalError(cat('trtType1 mismatch in year', yr, 'trtLevel', tr))
    #Create list of current trtType1s
    sfxs <- unique(nRatesYrTr$plotSuffix)
    
    for(sfx in sfxs) {  #For each plotSuffix
      ardecYrTrSfx <- filter(ardecYrTr, plotSuffix == sfx)  #Subset by plotSuffix
      nRatesYrTrSfx <- filter(nRatesYrTr, plotSuffix == sfx)  #Subset by plotSuffix
      if(nrow(ardecYrTrSfx) == 0 | nrow(nRatesYrTrSfx) == 0)
        break
#         fatalError(cat('zero rows in year', yr, 'trtLevel', tr,
#                        'suffix', sfx))
      
      for(rowNum in 1:nrow(ardecYrTrSfx)) {
        
        #*** This if statement throws an error due to the presence of E or W
        #as a plot suffix (due to tillage difference), when no plot suffix appears
        #in the N-rate file (no E/W fertilization difference).
        if((ardecYrTrSfx$trtType1[rowNum] != nRatesYrTrSfx$trtType1[1]) |
           (ardecYrTrSfx$trtType2[rowNum] != nRatesYrTrSfx$trtType2[1]) |
           (ardecYrTrSfx$trtRateN1_kg_per_ha[rowNum] -
            nRatesYrTrSfx$trtRateN1_kg_per_ha[1] > 0.1) |
           (ardecYrTrSfx$trtRateN2_kg_per_ha[rowNum] -
            nRatesYrTrSfx$trtRateN2_kg_per_ha[1]) > 0.1) 
          cat('mismatch in year', yr, 'trtLevel', tr, 'suffix', sfx, 'row',
              rowNum, '\n')
      }
    }
  }
}

 
# ardecYrTrSfx$trtType2[rowNum] == nRatesYrTrSfx$trtType2[1]
# 
# all.equal(ardecYrTrSfx$trtRateN2_kg_per_ha[rowNum], nRatesYrTrSfx$trtRateN2_kg_per_ha[1])
# 
# all.equal(.01, 4, tolerance = .01)
# 
# ardecYrTrSfx$trtRateN2_kg_per_ha[rowNum] == nRatesYrTrSfx$trtRateN2_kg_per_ha[1]
# ardecYrTrSfx$trtRateN2_kg_per_ha[rowNum] <- 3
