################################################################################
#
# Function chooseExcelFile
# Spawn a file chooser for selecting appropriate Excel soil file
#
################################################################################

chooseExcelFile <- function() {
  
  path <- 'C:/Users/Robert.Dadamo/Google Drive/USDA/ARDEC LTAR projects/'
  fileExt <- '.xlsx'
  defaultFile <- paste(path, 'Soil_or_Plant', fileExt, sep = '')
  filename <- choose.files(default = defaultFile)
  
  # xlsx provides read/write functions for Excel files
  library(xlsx)
  excelSheet <- read.xlsx2(filename, sheetIndex = 1, stringsAsFactors = FALSE)
  excelSheet  # Return value
}


################################################################################
#
# Convert an Excel date/time code to an R date object
#
################################################################################

excelDateToR <- function(excelDate) {
  rDate <- as.numeric(excelDate)
  as.Date(rDate, origin="1899-12-30")  #Return values
}


################################################################################
#
# Extract numeric month, day and year from an R date object
#
################################################################################

extractMDY <- function(rDate) {
  monthNum <- as.numeric(format(rDate, '%m'))
  dayNum <- as.numeric(format(rDate, '%d'))
  yearNum <- as.numeric(format(rDate, '%Y'))
  data.frame(sampYear = yearNum, sampMonth = monthNum, sampDay = dayNum)
}


################################################################################
#
# Handle fatal errors with message output
#
################################################################################

fatalError <- function (errorMessage) {
  cat('\n\n')
  fullMessage <- paste(errorMessage, '. Program halted.', sep = '')
  stop(fullMessage, call. = FALSE)
}


#----------------------------------------------
#
#Compare treatment values in soil and plant files to ensure consistency

#Read files
soil <- chooseExcelFile()
plant <- chooseExcelFile()

#Transform date codes into R-formatted dates
rDate <- excelDateToR(soil$sampDate)

dateDF <- extractMDY(rDate)  #Extract numeric month, day and year values
soil <- cbind(dateDF, soil)  #Add numeric columns to soil df
soil$sampDate <- NULL  #Delete sampDate column

#Coerce plant date values to numeric for comparison with corresponding soil values
plant$sampYear <- as.numeric(plant$sampYear)
plant$sampMonth <- as.numeric(plant$sampMonth)
plant$sampDay <- as.numeric(plant$sampDay)

#Coerce treatment rate values to numeric, rounding to 3 decimal places
soil$trtRateN1_kg_per_ha <- as.numeric(soil$trtRateN1_kg_per_ha) %>% round(, 3)
plant$trtRateN1_kg_per_ha <- as.numeric(plant$trtRateN1_kg_per_ha) %>% round(, 3)
soil$trtRateN2_kg_per_ha <- as.numeric(soil$trtRateN2_kg_per_ha) %>% round(, 3)
plant$trtRateN2_kg_per_ha <- as.numeric(plant$trtRateN2_kg_per_ha) %>% round(, 3)

#Form unique lists of appropriate variables
years <- c(2000:2013)  #Ignore 2014 because of E/W fert split, and ignore 1999.
plots <- unique(soil$plotNumber)
#suffixes <- unique(soil$plotSuffix)

library(dplyr)
#Loop over years and plot numbers
for(yr in years) {
  for(pn in plots) {
    
    plantSub <- dplyr::filter(plant, sampYear == yr, plotNumber == pn)
    soilSub <- dplyr::filter(soil, sampYear == yr, plotNumber == pn)
    
    #Check for trtLevel mismatch or multiple values
    lapply(soilSub$trtLevel, function(x) {
      if(x != plantSub$trtLevel[1] | length(unique(plantSub$trtLevel)) != 1)
        fatalError(paste('trtLvl', yr, pn, x, plantSub$trtLevel[1]))
    }) #Close lapply statement
    
    #Check for trtType1 mismatch
    lapply(soilSub$trtType1, function(x) {
      if(x != plantSub$trtType1[1] | length(unique(plantSub$trtType1)) != 1)
        fatalError(paste('trtType1', yr, pn, x, plantSub$trtType1[1]))
    })
    
    #Check for trtType2 mismatch
    lapply(soilSub$trtType2, function(x) {
      if(x != plantSub$trtType2[1] | length(unique(plantSub$trtType2)) != 1)
        fatalError(paste('trtType2', yr, pn, x, plantSub$trtType2[1]))
    })
    
    #Check for trtRate1 mismatch
    lapply(soilSub$trtRateN1_kg_per_ha, function(x) {
      if(round(x, 3) != round(plantSub$trtRateN1_kg_per_ha[1], 3) |
         length(unique(plantSub$trtRateN1_kg_per_ha)) != 1)
        fatalError(paste('trtRate1', yr, pn, x, plantSub$trtRateN1_kg_per_ha[1]))
    })
    
    #Check for trtRate2 mismatch
    lapply(soilSub$trtRateN2_kg_per_ha, function(x) {
      if(round(x, 3) != round(plantSub$trtRateN2_kg_per_ha[1], 3) |
         length(unique(plantSub$trtRateN2_kg_per_ha)) != 1)
        fatalError(paste('trtRate2', yr, pn, x, plantSub$trtRateN2_kg_per_ha[1]))
    })
  }  #Close inner for-loop
}  #Close outer for-loop


