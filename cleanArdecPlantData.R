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
# This function saves a data frame to an Excel worksheet
#
################################################################################
saveDF <- function(dataFrame, fullFilename) {
  write.xlsx2(dataFrame, file = fullFilename, showNA = FALSE, row.names = FALSE)
}



################################################################################
#
#  Reads and tidies raw data summary files 
#
################################################################################

# Read raw data file
rawDataPath <- 'W:/R/SPNR/ARDEC projects/'
rawDataFile <- 'ARDEC_R1_Plant_All_Years_Raw.xlsx'
fullRawFilename <- paste(rawDataPath, rawDataFile, sep = '')
rawData <- readWithClasses(fullRawFilename)

# Fill in crop
rawData$crop <- 'Corn'
# Remove rows with missing values for trtType1
rawData <- filter(rawData, !(is.na(trtType1) | trtType1 == ''))

# Form vectors for use in tidyData DF
grainField <- rawData$yieldGrainFieldDry_kg_per_ha
grainOD <- rawData$yieldGrainOvenDry_kg_per_ha
stalkOD <- rawData$yieldStalkOvenDry_kg_per_ha
cobOD <- rawData$yieldCobOvenDry_kg_per_ha
# Create tidyData with enough rows to accommodate all yield values from all
# plant segments in a single column
tidyData <- rbind(rawData, rawData, rawData)
# Fill plantSegment column
tidyData$plantSegment <- c(rep('Grain', nrow(rawData)),
                          rep('Stalk', nrow(rawData)),
                          rep('Cob', nrow(rawData)))
# Populate oven dry yields
tidyData$yieldOvenDry_kg_per_ha <- c(grainOD, stalkOD, cobOD)

# The next column is first established as NA so that the following statement can
# populate just a specific range.
tidyData$yieldFieldDry_kg_per_ha <- NA_real_
tidyData$yieldFieldDry_kg_per_ha[1:nrow(rawData)] <- grainField
tidyData$yieldGrainFieldDry_kg_per_ha <- NULL
tidyData$yieldGrainOvenDry_kg_per_ha <- NULL
tidyData$yieldStalkOvenDry_kg_per_ha <- NULL
tidyData$yieldCobOvenDry_kg_per_ha <- NULL

savePath <- 'W:/R/SPNR/ARDEC projects/'
saveFile <- 'ARDEC_R1_Plant_All_Years_Tidy.xlsx'
fullSaveFilename <- paste(savePath, saveFile, sep = '')
saveDF(tidyData, fullSaveFilename)
