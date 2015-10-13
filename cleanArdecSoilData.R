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

library(dplyr); library(xlsx)

# Read raw data file
rawDataPath <- 'W:/R/SPNR/ARDEC projects/'
#rawDataPath <- 'C:/Users/Robert/Documents/R/ARDEC/'
rawDataFile <- 'R1W_Soil_Lite_Raw.xlsx'
fullRawFilename <- paste(rawDataPath, rawDataFile, sep = '')
# Worksheet 1 contains nitrate values; worksheet 2 contains ammonium values
rawData <- read.xlsx2(fullRawFilename, sheetIndex = 2, stringsAsFactors = FALSE)

# Remove rows with missing values for trt
rawData <- filter(rawData, !(is.na(trt) | trt == ''))

# Depth increment lists to be used in tidy DF
depthTopList <- rep(c(0, 3, 6, 12, 24, 36, 48, 60), nrow(rawData))
depthBottomList <- rep(c(3, 6, 12, 24, 36, 48, 60, 72), nrow(rawData))

depthTopVec <- as.vector(depthTopList)
depthBottomVec <- as.vector(depthBottomList)

# Number of columns in rawData
rawDataCols <- ncol(rawData)
# Initialize vector for ammonium values
nh4Complete <- as.vector(numeric())

# Construct a vector of ordered nh4 values
for(rowNum in 1:nrow(rawData)) {
  nh4SubVec <- rawData[rowNum, 7:rawDataCols]
  nh4Complete <- append(nh4Complete, nh4SubVec)
}

# Initialize currentDate
currentDate <- rawData$sampDate[1]
# Populate entire sampDate column in rawData
for(rowNum in 1:nrow(rawData)) {
  if(is.na(rawData$sampDate[rowNum]) | rawData$sampDate[rowNum] == '') {
    rawData$sampDate[rowNum] <- currentDate
  } else {
    currentDate <- rawData$sampDate[rowNum]
  }
}

# Initialize vectors
dateVector <- as.vector(character())
trtVector <- as.vector(integer())
rateVector <- as.vector(numeric())
repVector <- as.vector(integer())
plotVector <- as.vector(integer())

# Populate vectors
for(rowNum in 1:nrow(rawData)) {
  dateVector <- append(dateVector, rep(rawData$sampDate[rowNum], 8))
  trtVector <- append(trtVector, rep(rawData$trt[rowNum], 8))
  rateVector <- append(rateVector, rep(rawData$rateN[rowNum], 8))
  repVector <- append(repVector, rep(rawData$rep[rowNum], 8))
  plotVector <- append(plotVector, rep(rawData$plot[rowNum], 8))
}

# Coerce to DF
tidyData <- as.data.frame(cbind(dateVector, trtVector, rateVector, repVector,
                                plotVector, depthTopVec, depthBottomVec,
                                nh4Complete))
# Rename columns
names(tidyData) <- c('sampDate', 'trtLevel', 'trtRate', 'rep', 'plotNumber',
                     'depthTop', 'depthBottom', 'ammonium')

# Remove rows that are missing a ammonium value (some depths not sampled)
tidyData <- filter(tidyData, !(is.na(ammonium) | ammonium == ''))

# Convert Excel date code into an R time value
tidyData$sampDate <- as.POSIXct(as.numeric(tidyData$sampDate) * (60*60*24),
           origin="1899-12-30", tz="GMT")

#tidyDataPath <- 'C:/Users/Robert/Documents/R/ARDEC/'
tidyDataPath <- rawDataPath
tidyDataFile <- 'R1W_Soil_Lite_Tidy_NH4.xlsx'
fullTidyFilename <- paste(tidyDataPath, tidyDataFile, sep = '')
saveDF(tidyData, fullTidyFilename)


