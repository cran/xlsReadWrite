###{{ setup

  # xyDir variables from 'runner.R' or define manually in .GlobalEnv
rfile <- file.path( dataDir, "origData.xls" )
wfile <- file.path( outputDir, "tmpWriteData.xls" )


###}}
###{{ test: read colNames

test.readColNamesGiven <- function() {
  mycol <- c( "one", "two", "three" )
  rdata <- read.xls( rfile, colNames = mycol, from = 3, rowNames = FALSE, type = "double" )
  checkIdentical( colnames( rdata ), mycol )
  checkEqualsNumeric( rdata[2,1], 2.13185270174645 )
  rdata <- read.xls( rfile, colNames = mycol, from = 3, rowNames = FALSE, type = "data.frame" )
  checkIdentical( colnames( rdata ), mycol )
  checkEqualsNumeric( rdata[2,1], 2.13185270174645 )
}

test.readColNamesGivenWrong <- function() {
  mycol <- c( "one", "two" )
  checkException( read.xls( rfile, colNames = mycol, from = 3, rowNames = FALSE, type = "double" ), silent = TRUE )
  checkException( read.xls( rfile, colNames = mycol, from = 3, rowNames = FALSE, type = "data.frame" ), silent = TRUE )
}

  # TODO: for empty and numeric colnames the behaviour is different (also depending
  #       on 'checkNames' argument. Especially (my hardcoded) V1, V2 should be checked
  #       and maybe replaced with 'X' or '""'. What does the read.table function do?
  
test.readColNamesEmpty <- function() {
    # checkNames is TRUE
  rdata <- read.xls( rfile, colNames = TRUE, from = 1, rowNames = FALSE, type = "double" )
  checkIdentical( colnames( rdata ), c( "X", "X", "X" ) )
  checkEqualsNumeric( rdata[3,1], 2.13185270174645 )
  rdata <- read.xls( rfile, colNames = TRUE, from = 1, rowNames = FALSE, type = "data.frame" )
  checkIdentical( colnames( rdata ), c( "V1", "V2", "V3" ) )
  checkEqualsNumeric( as.numeric( as.character( rdata[3,1] ) ), 2.13185270174645 )
    # checkNames is FALSE
  rdata <- read.xls( rfile, colNames = TRUE, from = 1, rowNames = FALSE, type = "double", checkNames = FALSE )
  checkIdentical( colnames( rdata ), c( "", "", "" ) )
  rdata <- read.xls( rfile, colNames = TRUE, from = 1, rowNames = FALSE, type = "data.frame", checkNames = FALSE )
  checkIdentical( colnames( rdata ), c( "V1", "V2", "V3" ) )
}

test.readColNamesNumeric <- function() {
    # checkNames is TRUE
  rdata <- read.xls( rfile, colNames = TRUE, from = 3, rowNames = FALSE, type = "double" )
  checkIdentical( colnames( rdata ), c( "X42.345", "X3.39500681662056", "X3.80052334013015" ) )
  checkEqualsNumeric( rdata[1,1], 2.13185270174645 )
  rdata <- read.xls( rfile, colNames = TRUE, from = 3, rowNames = FALSE, type = "data.frame" )
  checkIdentical( colnames( rdata ), c( "X42.345", "X3.39500681662056", "X3.80052334013015" ) )
  checkEqualsNumeric( rdata[1,1], 2.13185270174645 )
    # checkNames is FALSE
  rdata <- read.xls( rfile, colNames = TRUE, from = 3, rowNames = FALSE, type = "double", checkNames = FALSE )
  checkIdentical( colnames( rdata ), c( "42.345", "3.39500681662056", "3.80052334013015" ) )
  rdata <- read.xls( rfile, colNames = TRUE, from = 3, rowNames = FALSE, type = "data.frame", checkNames = FALSE )
  checkIdentical( colnames( rdata ), c( "42.345", "3.39500681662056", "3.80052334013015" ) )
}

test.readColNamesNormalHasRowNames <- function() {
  rdata <- read.xls( rfile, from = 2, rowNames = TRUE, type = "double" )
  checkIdentical( colnames( rdata ), c( "Kol2", "Kol3") )
  rdata <- read.xls( rfile, from = 2, rowNames = TRUE,  )
  checkIdentical( colnames( rdata ), c( "Kol2", "Kol3") )  
}

test.readColNamesNormalNoRowNames <- function() {
  rdata <- read.xls( rfile, from = 2, type = "double", rowNames = FALSE )
  checkIdentical( colnames( rdata ), c( "Kol1", "Kol2", "Kol3") )
  rdata <- read.xls( rfile, from = 2, rowNames = FALSE )
  checkIdentical( colnames( rdata ), c( "Kol1", "Kol2", "Kol3") )  
}

test.readColNamesNormalAutoRowNames <- function() {
  mycol <- c( "Fertility", "Agriculture", "Testlogical", "Education", "Catholic", "Infant.Mortality", "Testcharacter" )
  rdata <- read.xls( rfile, sheet = "dfSht", from = 5, type = "double" )
  checkIdentical( colnames( rdata ), mycol )
  rdata <- read.xls( rfile, sheet = 5, from = 5, type = "data.frame" )
  checkIdentical( colnames( rdata ), mycol )
}

###}}