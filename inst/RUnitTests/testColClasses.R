###{{ setup

  # xyDir variables from 'runner.R' or define manually in .GlobalEnv
rfile <- file.path( dataDir, "origData.xls" )
wfile <- file.path( outputDir, "tmpWriteData.xls" )

###}}
###{{ test: read with given colClasses

test.readColClassesGivenDateTime <- function() {
  myidx <- c( 1, 19 )
  myclsIn <-  c( "integer", "integer", "isodate",   "integer", "integer", "double", "double", "double", "isotime",   "double", "double", "isodatetime"  )
  myclsOut <- c( "integer", "integer", "character", "integer", "integer", "double", "double", "double", "character", "double", "double", "character"  )
  myval <- data.frame( OleDate = c(38429L, 38460L), AsDate = c(38429L, 38460L), IsoDate = c("2005-03-18", "2005-04-18"), 
    Hour = c(14L, 11L), Minute = c(42L, 0L), Sec = c(18.005, 31.002), OleTime = c(0.612708391203704, 0.458692152777778), AsTime = c(0.612708391203704, 0.458692152777778), IsoTime = c("14:42:18", "11:00:31"), 
    OleDateTime = c(38429.6127083912, 38460.4586921528), AsDateTime = c(38429.6127083912, 38460.4586921528), IsoDateTime = c("2005-03-18 14:42:18", "2005-04-18 11:00:31"), stringsAsFactors = FALSE )
  rdata <- read.xls( rfile, sheet = "dateTime", colClasses = myclsIn )

  checkIdentical( colnames( rdata ), colnames( myval ) )  
  checkIdentical( as.vector( sapply( rdata, storage.mode ) ), myclsOut )
  for (co in 1:length( myval )) {
    if ( myclsOut[co] == "double" ) chk <- checkEquals else chk <- checkIdentical
    chk( rdata[myidx,co], myval[,co] )
  }
}

test.readColClassesGivenDoubleIntChar <- function() {
  myidx <- c( 1, 13 )
  myval <- matrix( c( 42.34500000, 3, 3.80052334013015,
                       6.94788636, 6, 4.9871624111902 ), ncol = 3, byrow = TRUE )
  mycls <- c( "double", "integer", "character" )
  rdata <- read.xls( rfile, from = 2, colClasses = mycls )
  checkIdentical( as.vector( sapply( rdata, storage.mode ) ), mycls )
  checkEquals( rdata[myidx,1], as.double( myval[,1] ) )
  checkIdentical( rdata[myidx,2], as.integer( myval[,2] ) )
  checkIdentical( rdata[myidx,3], as.character( myval[,3] ) )
}

test.readColClassesGivenLogicalFactorNA <- function() {
  myval <- structure( list(
    dd = c( FALSE, TRUE ), 
    ii = structure( 1:2, .Label = c( "2.80830928384236", "6.86965478948502" ), class = "factor" ), 
    cc = c( 8.4730312838147, 4.9871624111902 ) ), .Names = c( "dd", "ii", "cc" ), class = "data.frame", row.names = c( "1", "2" ) )
	# FIXME: 0.1904 gives false which is a BUG. I casted to integer and then to logical.
  myclsIn <- c( "logical", "factor", "NA" )
  myclsOutClass <- c( "logical", "factor", "numeric" )
  myclsOutStorageMode <- c( "logical", "integer", "double" )
  rdata <- read.xls( rfile, colNames = c( "dd", "ii", "cc" ), from = 14, colClasses = myclsIn )
  checkIdentical( as.vector( sapply( rdata, class ) ), myclsOutClass )
  checkIdentical( as.vector( sapply( rdata, storage.mode ) ), myclsOutStorageMode )
  checkEquals( rdata, myval )
}

###}}
###{{ test: read with auto-determined colClasses

test.readColClassesAuto <- function() {
  myclsOutClass <-    c( "numeric", "numeric", "character", "numeric", "numeric", "numeric", "character" )
  myclsOutStorageMode <- c( "double", "double", "character", "double", "double", "double", "character" )
  # TODO: TRUE should be recognized as logical (does this work on pro?)
  # TODO: Why is the 'Education' not an integer (does this work on pro?)
  rdata <- read.xls( rfile, colNames = TRUE, "dfSht", from = 5, stringsAsFactors = FALSE )
  checkIdentical( as.vector( sapply( rdata, storage.mode ) ), myclsOutStorageMode )  
}

test.readColClassesAutoFirst16Rows <- function() {
  mywarn <- 0
  withCallingHandlers( rdata <- read.xls( rfile, sheet = "autoCls", from = 2, stringsAsFactors = FALSE ), 
    warning = function(w) { mywarn <<- mywarn + 1; invokeRestart("muffleWarning") } )
  checkIdentical( mywarn, 1 )
  checkIdentical( class( rdata$just_ok ), "character" )  
  checkIdentical( rdata$just_ok[16], "Palm" )
  checkIdentical( class( rdata$has_warnings ), "logical" )  
  checkIdentical( rdata$has_warnings[c( 16, 17 )], c( NA, NA ) )
}

test.readColClassesAutoProgression <- function() {
    # the first found cell value determines class
    # (there's a progression on the pro version but it's too complicated to implement here)
    # check first row and...
  myclsOutClass <- c( "factor", rep( "numeric", 8 ), "factor", "logical" )
  myclsOutStorageMode <- c( "integer", rep( "double", 8 ), "integer", "logical" )
  checkIdentical( as.vector( sapply( suppressWarnings( read.xls( rfile, sheet = "autoCls", from = 2 ) ), class ) ), myclsOutClass )
  checkIdentical( as.vector( sapply( suppressWarnings( read.xls( rfile, sheet = "autoCls", from = 2 ) ), storage.mode ) ), myclsOutStorageMode )
    # ...each row separately (on free version)
  mycls <- c( "double", "double", rep( "character", 4 ), rep( "double", 4 ), rep( "character", 6 ) )
  for (ro in 3:18) {
    checkIdentical( storage.mode( suppressWarnings( read.xls( rfile, colNames = FALSE, sheet = "autoCls", from = ro, stringsAsFactors = FALSE ) )[,2] ), mycls[ro - 2] )
  }
}


##}}