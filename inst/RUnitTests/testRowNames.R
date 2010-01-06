###{{ setup

  # xyDir variables from 'runner.R' or define manually in .GlobalEnv
rfile <- file.path( dataDir, "origData.xls" )
wfile <- file.path( outputDir, "tmpWriteData.xls" )

###}}
###{{ test: read rowNames

test.readRowNamesTrue <- function() {
  rdata <- read.xls( rfile, from = 2, rowNames = TRUE, type = "double" )
	checkTrue( rownames( rdata )[1] == 42.345 )
  checkIdentical( colnames( rdata ), c( "Kol2", "Kol3") )
}

test.readRowNamesGiven <- function() {
  myrow <- paste( "r", 1:13, sep = "" )
  rdata <- read.xls( rfile, from = 2, rowNames = myrow, type = "double" )
  checkIdentical( rownames( rdata ), myrow )
  rdata <- read.xls( rfile, from = 2, rowNames = myrow, type = "data.frame" )
  checkIdentical( rownames( rdata ), myrow )
}

test.readRowNamesGivenWrong <- function() {
  myrow <- paste( "r", 1:19, sep = "" )
  checkException( read.xls( rfile, from = 2, rowNames = myrow, type = "double" ), silent = TRUE )
  checkException( read.xls( rfile, from = 2, rowNames = myrow, type = "data.frame" ), silent = TRUE )
}

test.readRowlNamesNormal <- function() {
  myrow <- c( "Courtelary", "Delemont", "Franches-Mnt", "Moutier", "Neuveville", 
              "Porrentruy", "Broye", "Glane", "Gruyere", "Sarine", "Veveyse", "Aigle" )
  rdata <- read.xls( rfile, sheet = 5, from = 5, rowNames = TRUE, type = "double" )
  checkIdentical( rownames( rdata ), myrow )
  rdata <- read.xls( rfile, sheet = 5, from = 5, rowNames = TRUE, type = "data.frame" )
  checkIdentical( rownames( rdata ), myrow )
}

###}}

###{{ test: write rowNames

xMatrix1 <- xMatrix2 <- matrix( 1:12, 3, 4, dimnames = list( letters[1:3], LETTERS[1:4] ) )

test.writeRowNamesTrueForMatrix <- function() {
	write.xls( xMatrix1, wfile, colNames = TRUE, rowNames = TRUE )
  rdata <- read.xls( wfile, type = "integer" )
  checkIdentical( xMatrix1, rdata )
}

test.writeRowNamesTrueForBareMatrix <- function() {
	rownames(xMatrix2) <- NULL
  mywarn <- 0
  withCallingHandlers( write.xls( xMatrix2, wfile, colNames = TRUE, rowNames = TRUE ), 
      warning = function(w) { mywarn <<- mywarn + 1; invokeRestart("muffleWarning") } )
  checkIdentical( mywarn, 1 )
  rdata <- read.xls( wfile, type = "integer", rowNames = TRUE )
	rownames(xMatrix2) <- as.character( 1:3 )
  checkIdentical( xMatrix2, rdata )
}

test.writeRowNamesTrueForFrame <- function() {
  xFrame <- data.frame( xMatrix1 )
	write.xls( xFrame, wfile, colNames = TRUE, rowNames = TRUE )
  rdata <- read.xls( wfile, colClasses = "integer" )
  identical( xFrame, rdata )
  checkIdentical( xFrame, rdata )
}

test.writeIntegerRowNamesTrue <- function() {
	x <- data.frame( BD = 1:3, BF = 4:6, BO = 12:14 )
	# (earlier this broke because rownames ended as (unsupported) integer types
	write.xls( x, wfile, rowNames = TRUE )
}

###}}
