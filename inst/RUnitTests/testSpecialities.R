###{{ setup

  # xyDir variables from 'runner.R' or define manually in .GlobalEnv
rfile <- file.path( dataDir, "origData.xls" )
wfile <- file.path( outputDir, "tmpWriteData.xls" )

###}}
###{{ tests

test.readRKCellValue <- function() {
    # in an old (library) version RKRecord value was not divided by 100
  rdata <- read.xls( rfile, sheet = "specialities", type = "double", from = 12 )[1]
  checkIdentical( rdata, 143.28 )
}

test.read1ColDataFrame <- function() {
  myval <- c( "Courtelary", "Delemont", "Franches-Mnt", "Moutier", "Neuveville", "Porrentruy",
            "", "", "OLD BUG WITH RKRecord:", "143.28" )
  rdata <- read.xls( rfile, FALSE, sheet = "specialities", from = 4, stringsAsFactor = FALSE )
  checkIdentical( rdata[,1], myval )
  
  write.xls( rdata, wfile, colNames = FALSE )
  wdata <- read.xls( wfile, FALSE, stringsAsFactor = FALSE )
  checkIdentical( rdata, wdata )
}  

test.readWithBadFrom <- function() {
	checkException( read.xls( rfile, from = "bad" ), silent = TRUE )
	  # check if wrong type has been catched correctly
	checkEquals( grep( "must be a scalar integer or double", geterrmessage() ), 1 )
}


###}}