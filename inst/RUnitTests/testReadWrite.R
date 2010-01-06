###{{ setup

  # xyDir variables from 'runner.R' or define manually in .GlobalEnv
rfile <- file.path( dataDir, "origData.xls" )
wfile <- file.path( outputDir, "tmpWriteData.xls" )

###}}
###{{ test: read and write/read

test.readWriteDouble <- function() {
  myval <- cbind( c( 42.345000000000000,2.131852701746450,5.196028573821160,4.052032658276100,5.455442814521230,0.920121101217604,4.320537450438270,9.228986838061070,3.495177339311360,7.162184977985050,0.079735376876580,0.190438423685664,6.947886357716580 ),
                 c( 3.395006816620560,3.688808025310860,1.182829713821520,5.286272675203200,2.772115393878420,4.547634389328750,9.723958473331380,4.811160580044460,6.250465426071360,9.579865360317900,6.635986679801000,2.808309283842360,6.869654789485020 ),
                 c( 3.800523340130150,3.147041728858900,7.401595346568170,4.757634840096100,7.530930520646660,2.108720283363290,8.516694933737470,2.070764974017880,9.260555562457060,8.930731727676670,6.070516769429020,8.473031283814710,4.987162411190200 ) )
  rdata <- read.xls( rfile, sheet = "dSht", type = "double", from = 2 )
  checkEquals( rdata, myval, check.attributes = FALSE )
  checkIdentical( ncol( rdata ), ncol( myval ) ); checkIdentical( nrow( rdata ), nrow( myval ) )
  checkIdentical( colnames( rdata ), c( "Kol1", "Kol2", "Kol3") )
  
  write.xls( rdata, wfile )
  wdata <- read.xls( wfile, type = "double" )
  checkIdentical( wdata, rdata )
}
  
test.readWriteInteger <- function() {
  myidx <- c( 1, 5, 19 )
  myval <- rbind( c( 1L, 2L, 3L, NA, NA, NA, NA, NA, NA, NA, NA, NA ),
                  c( 1L, 2L, 3L, 4L, 5L, 6L, 7L, 8L, 9L, 10L, 11L, 12L ),
                  c( 0L, 22L, 33L, NA, NA, NA, NA, NA, NA, NA, NA, NA ) )
  rdata <- read.xls( rfile, sheet = "intSht", type = "integer" )
  checkEquals( rdata[myidx,], myval, check.attributes = FALSE )
  checkIdentical( ncol( rdata ), ncol( myval ) ); checkIdentical( nrow( rdata ), 19L )
  checkIdentical( colnames( rdata ), c( "X1", "X2", "X3", rep( "X", 9 ) ) )

  write.xls( rdata, wfile )
  wdata <- read.xls( wfile, type = "integer"  )
  checkIdentical( wdata, rdata )
}
  
test.readWriteLogical <- function() {
  myval <- cbind( c( T, T, T, F, F, F, T, T, T, F, F ), rep( F, 11 ) )
	# FIXME: 0.4 gives false which is a BUG. I casted to integer and then to logical.
  rdata <- read.xls( rfile, colNames = FALSE, sheet = "logSht", type = "logical", from = 4 )
  checkEquals( rdata, myval, check.attributes = FALSE )
  checkIdentical( ncol( rdata ), ncol( myval ) ); checkIdentical( nrow( rdata ), nrow( myval ) )
  checkIdentical( colnames( rdata ), NULL )

  write.xls( rdata, wfile, colNames = FALSE )
  wdata <- read.xls( wfile, colNames = FALSE, type = "logical" )
  checkIdentical( wdata, rdata )
}

test.readWriteCharacter <- function() {
	myval <- matrix( c( 
	  "Sind hierorts Häuser grün, tret ich noch in ein Haus.", "Sind hier die Brücken heil, geh ich auf gutem Grund.", 
    "I'd agree with that,' said Arkady.", "The world, if it has a future, has an ascetic future." ), ncol = 2 )
# TODO: the 'ü' literal above doesn't work (with RMate). Temporary removed in equals check below
  rdata <- read.xls( rfile, colNames = TRUE, "charSht", "character" )
  checkEquals( rdata[,2], myval[,2], check.attributes = FALSE )
  checkIdentical( colnames( rdata ), c( "Bachmann", "Chatwin" ) )
  checkIdentical( ncol( rdata ), ncol( myval ) ); checkIdentical( nrow( rdata ), nrow( myval ) )
  checkIdentical( colnames( rdata ), c( "Bachmann", "Chatwin" ) )
  
  write.xls( rdata, wfile, colNames = TRUE )
  wdata <- read.xls( wfile, TRUE, 1, "character" )
  checkIdentical( wdata, rdata )
}

test.readWriteDataFrame <- function() {
  myval <- data.frame( Fertility = c(80.2, 83.1, 92.5, 85.8, 76.9, 76.1, 83.8, 92.4, 82.4, 82.9, 87.1, 64.1), 
    Agriculture = c(17, 45.1, 39.7, 36.5, 43.5, 35.3, 70.2, 67.8, 53.3, 45.2, 64.5, 62), 
    Testlogical = rep( c(T,T,F),4), 
    Education = c(12, 9, 5, 7, 15, 7, 7, 8, 7, 13, 6, 12), 
    Catholic = c(9.96, 84.84, 93.4, 33.77, 5.16, 90.57, 92.85, 97.16, 97.67, 91.38, 98.61, 8.52), 
    Infant.Mortality = c(22.2, 22.2, 20.2, 20.3, 20.6, 26.6, 23.6, 24.9, 21, 24.4, 24.5, 16.5), 
    Testcharacter = c("Co", "De", "Fr", "Mo", "Ne", "Po", "Br", "Gl", "Gr", "Sa", "Ve", "Ai" ), stringsAsFactors = FALSE )
  myclsIn <-  c( "double", "double", "logical", "integer", "double", "double", "character" )
  rdata <- read.xls( rfile, colNames = TRUE, "dfSht", from = 5, colClasses = myclsIn )
  checkEquals( rdata, myval, check.attributes = FALSE )
  checkIdentical( ncol( rdata ), ncol( myval ) ); checkIdentical( nrow( rdata ), nrow( myval ) )
  checkIdentical( colnames( rdata ), colnames( myval ) )
  checkIdentical( rownames( rdata ), c( "Courtelary", "Delemont", "Franches-Mnt", "Moutier", "Neuveville", "Porrentruy", "Broye", "Glane", "Gruyere", "Sarine", "Veveyse", "Aigle") )

  write.xls( rdata, wfile )
  wdata <- read.xls( wfile, colClasses = myclsIn )
  checkIdentical( wdata, rdata )
}

# more data.frame tests in 'testColClasses.R'

###}}
