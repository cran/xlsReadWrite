# TODO: not yet integrated in RUnit

    # check user colNames (note: "from" has been increased by one in order to point to the first data row)
  res1 <- read.xls( rfile, colNames = c( "MyCol1", "MyCol2", "MyCol3" ), "doubleSheet", "double", from = 3 )
  if (!isTRUE( all.equal( res1, orig, check.attributes = FALSE ) )) stop( "double read/write (data not equal)" )
  if (!all( colnames( res1 ) == c( "MyCol1", "MyCol2", "MyCol3" ) )  ) stop( "double read/write (data not equal)" )
  write.xls( res1, wfile, colNames = TRUE )
  res2 <- read.xls( wfile, TRUE,, "double" )
  if (!isTRUE( all.equal( res1, res2 ))) stop( "double read/write (data not equal)" )

    # habe colNames auf FALSE gesetzt
  rdata <- read.xls( rfile, colNames = TRUE, 3, "logical" )


  # not sure if to put into runit.DateTime
  # isodatetime, isodate and isotime
cc <- c( "numeric", "isodate", "numeric", "numeric", "numeric", "isotime", "isodatetime" )
res1 <- read.xls( rfile, sheet = "specialities", colClasses = cc, from = 15 )
if (!all( colnames( res1 ) == c( "DateAsNumber", "DateAsDate", "Hour", "Minute", "Sec", "TimeAsTime", "DateTimeAsDateTime" ) )  ) stop( "date/time (data not equal)" )
if (!dateTimeToStr( res1$DateAsNumber[7], "YYYY-MM-DD" ) == "2005-03-30" ) stop( "date/time (data not equal)" )
if (!res1$DateAsDate[1] == "2005-03-18" ) stop( "date/time (data not equal)" )
if (!res1$TimeAsTime[19] == "11:00:31" ) stop( "date/time (data not equal)" )
if (!res1$DateTimeAsDateTime[14] == "2005-04-08 10:26:57" ) stop( "date/time (data not equal)" )

#  # check double/integer/character
#res1 <- read.xls( rfile, colClasses = c( "double", "integer", "character" ), from = 2 )
#if (!isTRUE( all.equal( res1$Kol1, orig[,1] ) )) stop( "frame read colClasses/rowNames/colNames (data not equal)" )
#if (!isTRUE( all.equal( res1$Kol2, as.integer(orig[,2]) ))) stop( "frame read colClasses/rowNames/colNames (data not equal)" )
#if (!isTRUE( all.equal( res1$Kol3, as.character(orig[,3]) ))) stop( "frame read colClasses/rowNames/colNames (data not equal)" )
#
#  # check scalar colClass (character)
#res1 <- read.xls( rfile, colClasses = "character", from = 2 )
## don't check first column (the 11th row is different (decimal places))
#if (!isTRUE( all.equal( res1$Kol2, as.character(orig[,2]) ))) stop( "frame read colClasses/rowNames/colNames (data not equal)" )
#if (!isTRUE( all.equal( res1$Kol3, as.character(orig[,3]) ))) stop( "frame read colClasses/rowNames/colNames (data not equal)" )
#
#  # check logical/NA/character  (for logical, numbers will be truncated first)
#res1 <- read.xls( rfile, colClasses = c( "logical", "NA", "character" ), from = 2 )
#if (!isTRUE( all.equal( res1$Kol1, as.logical(as.integer(orig[,1])) ))) stop( "frame read colClasses/rowNames/colNames (data not equal)" )
#if (!isTRUE( all.equal( res1$Kol2, orig[,2] ))) stop( "frame read colClasses/rowNames/colNames (data not equal)" )
#if (!isTRUE( all.equal( res1$Kol3, as.character(orig[,3]) ))) stop( "frame read colClasses/rowNames/colNames (data not equal)" )
  