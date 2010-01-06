###{{ setup

myd <- c( 38429, 38432, 38433, 38434, 38435, 38440, 38441, 38442, 38443 )

##}}
###{{ test: basic date time functions (tests in 'TestColClasses.R')

test.dateTimeToStr_strToDateTime <- function() {
  s <- dateTimeToStr( myd )
  res <- strToDateTime( s )
  checkEquals( res, myd )
}

test.formattedDateTimeToStr <- function() {
  res <- dateTimeToStr( myd, 'mm-dd-yyyy' )
  checkEquals( res[c(3, 7)], c( "03-22-2005", "03-30-2005" ) )
}

test.dateTimeToIsoStr_isoStrToDateTime <- function() {
  res <- dateTimeToIsoStr( myd )
  res <- isoStrToDateTime( res )
  checkEquals( res, myd )
}

test.isoStrToDateTime <- function() {
  res <- isoStrToDateTime( c( "20070319", "2007-03-19", "20070319233112", "2007-03-19 23:31:12"  ) ) 
  checkEquals( res, c( 39160, 39160, 39160.98, 39160.98 ) )
}

###}}