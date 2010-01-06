#
# runRUnitTests.R
#
# - executes ~all tests in '../inst/RUnitTests'
# - intended to be run from CHECK only
#   (getwd() is supposed to be in <PATH_TO_PKG>.Rcheck/tests)

wd <- getwd()
pkgdir <- dirname( wd )
pkg <- sub( "\\.Rcheck$", '', basename( pkgdir ) )
shlib <- paste( pkgdir, "/", pkg, "/libs/", pkg, .Platform$dynlib.ext, sep = "" )
stopifnot( file.exists( shlib ) )

if (file.info( shlib )$size < 20000) {
	  # message not printed on the console, only appears in the log
  message( "tests not executed (cran placeholder shlib)" )

} else if (require( "RUnit", quietly = TRUE )) {
  rutdir <- file.path( pkgdir, pkg, "RUnitTests" )
  stopifnot( file.exists( rutdir ), file.info( rutdir )$isdir )
  source( file.path( rutdir, "runner.R") )
  runTestsForCHECK( rutdir )

} else {
  stop( "tests not executed ('RUnit' package is required)" )
}

