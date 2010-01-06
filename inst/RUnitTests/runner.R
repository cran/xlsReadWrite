#
# 'runner.R' is used to easily invoke the runit tests. Source it and
#
# - runTest           run a single 'testfile' (optionally with fctname)
# - runTests          run ~all testfiles in the 'dir' (RUnitTests) directory
# - runTestsForCHECK  only use for CHECK (stops on errors and makes file layout/naming assumptions)

library( RUnit )
library( tools )

runTest <- function( testfile, dir = dirname( testfile ), 
    dataDir = file.path( dir, "data" ), outputDir = file.path( dir, "output"),
    fctname = "^test.+" )
{
  fct <- if (is.na( fctname )) "^test.+" else paste( "^", fctname, "$", sep = "" )
  res <- runTestFile( .checkAndSetup( file.path( dir, testfile ), dataDir, outputDir ), testFuncRegexp = fct )
  .printResults( res )
}

runTests <- function( dir, 
    dataDir = file.path( dir, "data" ), outputDir = file.path( dir, "output") )
{
  .checkAndSetup( dir, dataDir, outputDir )

  suite <- defineTestSuite( name = "RUnit tests", dirs = dir, testFileRegex = "^test[[:upper:]].+\\.[rR]$" )
  res <<- runTestSuite( suite )
  .printResults( res )
}

runTestsForCHECK <- function( dir ) {
  .checkAndSetup( dir, file.path( dir, "data" ), file.path( dir, "output") )
	pkg <- sub( "\\.Rcheck$", '', basename( dirname( getwd() ) ) )
  library( package = pkg, character.only = TRUE )

  suite <- defineTestSuite( name = paste( pkg, "RUnitTests called from CHECK"), dirs = dir, testFileRegex = "^test[[:upper:]].+\\.[rR]$" )
  res <- runTestSuite( suite )
  tmp <- getErrors( res )
  hasErrors <- tmp$nFail > 0 || tmp$nErr > 0
  if (hasErrors) {
  	  # warnings won't be displayed, we have to stop
    stop( "RUnitTests failed:\n->", tmp$nFail, " failure(s), ", tmp$nErr, " error(s))" )
  }
}

.checkAndSetup <- function( target, dataDir, outputDir ) {
  stopifnot( file.exists( target ) )
  stopifnot( file.exists( dataDir ) )
  if (!file.exists( outputDir )) dir.create( outputDir ) 
  if (!exists( "dataDir", envir = .GlobalEnv )) assign( "dataDir", dataDir, envir = .GlobalEnv )
  if (!exists( "outputDir", envir = .GlobalEnv )) assign( "outputDir", outputDir, envir = .GlobalEnv )
  target
}
.printResults <- function( res ) {
  printTextProtocol( res, showDetails = FALSE, fileName = file.path( outputDir, "Results.txt" ) )
  printTextProtocol( res, showDetails = TRUE, fileName = file.path( outputDir, "Details.txt" ) )
  tmp <- getErrors( res )
  hasErrors <- tmp$nFail > 0 || tmp$nErr > 0
  if (hasErrors) {
    message( "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n" )
    message( "RUnit test failed:\n\n-> ", tmp$nFail, " failure(s), ", tmp$nErr, " error(s)" )
    message( "-> see '", normalizePath( outputDir ), "'\n" )
    message( "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n" )
  } else {
    message( "\n" )
    message( "Ok (", tmp$nTestFunc, " test function(s), ", tmp$nDeactivated, " deactivated)\n" )
    message( "-> info see '", normalizePath( outputDir ), "'\n" )
  }
}