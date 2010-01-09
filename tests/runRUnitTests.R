#
# runRUnitTests.R
#
# Executes tests in '../inst/RUnitTests'.

pkg <- "xlsReadWrite"
shlib <- system.file("libs",
		          paste(pkg, .Platform$dynlib.ext, sep=""),
              package = pkg)
stopifnot(file.exists(shlib))

if (file.info( shlib )$size < 20000) {
 ## message not printed on the console, only appears in the log
 message( "tests not executed (cran placeholder shlib)" )
} else if (require( "RUnit", quietly = TRUE )) {
 rutdir <- system.file("RUnitTests", package = pkg)
 stopifnot(file.exists(rutdir), file.info(rutdir)$isdir)
 source(system.file("RUnitTests", "runner.R", package = pkg))
 runTestsForCHECK(rutdir)
} else {
 stop( "tests not executed ('RUnit' package is required)" )
}