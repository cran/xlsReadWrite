### execute RUnit tests from within R CMD check
#
# - the test files are located at: '../inst/unitTests/runit*.R'
# - to execute them manually see:  '../inst/unitTests/execManually.R'

pkg <- "xlsReadWrite"
shlib <- system.file("libs", if (nzchar(arch <- .Platform$r_arch)) arch else "",
                     paste(pkg, .Platform$dynlib.ext, sep = ""), package = pkg)
stopifnot(file.exists(shlib))

if (file.info(shlib)$size < 20000)  {
        # cran version with dummy shlib (msg appears in the log)
    message("tests not executed (cran placeholder shlib)")
} else if (require("RUnit", quietly = TRUE)) {

    # path to folder
    rutdir <- system.file("unitTests", package = pkg)
    stopifnot(file.exists(rutdir), file.info(rutdir)$isdir)

    # load package - setup file paths and flags
    library(package = pkg, character.only = TRUE)
    rfile <<- file.path(rutdir, "data/origData.xls" )
    rfile.img <<- file.path(rutdir, "data/origImage.xls" )
    wfile <<- file.path(getwd(), "tmp_wData.xls" )
    wfile.img <<- file.path(getwd(), "tmp_wImage.xls" )
    cfile <<- file.path(getwd(), "tmp_cData.xls" )
    cfile.img <<- file.path(getwd(), "tmp_cImage.xls" )
    isFreeVersion <<- length(grep("cells", names(formals(read.xls)))) == 0

    # execute all tests
    suite <- defineTestSuite(name = paste(pkg, "RUnitTests called from CHECK"), dirs = rutdir, testFileRegex = "^runit[[:upper:]].+\\.[rR]$")
    res <- runTestSuite(suite)
    err <- getErrors(res)
    hasErrors <- err$nFail > 0 || err$nErr > 0
    if (hasErrors) {
        # warnings won't be displayed, we have to stop
        msg <- paste("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n",
             " RUnit test failed\n",
             " - ", err$nFail, " failure(s), ", err$nErr, " error(s), ", err$nDeactivated, " deactivated, ",
             err$nTestFunc - err$nFail - err$nErr, " successful\n",
             " - details in <pkg>.Rcheck/tests/execRUnit.Rout.fail\n",
             " !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n", sep = "" )
        stop(msg)
    }
} else {
    stop( "'RUnit' package is not installed. Tests can not be executed." )
}
