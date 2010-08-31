# loads RUnit and provides helper functions to execute tests

### manual execution by USER (see debug.template.R)

    # execute *all* testfiles in 'runitDir' folder
execTestSuite <- function(runitDir, outDir = runitDir) {
    suite <- defineTestSuite(name = "RUnit tests", dirs = runitDir, testFileRegex = "^runit[[:upper:]].+\\.[rR]$")
    res <<- runTestSuite(suite)
    .printResults(res, outDir)
}

    # execute 'testfile' with all or one single 'fctname' functions
execTestFile <- function(testfile, outDir, fctname = NA) {
    runitDir <- dirname(testfile)
    fct <- if (is.na(fctname)) "^test.+" else paste("^", fctname, "$", sep = "")
    if (is.na(outDir)) outDir <- runitDir
    res <- runTestFile(testfile, testFuncRegexp = fct)
    .printResults(res, outDir)
}

### executed from within R CMD CHECK

execTestSuiteCHECK <- function(runitDir) {
    pkg <- sub("\\.Rcheck$", '', basename(dirname(getwd())))
    library(package = pkg, character.only = TRUE)
    .setup(runitDir)  # read.xls must be available when we call this

    suite <- defineTestSuite(name = paste(pkg, "RUnitTests called from CHECK"), dirs = runitDir, testFileRegex = "^runit[[:upper:]].+\\.[rR]$")
    res <- runTestSuite(suite)
    err <- getErrors(res)
    hasErrors <- err$nFail > 0 || err$nErr > 0
    if (hasErrors) {
            # warnings won't be displayed, we have to stop
        stop(.printError(err, "<pkg>.Rcheck/tests/execRUnit.Rout.fail"))
    }
}

### helper function

.printError <- function(err, details) {
    msg <- paste("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n",
                 "  RUnit test failed\n", 
                 "  - ", err$nFail, " failure(s), ", err$nErr, " error(s), ", err$nDeactivated, " deactivated, ", 
                 err$nTestFunc - err$nFail - err$nErr, " successful\n",
                 "  - details in ", details, "\n", 
                 "  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n", sep = "" )
}

.printResults <- function(res, outDir) {
    printTextProtocol(res, showDetails = FALSE, fileName = file.path(outDir, "_Results.txt"))
    printTextProtocol(res, showDetails = TRUE, fileName = file.path(outDir, "_Details.txt"))
    err <- getErrors(res)
    hasErrors <- err$nFail > 0 || err$nErr > 0
    if (hasErrors) {
        .printError(err, normalizePath(outDir))
    } else {
        message("\n")
        message("RUnit test ok: ", err$nTestFunc, " test function(s), ", err$nDeactivated, " deactivated\n")
        message("(Log: ", normalizePath(outDir), ")\n")
    }
}

### setup and global settings

.setup <- function(runitDir, outDir = runitDir){
    assign("rfile", file.path(runitDir, "data/origData.xls"), envir = .GlobalEnv)
    assign("rfile.img", file.path(runitDir, "data/origImage.xls"), envir = .GlobalEnv)
    assign("wfile", file.path(outDir, "tmpWriteData.xls"), envir = .GlobalEnv)
    assign("wfile.img", file.path(outDir, "tmpImageOut.xls"), envir = .GlobalEnv)
    assign("isFreeVersion", length(grep("cells", names(formals(read.xls)))) == 0, envir = .GlobalEnv)
}

library(RUnit)
library(tools)

