### execute RUnit tests:

#   1. source (or copy/paste) this file
#   2. use snippet below to execute tests in current directory
#   (for options see adapt block (rarely needed))


### snippets

if (FALSE) { # prevent execution when file is sourced

# check suite, i.e. all test files
checkSuite()

# check the indicated file
checkFile("runitReadWrite.R")

# check the indicated test-function
checkFunc("runitReadWrite.R", fct = "test.readWrite.integer")

}


### adapt

    # "pro" (xlsReadWritePro), "free" (xlsReadWrite) or "" (dyn.lib, devel)
withLib <- "free"  
    # when set to TRUE the .GlobalEnv will be cleared before running tests
cleanFirst <- FALSE
    # the following is only considered for "", i.e. lowlevel dyn.lib
#pkgroot <- "T:/swissr_repos/xlsReadWrite"
#runInvisible <- TRUE


### helpers

.setup <- function() {
    library(RUnit)
    if (withLib != "") {
        pkg <<- if (withLib == "free") "xlsReadWrite" else if (withLib == "pro") "xlsReadWritePro" else stopifnot(FALSE)
        library(pkg, character.only = TRUE)
        testdir <<- file.path(system.file(package = pkg), "unitTests");
    } else {
        runFct <- function(func) if (runInvisible) invisible(func) else func
        runFct(sapply(dir(file.path(pkgroot, "R"), full.names = TRUE), source))
        if (!is.null(getLoadedDLLs()$xlsReadWrite)) dyn.unload(getLoadedDLLs()$xlsReadWrite[["path"]])
        runFct(dyn.load(file.path(pkgroot, "src/pas/xlsReadWrite.dll")))
        testdir <<- file.path(pkgroot, "inst/unitTests");
    }
    if (cleanFirst) rm(list = ls(), envir = .GlobalEnv)
    rfile <<- file.path(testdir, "data/origData.xls" )
    rfile.img <<- file.path(testdir, "data/origImage.xls" )
    wfile <<- file.path(getwd(), "tmp_wData.xls" )
    wfile.img <<- file.path(getwd(), "tmp_wImage.xls" )
    cfile <<- file.path(getwd(), "tmp_cData.xls" )
    cfile.img <<- file.path(getwd(), "tmp_cImage.xls" )

    isFreeVersion <<- length(grep("cells", names(formals(read.xls)))) == 0
}
.setup()

.printResults <- function(res, outDir) {
    printTextProtocol(res, showDetails = FALSE, fileName = file.path(outDir, "_Results.txt"))
    printTextProtocol(res, showDetails = TRUE, fileName = file.path(outDir, "_Details.txt"))
    err <- getErrors(res)
    hasErrors <- err$nFail > 0 || err$nErr > 0
    if (hasErrors) {
      msg <- paste("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n",
                   " RUnit test failed\n",
                   " - ", err$nFail, " failure(s), ", err$nErr, " error(s), ", err$nDeactivated, " deactivated, ",
                   err$nTestFunc - err$nFail - err$nErr, " successful\n",
                   " - details in ", normalizePath(outDir), "\n",
                   " !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n", sep = "" )
      message("\n", msg, "\n")
    } else {
        message("\n")
        message("RUnit test ok: ", err$nTestFunc, " test function(s), ", err$nDeactivated, " deactivated\n")
        message("(Log: ", normalizePath(outDir), ")\n")
    }
}

checkSuite <- function() {
    suite <- defineTestSuite(name = "RUnit tests", dirs = testdir, testFileRegex = "^runit[[:upper:]].+\\.[rR]$")
    res <- runTestSuite(suite)
    .printResults(res, getwd())
}

checkFile <- function(fn) {
    testfile <- file.path(testdir, fn)
    res <- runTestFile(testfile, testFuncRegexp = "^test.+")
    .printResults(res, getwd())
}

checkFunc <- function(fn, fct) {
    testfile <- file.path(testdir, fn)
    res <- runTestFile(testfile, testFuncRegexp = paste("^", fct, "$", sep = ""))
    .printResults(res, getwd())
}
