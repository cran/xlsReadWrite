# helper code to execute RUnit tests manually
# - 1) make a copy of this file, e.g. as 'debug.R'
# - 2) adapt myroot (mytest, mywork, withLib, runInvisible)
# - 3) source and execute test (whole suite or subset, ev. run visible)

### settings
myroot <- "V:/swissrRepos/public/xlsReadWrite"  # HS
mytest <- file.path(myroot, "inst/unitTests"); mywork <- getwd()
withLib <- ""  # "free", "pro" or "" (meaning: source code)
runInvisible <- function(func) if (FALSE) invisible(func) else func

### source/load the code
source(file.path(mytest, "loadRUnit.R"))
if (withLib == "free") library(xlsReadWrite) else
if (withLib == "pro") library(xlsReadWritePro) else { stopifnot(withLib == "")
    runInvisible(sapply(dir(file.path(myroot, "R"), full.names = TRUE), source))
    if (!is.null(getLoadedDLLs()$xlsReadWrite)) dyn.unload(getLoadedDLLs()$xlsReadWrite[["path"]])
    runInvisible(dyn.load(file.path(myroot, "src/pas/xlsReadWrite.dll")))
}
.setup(mytest, mywork)
 
### execute tests

# suite
execTestSuite(mytest, mywork)

# or single files
execTestFile(file.path(mytest, "runitColClasses.R"), mywork)
execTestFile(file.path(mytest, "runitColNames.R"), mywork)
execTestFile(file.path(mytest, "runitDateTime.R"), mywork)
execTestFile(file.path(mytest, "runitNaNaN.R"), mywork)
execTestFile(file.path(mytest, "runitReadWrite.R"), mywork)
execTestFile(file.path(mytest, "runitRowNames.R"), mywork)
execTestFile(file.path(mytest, "runitSpeciality.R"), mywork)

# or single function (example)
execTestFile(file.path(mytest, "runitReadWrite.R"), mywork, "test.readWrite.logical")
