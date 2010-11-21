# * this files contains  bugs which have been fixed
#   and do not fit nicely into the normal unit tests
# * each issue should be separated by a '#...' line
#   ending with the redmine ticket number (if exists)
# * shortly describe what went wrong / has been fixed

########################################################### (i711)

# result pointer unprotected too early, resulting - for larger data - in an AV:
# Error in read.xls(fnorig, colNames = FALSE, sheet = 1, type = "data.frame",  : 
#  Unexpected error. Message: Access violation at address 6C7CB98B in module 'R.dll'. Read of address 00000018

# Only affected free version, but we now test both version. The following code
# triggered the AV with xlsReadWrite version 1.5.2 on my ThinkPad, Win7-64.

test.fixedBugs.unprotectAV <- function() {
    x <- matrix(42.42, nrow = 1400, ncol = 29)
    write.xls(x, wfile, colNames = FALSE)
    gctorture(on = TRUE)
    y <- read.xls(wfile, colNames = FALSE, colClasses = "double")
    mydim <- dim(y)
    y <- read.xls(wfile, colNames = FALSE, colClasses = "double")
    mydim <- dim(y)
    y <- read.xls(wfile, colNames = FALSE, colClasses = "double")
    mydim <- dim(y)
    gctorture(on = FALSE)
    checkIdentical(mydim, c(1400L, 29L))
}

###########################################################
