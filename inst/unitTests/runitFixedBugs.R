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

########################################################### (i713)

# integer values outside the integer range (-2147483648..2147483647) have been
# converted to double values wrongly. New behaviour:
# - auto class: numbers are always "double" (as in the free version already)
# - colClasses = "integer" stops with an error if a value is outside allowed range
#                (technically we check in the lowlevel 'VarAsInt' call)
# - (note: in R 'as.integer(-2147483648)' gives NA)

test.fixedBugs.integerCheckRange <- function() {
    x <- c(-2147483649, -2147483648, -2147483647,
           2147483646, 2147483647, 2147483648)
    write.xls(x, wfile, colNames = FALSE)
    # make sure the correct values have been written
    checkIdentical(as.character(read.xls(wfile, colNames = FALSE, type = "character")),
                   as.character(x))
    # double works (and is default for data.frame)
    checkIdentical(as.numeric(read.xls(wfile, colNames = FALSE, type = "double")), x)
    checkIdentical(as.numeric(read.xls(wfile, colNames = FALSE)[,1]), x)
    # integer raises an error
    checkException(read.xls(wfile, colNames = FALSE, type = "integer"), silent = TRUE)

    # note: in R '-2147483648' is not a valid integer (gives NA)
    mywarn <- 0
    withCallingHandlers(y <- as.integer(x[2]), warning = function(w) {
        mywarn <<- mywarn + 1; invokeRestart("muffleWarning") })
    checkIdentical(mywarn, 1)
    checkTrue(is.na(y))
    
    # check the individual values
    if (!isFreeVersion) {
        checkException(read.xls(wfile, cells = c(1,1), type = "integer"), silent = TRUE)
        checkException(read.xls(wfile, cells = c(2,1), type = "integer"), silent = TRUE)
        checkIdentical(read.xls(wfile, cells = c(3,1), type = "integer"), as.integer(x[3]))
        checkIdentical(read.xls(wfile, cells = c(4,1), type = "integer"), as.integer(x[4]))
        checkIdentical(read.xls(wfile, cells = c(5,1), type = "integer"), as.integer(x[5]))
        checkException(read.xls(wfile, cells = c(6,1), type = "integer"), silent = TRUE)
    }
}


test.fixedBugs.integerOutsideAsReported <- function() {
    x <- c(1, 12, 123, 1234, 12345, 123456, 1234567, 12345678, 123456789,
           1234567890, 12345678901, 123456789012)
    write.xls(x, wfile, colNames = FALSE)
    # make sure the correct values have been written
    checkIdentical(as.character(read.xls(wfile, colNames = FALSE, type = "character")),
                   as.character(x))
    # double works (and is default for data.frame)
    checkIdentical(as.numeric(read.xls(wfile, colNames = FALSE, type = "double")), x)
    checkIdentical(as.numeric(read.xls(wfile, colNames = FALSE)[,1]), x)
    # integer raises an error
    checkException(read.xls(wfile, colNames = FALSE, type = "integer"), silent = TRUE)
}

###########################################################
