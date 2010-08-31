### given colClasses

test.colClasses.givenDateTime <- function() {
    myidx <- c(1, 19)
    mycls <-  c("integer", "integer", "isodate",   "integer", "integer", "double", "double", "double", "isotime",   "double", "double", "isodatetime" )
    mystorage <- c("integer", "integer", "character", "integer", "integer", "double", "double", "double", "character", "double", "double", "character" )
    myval <- data.frame(IntDate = c(38429L, 38460L), AsDate = c(38429L, 38460L), AsIsoDate = c("2005-03-18", "2005-04-18"), 
        Hour = c(14L, 11L), Minute = c(42L, 0L), Sec = c(18.005, 31.002), 
        IntTime = c(0.612708391203704, 0.458692152777778), AsTime = c(0.612708391203704, 0.458692152777778),
        AsIsoTime = c("14:42:18", "11:00:31"), 
        IntDateTime = c(38429.6127083912, 38460.4586921528), AsDateTime = c(38429.6127083912, 38460.4586921528),
        AsIsoDateTime = c("2005-03-18 14:42:18", "2005-04-18 11:00:31"), stringsAsFactors = FALSE)
    rdata <- read.xls(rfile, sheet = "dateTime", colClasses = mycls)

    checkIdentical(colnames(rdata), colnames(myval))  
    checkIdentical(as.vector(sapply(rdata, storage.mode)), mystorage)
    for (co in 1:length(myval)) {
        if (mycls[co] == "double") chk <- checkEquals else chk <- checkIdentical
        chk(rdata[myidx,co], myval[,co])
  }
}

test.colClasses.givenDoubleIntCharNum <- function() {
    myidx <- c(1, 13)
    myval <- matrix(c(42.34500000, 3, 3.8005233, -20.4994424,
                       6.9478864, 6, 4.9871624, 4.4270298), ncol = 4, byrow = TRUE)
    mycls <- c("double", "integer", "character", "numeric")
    mystorage <- c("double", "integer", "character", "double")
    rdata <- read.xls(rfile, colClasses = mycls)
    checkIdentical(as.vector(sapply(rdata, storage.mode)), mystorage)
    checkEquals(rdata[myidx,1], myval[,1])
    checkIdentical(rdata[myidx,2], as.integer(myval[,2]))
    checkIdentical(rdata[myidx,3], as.character(myval[,3]))
    checkIdentical(rdata[myidx,4], myval[,4])
}

test.colClasses.givenLogicalFactorNA <- function() {
    myval <- structure(list(
        c(TRUE, TRUE), 
        structure(1:2, .Label = c("2.8083093", "6.8696548"), class = "factor"), 
        c(8.4730312838147, 4.9871624111902),
        c(-35.3584594, 4.4270298)),
        .Names = c("l_1", "f_2", "d_3", "d_4"), class = "data.frame", row.names = 1:2)
    myclsIn <- c("logical", "factor", "NA", NA)
    myclsOut <- c("logical", "factor", "numeric", "numeric")
    mystorage <- c("logical", "integer", "double", "double")
    rdata <- read.xls(rfile, colNames = c("l_1", "f_2", "d_3", "d_4"), from = 13, colClasses = myclsIn)
    checkIdentical(as.vector(sapply(rdata, class)), myclsOut)
    checkIdentical(as.vector(sapply(rdata, storage.mode)), mystorage)
    checkEquals(rdata, myval)
}


### auto-determine colClasses

test.colClasses.auto <- function() {
    if (isFreeVersion) {
      mycls <-    c("numeric", "numeric", "logical", "numeric", "numeric", "numeric", "character")
      mystorage <- c("double", "double", "logical", "double", "double", "double", "character")
    } else {
      mycls <-    c("numeric", "numeric", "integer", "integer", "numeric", "numeric", "character")
      mystorage <- c("double", "double", "integer", "integer", "double", "double", "character")
    }
    rdata <- read.xls(rfile, colNames = TRUE, "dfSht", from = 5, stringsAsFactors = FALSE)
    checkIdentical(as.vector(sapply(rdata, class)), mycls)  
    checkIdentical(as.vector(sapply(rdata, storage.mode)), mystorage)  
}

test.colClasses.autoFirst16Rows <- function() {
    mywarn <- 0
    withCallingHandlers(rdata <- read.xls(rfile, sheet = "autoCls", from = 2, stringsAsFactors = FALSE), 
        warning = function(w) { mywarn <<- mywarn + 1; invokeRestart("muffleWarning") })
    checkIdentical(mywarn, 1)
    checkIdentical(class(rdata$just_ok), "character")
    checkIdentical(rdata$just_ok[16], "and")
    checkIdentical(class(rdata$has_warnings), "logical")  
    checkIdentical(rdata$has_warnings[c(16, 17)], c(NA, NA))
}

test.colClasses.autoProgression <- function() {
        # the first non-empty cell value determines the class
        # (there's a progression on the pro version but it's too complicated to implement here)
    if (isFreeVersion) {
            # check first row and then...
        mycls <-    c(rep("numeric", 8), "factor", "logical")
        mystorage <- c(rep("double", 8), "integer", "logical")
        rdata <- suppressWarnings(read.xls(rfile, sheet = "autoCls", from = 2))
        checkIdentical(as.vector(sapply(rdata, function(x) class(x)[1])), mycls)
        checkIdentical(as.vector(sapply(rdata, storage.mode)), mystorage)
        # ...loop and check each row of 2nd column separately
        mycls <- c("double", "double", rep("logical", 3), rep("double", 6), rep("character", 5))
        for (ro in 3:18) {
            checkIdentical(storage.mode(suppressWarnings(read.xls(
                rfile, colNames = FALSE, sheet = "autoCls", from = ro, 
                stringsAsFactors = FALSE))[,2]), mycls[ro - 2])
        }
    }
}
