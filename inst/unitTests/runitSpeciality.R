### misc. old bugs

test.speciality.RKCellValue <- function() {
        # in an old flexcel version RKRecord values were not divided by 100
    rdata <- read.xls(rfile, sheet = "spec_1", type = "double", from = 12)[1]
    checkIdentical(rdata, 143.28)
}

test.speciality.1ColDataFrame <- function() {
    myval <- c("Courtelary", "Delemont", "Franches-Mnt", "Moutier", "Neuveville", "Porrentruy",
               "", "", "OLD BUG WITH RKRecord:", "143.28")
    rdata <- read.xls(rfile, FALSE, sheet = "spec_1", from = 4, stringsAsFactor = FALSE)
    checkIdentical(rdata[,1], myval)

    write.xls(rdata, wfile, colNames = FALSE)
    wdata <- read.xls(wfile, FALSE, stringsAsFactor = FALSE)
    checkIdentical(rdata, wdata)
}

test.speciality.readWithBadFrom <- function() {
    checkException(read.xls(rfile, from = "bad"), silent = TRUE)
}


### relative paths

test.speciality.readRelativePathNames <- function() {
    oldwd <- getwd()
    rf1 <- "origData.xls"
        # (the following two paths may fail for non-standard dataDir)
    rf2 <- "..\\..\\unitTests\\data\\origData.xls"
    rf3 <- "../../unitTests/data/origData.xls"
    setwd(dirname(rfile))

    rdata <- read.xls(rf1)
    rdata <- read.xls(rf2)
    rdata <- read.xls(rf3)

    setwd(oldwd)
}

