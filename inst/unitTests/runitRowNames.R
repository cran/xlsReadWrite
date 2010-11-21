### true and given (false already tested in runitReadWrite.R)

test.rowNames.true <- function() {
    myrow <- c(42.345, 2.1318527, 5.1960286, 4.0520327, 5.4554428, 0.9201211, 4.3205375, 9.2289868, 3.4951773, 7.162185, 0.0797354, 0.1904384, 6.9478864)
    rdata <- read.xls(rfile, rowNames = TRUE, type = "double")
    checkIdentical(rownames(rdata), as.character(myrow))
    checkEquals(rdata[2,3], -22.6819437)
    write.xls(rdata, wfile, rowNames = TRUE)
    wdata <- read.xls(wfile, type = "double")
    checkIdentical(wdata, rdata)

    rdata <- read.xls(rfile, rowNames = TRUE, type = "data.frame")
    checkIdentical(rownames(rdata), as.character(myrow))
    checkEquals(rdata[2,3], -22.6819437)
    write.xls(rdata, wfile, rowNames = TRUE)
    wdata <- read.xls(wfile, type = "data.frame")
    checkIdentical(wdata, rdata)
}

test.rowNames.given <- function() {
    myrow <- paste("r", 1:13, sep = "")
    rdata <- read.xls(rfile,  rowNames = myrow, type = "double")
    checkIdentical(rownames(rdata), myrow)
    checkEquals(rdata[2,4], -22.6819437)
    write.xls(rdata, wfile, rowNames = myrow)
    wdata <- read.xls(wfile, type = "double")
    checkIdentical(wdata, rdata)

    rdata <- read.xls(rfile,  rowNames = myrow, type = "data.frame")
    checkIdentical(rownames(rdata), myrow)
    checkEquals(rdata[2,4], -22.6819437)
    write.xls(rdata, wfile, rowNames = myrow)
    wdata <- read.xls(wfile, type = "data.frame")
    checkIdentical(wdata, rdata)
}

test.rowNames.readGivenWrong <- function() {
    myval <- 1:4
    myrow <- paste("r", 1:19, sep = "")
    checkException(read.xls(rfile, from = 2, rowNames = myrow, type = "double"), silent = TRUE)
    checkException(write.xls(myval, wfile, rowNames = myrow), silent = TRUE)  
    checkException(read.xls(rfile, from = 2, rowNames = myrow, type = "data.frame"), silent = TRUE)
    checkException(write.xls(as.data.frame(t(myval)), wfile, rowNames = myrow), silent = TRUE)  
}

test.rowNames.givenEmpty <- function() {
    myrow <- as.character(1:3)

    rdata <- rdata.orig <- read.xls(rfile,  sheet = "spec_pro", rowNames = TRUE, type = "character", from = 4)
    checkIdentical(rownames(rdata), myrow)
    rownames(rdata) <- NULL
    write.xls(rdata, wfile, rowNames = TRUE)
    wdata <- read.xls(wfile, type = "character", rowNames = TRUE)
    checkIdentical(wdata, rdata.orig)

    rdata <- rdata.orig <- suppressWarnings(read.xls(rfile,  sheet = "spec_pro", rowNames = TRUE, type = "data.frame", from = 4))
    checkIdentical(rownames(rdata), myrow)
    rownames(rdata) <- NULL
    write.xls(rdata, wfile, rowNames = TRUE)
    wdata <- suppressWarnings(read.xls(wfile, type = "data.frame", rowNames = TRUE))
    checkIdentical(wdata, rdata.orig)
}

test.rowNames.partlyMissing <- function() {
    myrow <- c("1", "formula", "3", "4", "logical", "oledate", "integer", "oletime", "oledatetime",
               "10", "double", "12", "13", "14", "15", "character/factor", "17")
    rdata <- read.xls(rfile,  sheet = "autoCls", rowNames = TRUE, type = "character", from = 2)
    checkIdentical(rownames(rdata), myrow)
}


### disallow duplicated rownames

test.rowNames.duplicated <- function() {
    checkException(rdata <- read.xls(rfile, sheet = "intSht", rowNames = TRUE, type = "integer"), silent = TRUE)
    checkException(rdata <- read.xls(rfile, sheet = "intSht", rowNames = TRUE, type = "data.frame"), silent = TRUE)
}


### auto read

test.proRowNames.autoReadYes <- function() {
    col <- c("Fertility", "Agriculture", "Testlogical", "Education", "Catholic", "Infant.Mortality", "Testcharacter")
    colwithrowcol <- c("", col)
    row <- c("Courtelary", "Delemont", "Franches-Mnt", "Moutier", "Neuveville", "Porrentruy", "Broye", "Glane", "Gruyere", "Sarine", "Veveyse", "Aigle")

    rdata <- read.xls(rfile, from = 5, sheet = "dfSht", type = "double")
    checkIdentical(colnames(rdata), col)
    checkIdentical(rownames(rdata), row)
    rdata <- read.xls(rfile, from = 5, sheet = "dfSht")
    checkIdentical(colnames(rdata), col)
    checkIdentical(rownames(rdata), row)

    rdata <- read.xls(rfile, from = 6, colNames = colwithrowcol, sheet = "dfSht", type = "double")
    checkIdentical(colnames(rdata), col)
    checkIdentical(rownames(rdata), row)
    rdata <- read.xls(rfile, from = 6, colNames = colwithrowcol, sheet = "dfSht")
    checkIdentical(colnames(rdata), col)
    checkIdentical(rownames(rdata), row)
}

test.proRowNames.autoReadNo <- function() {
    col1 <- c("Courtelary", "X80.2", "X17", "TRUE.", "X12", "X9.96", "X22.2", "Co" )
    row1 <- as.character(1:11)
    col2 <- paste("V", 1:8, sep = "")
    row2 <- as.character(1:13)

    rdata <- read.xls(rfile, from = 6, sheet = "dfSht", type = "double")
    checkIdentical(colnames(rdata), col1)
    checkIdentical(rownames(rdata), row1)
    rdata <- read.xls(rfile, from = 6, sheet = "dfSht")
    checkIdentical(colnames(rdata), col1)
    checkIdentical(rownames(rdata), row1)

    rdata <- read.xls(rfile, from = 5, colNames = FALSE, sheet = "dfSht", type = "double")
    checkIdentical(colnames(rdata), col2)
    checkIdentical(rownames(rdata), row2)
    rdata <- read.xls(rfile, from = 5, colNames = FALSE, sheet = "dfSht")
    checkIdentical(colnames(rdata), col2)
    checkIdentical(rownames(rdata), row2)
}

### auto write

test.proRowNames.autoWriteNo <- function() {
    myval <- cbind( 42:45, 22:25 )
    rownames(myval) <- 1:4

    write.xls(myval, wfile)
    wdata <- read.xls(wfile, type = "double", rowNames = FALSE)
    checkIdentical(ncol(wdata), 2L)
    write.xls(myval, wfile)
    wdata <- read.xls(wfile, rowNames = FALSE)
    checkIdentical(ncol(wdata), 2L)
}

test.proRowNames.autoWriteYes <- function() {
    myval <- cbind( 42:45, 22:25 )
    rownames(myval) <- 2:5

    write.xls(myval, wfile)
    wdata <- read.xls(wfile, type = "double", rowNames = FALSE)
    checkIdentical(ncol(wdata), 3L)
    write.xls(myval, wfile)
    wdata <- read.xls(wfile, rowNames = FALSE)
    checkIdentical(ncol(wdata), 3L)

    write.xls(myval, wfile, colNames = c("colrow","col1", "col2"))
    wdata <- read.xls(wfile, type = "double", rowNames = FALSE)
    checkIdentical(ncol(wdata), 3L)
    write.xls(myval, wfile)
    wdata <- read.xls(wfile, rowNames = FALSE)
    checkIdentical(ncol(wdata), 3L)
}

