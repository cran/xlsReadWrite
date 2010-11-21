### setup

readAndCheckNull <- function(cn, type = NA) {
    if (is.na(type)) {
        dat <- read.xls(wfile, colNames = cn)
    } else {
        dat <- read.xls(wfile, colNames = cn, type = type)
    }
    checkIdentical(dat, NULL)
    checkIdentical(colnames(dat), NULL)
    checkIdentical(rownames(dat), NULL)
}


### colNames given

test.colNames.given <- function() {
    mycol <- c("one", "two", "three", "four")
      # double
    rdata <- read.xls(rfile, colNames = mycol, type = "double", from = 2)
    checkIdentical(colnames(rdata), mycol)
    checkEqualsNumeric(rdata[2,1], 2.13185270174645)
    write.xls(rdata, wfile, colNames = mycol)
    wdata <- read.xls(wfile, type = "double")
    checkIdentical(wdata, rdata)
        # data.frame
    rdata <- read.xls(rfile, colNames = mycol, type = "data.frame", from = 2)
    checkIdentical(colnames(rdata), mycol)
    checkEqualsNumeric(rdata[2,1], 2.13185270174645)
    write.xls(rdata, wfile, colNames = mycol)
    wdata <- read.xls(wfile, type = "data.frame")
    checkIdentical(wdata, rdata)
}

test.colNames.givenWrong <- function() {
    myval <- 1:4
    mycol <- c("one", "two")
        # double
    checkException(read.xls(rfile, colNames = mycol, type = "double", from = 2), silent = TRUE)
    checkException(write.xls(myval, wfile, colNames = mycol), silent = TRUE)
        # data.frame
    checkException(read.xls(rfile, colNames = mycol, type = "data.frame", from = 2), silent = TRUE)
    checkException(write.xls(as.data.frame(t(myval)), wfile, colNames = mycol), silent = TRUE)
}

### colNames true but (partly) missing

test.colNames.trueButEmpty <- function() {
        # character
    rdata <- read.xls(rfile, colNames = TRUE, sheet = "spec_1", type = "character", from = 3)
    checkIdentical(colnames(rdata), "V1")
    checkIdentical(rdata[1], "Courtelary")
    colnames(rdata) <- NULL
    write.xls(rdata, wfile, colNames = TRUE)
    wdata <- read.xls(wfile, type = "character", colNames = FALSE)
    checkIdentical(wdata[1,], "V1")
    checkEquals(wdata[-1,,drop = FALSE], rdata, check.attributes = FALSE)
        # data.frame
    rdata <- read.xls(rfile, colNames = TRUE, sheet = "spec_1", type = "data.frame", from = 3)
    checkIdentical(colnames(rdata), "V1")
    checkIdentical(as.character(rdata[1,1]), "Courtelary")
    colnames(rdata) <- NULL
    write.xls(rdata, wfile, colNames = TRUE)
    wdata <- read.xls(wfile, type = "data.frame", colNames = FALSE)
    checkIdentical(as.character(wdata[1,]), "V1")
    checkEquals(wdata[-1,,drop = FALSE], rdata, check.attributes = FALSE)
}

test.colNames.truePartlyEmpty <- function() {
    mycol <- c("Linktype", "Name.and.Link", "Linkaddress", "V4", "V5", "Col1", "Col2")
    mycol.ex <- c("V1", "Linktype", "Name.and.Link", "Linkaddress", "V5", "V6", "Col1", "Col2")
        # character
    rdata <- read.xls(rfile, sheet = "spec_pro", rowNames = TRUE, type = "character", from = 4)
    checkIdentical(colnames(rdata), mycol)
    rdata <- read.xls(rfile,  sheet = "spec_pro", rowNames = FALSE, type = "character", from = 4)
    checkIdentical(colnames(rdata), mycol.ex)
    rdata <- read.xls(rfile,  sheet = "spec_pro", type = "character", from = 4)
    checkIdentical(colnames(rdata), mycol.ex)
        # data.frame
    rdata <- suppressWarnings(read.xls(rfile, sheet = "spec_pro", rowNames = TRUE, type = "data.frame", from = 4))
    checkIdentical(colnames(rdata), mycol)
    rdata <- suppressWarnings(read.xls(rfile,  sheet = "spec_pro", rowNames = FALSE, type = "data.frame", from = 4))
    checkIdentical(colnames(rdata), mycol.ex)
    rdata <- read.xls(rfile,  sheet = "spec_pro", type = "character", from = 4)
    checkIdentical(colnames(rdata), mycol.ex)
}


### numeric colNames

test.colNames.numericWithCheckNames <- function() {
    mycol <- c("X42.345", "X3.3950068", "X3.8005233", "X.20.4994424")
    rdata <- read.xls(rfile, colNames = TRUE, type = "double", from = 2)
    checkIdentical(colnames(rdata), mycol)
    checkIdentical(rdata[1,1], 2.1318527)
    write.xls(rdata, wfile, colNames = TRUE)
    wdata <- read.xls(wfile, type = "double", checkNames = FALSE)
    checkIdentical(wdata, rdata)

    rdata <- read.xls(rfile, colNames = TRUE, type = "data.frame", from = 2)
    checkIdentical(colnames(rdata), mycol)
    checkIdentical(rdata[1,1], 2.1318527)
    write.xls(rdata, wfile, colNames = TRUE)
    wdata <- read.xls(wfile, type = "data.frame", checkNames = FALSE)
    checkIdentical(wdata, rdata)
}

test.colNames.numericWithoutCheckNames <- function() {
    mycol <- c("42.345", "3.3950068", "3.8005233", "-20.4994424")
    rdata <- read.xls(rfile, colNames = TRUE, type = "double", from = 2, checkNames = FALSE)
    checkIdentical(colnames(rdata), mycol)
    write.xls(rdata, wfile, colNames = TRUE)
    wdata <- read.xls(wfile, type = "double", checkNames = FALSE)
    checkIdentical(wdata, rdata)

    rdata <- read.xls(rfile, colNames = TRUE, type = "data.frame", from = 2, checkNames = FALSE)
    checkIdentical(colnames(rdata), mycol)
    write.xls(rdata, wfile, colNames = TRUE)
    wdata <- read.xls(wfile, type = "data.frame", checkNames = FALSE)
    checkIdentical(wdata, rdata)
}


### colnames only (i.e. there is only one row)

test.colNames.only <- function() {
    col = c("col_1", "col_2")
    coldef = c("V1", "V2")
    colmy = c("my_1", "my_2")
    write.xls(t(col), wfile, colNames = FALSE)
        # character
    wdata <- read.xls(wfile, colNames = FALSE, type = "character")
    checkIdentical(colnames(wdata), coldef)
    checkIdentical(as.vector(wdata), col)
    wdata <- read.xls(wfile, colNames = TRUE, type = "character")
    checkIdentical(colnames(wdata), col)
    checkIdentical(as.vector(wdata), character())
    wdata <- read.xls(wfile, colNames = colmy, type = "character")
    checkIdentical(colnames(wdata), colmy)
    checkIdentical(as.vector(wdata), col)
        # data.frame
    wdata <- read.xls(wfile, colNames = FALSE, stringsAsFactors = FALSE)
    checkIdentical(colnames(wdata), coldef)
    checkIdentical(as.character(wdata), col)
    wdata <- read.xls(wfile, colNames = TRUE)
    checkIdentical(colnames(wdata), col)
    checkIdentical(as.character(wdata), c("logical(0)", "logical(0)"))
    wdata <- read.xls(wfile, colNames = colmy, stringsAsFactors = FALSE)
    checkIdentical(colnames(wdata), colmy)
    checkIdentical(as.character(wdata), col)
}


### colnames with empty vectors/matrix/data.frame

test.colNames.emptyVector <- function() {
    x <- double() # ncol(x) is NULL: nothing  will be written
    
    write.xls(x, wfile, colNames = FALSE)
    checkTrue(is.null(read.xls(wfile)))
    checkTrue(is.null(read.xls(wfile, type = "double")))

    write.xls(x, wfile, colNames = TRUE)
    checkTrue(is.null(read.xls(wfile)))
    checkTrue(is.null(read.xls(wfile, type = "double")))

    write.xls(x, wfile, colNames = "myColName")
    checkTrue(is.null(read.xls(wfile)))
    checkTrue(is.null(read.xls(wfile, type = "double")))
}


test.colNames.emptyMatrix <- function() {
    x <- as.matrix(integer()) # ncol(x) is 1: colnames will potentially be written
    
    write.xls(x, wfile, colNames = FALSE)
    checkTrue(is.null(read.xls(wfile)))
    checkTrue(is.null(read.xls(wfile, type = "double")))

    write.xls(x, wfile, colNames = TRUE)
    wdata <- read.xls(wfile, colNames = FALSE, stringsAsFactor = FALSE)
    checkIdentical(wdata[[1]], "V1")
    wdata <- read.xls(wfile, colNames = FALSE, type = "character")
    checkIdentical(wdata[[1]], "V1")

    write.xls(x, wfile, colNames = "myColName")
    wdata <- read.xls(wfile, colNames = FALSE, stringsAsFactor = FALSE)
    checkIdentical(wdata[[1]], "myColName")
    wdata <- read.xls(wfile, colNames = FALSE, type = "character")
    checkIdentical(wdata[[1]], "myColName")
}

test.colNames.emptyDataFrames <- function() {
    x <- data.frame() # ncol(x) is 0: nothing  will be written
    
    write.xls(x, wfile, colNames = FALSE)
    checkTrue(is.null(read.xls(wfile)))
    checkTrue(is.null(read.xls(wfile, type = "double")))
    
    write.xls(x, wfile, colNames = TRUE)
    checkTrue(is.null(read.xls(wfile)))
    checkTrue(is.null(read.xls(wfile, type = "double")))
    
    write.xls(x, wfile, colNames = "custCol")
    checkTrue(is.null(read.xls(wfile)))
    checkTrue(is.null(read.xls(wfile, type = "double")))

    x <- as.data.frame(double())
    colnames(x) <- "custCol" # ncol(x) is 1: colnames will potentially be written
    
    write.xls(x, wfile, colNames = FALSE)
    checkTrue(is.null(read.xls(wfile)))
    checkTrue(is.null(read.xls(wfile, type = "double")))

    write.xls(x, wfile, colNames = TRUE)
    wdata <- read.xls(wfile, colNames = FALSE, stringsAsFactor = FALSE)
    checkIdentical(wdata[[1]], "custCol")
    wdata <- read.xls(wfile, colNames = FALSE, type = "character")
    checkIdentical(wdata[[1]], "custCol")

    write.xls(x, wfile, colNames = "myColName")
    wdata <- read.xls(wfile, colNames = FALSE, stringsAsFactor = FALSE)
    checkIdentical(wdata[[1]], "myColName")
    wdata <- read.xls(wfile, colNames = FALSE, type = "character")
    checkIdentical(wdata[[1]], "myColName")
}


### colNames incl. rowNames

test.colNames.withRowNames <- function() {
        # colnames may optionally contain an entry for the rowname column
    mycol.1 <- c("Kol2", "Kol3", "Kol4")
    mycol.2 <- c("Kol1", "Kol2", "Kol3", "Kol4")

      # double
    rdata <- read.xls(rfile, colNames = TRUE, rowNames = TRUE, type = "double")
    checkIdentical(colnames(rdata), mycol.1)

    write.xls(rdata, wfile, colNames = TRUE, rowNames = TRUE)
    wdata <- read.xls(wfile, type = "double")
    checkIdentical(wdata, rdata)

    write.xls(rdata, wfile, colNames = mycol.1, rowNames = TRUE)
    wdata <- read.xls(wfile, type = "double")
    checkIdentical(wdata, rdata)

    write.xls(rdata, wfile, colNames = mycol.2, rowNames = TRUE)
    wdata <- read.xls(wfile, type = "double", rowNames = TRUE)
    checkIdentical(wdata, rdata)

        # data.frame
    rdata <- read.xls(rfile, colNames = TRUE, rowNames = TRUE, type = "data.frame")
    checkIdentical(colnames(rdata), mycol.1)

    write.xls(rdata, wfile, colNames = TRUE, rowNames = TRUE)
    wdata <- read.xls(wfile, type = "data.frame")
    checkIdentical(wdata, rdata)

    write.xls(rdata, wfile, colNames = mycol.1, rowNames = TRUE)
    wdata <- read.xls(wfile, type = "data.frame")
    checkIdentical(wdata, rdata)

    write.xls(rdata, wfile, colNames = mycol.2, rowNames = TRUE)
    wdata <- read.xls(wfile, type = "data.frame", rowNames = TRUE)
    checkIdentical(wdata, rdata)
}

test.colNames.withAutoRowNames <- function() {
    mycol <- c("Fertility", "Agriculture", "Testlogical", "Education", "Catholic", "Infant.Mortality", "Testcharacter")
        # double
    rdata <- read.xls(rfile, sheet = "dfSht", colNames = TRUE, type = "double", from = 5)
    checkIdentical(colnames(rdata), mycol)
    write.xls(rdata, wfile)
    wdata <- read.xls(wfile, type = "double")
    checkIdentical(wdata, rdata)
        # data.frame
        # note: free recognizes as logical, pro as logical->then->integer
    rdata <- read.xls(rfile, sheet = 5, type = "data.frame", from = 5)
    checkIdentical(colnames(rdata), mycol)
    write.xls(rdata, wfile)
    wdata <- read.xls(wfile, type = "data.frame")
    if (isFreeVersion) wdata[,3] <- as.logical(wdata[,3])
    checkIdentical(wdata, rdata)
}
