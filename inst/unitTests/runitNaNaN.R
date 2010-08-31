### setup

idx <- cbind(c(2,4),c(1,2))

naidx <- matrix(TRUE, nrow = 19, ncol = 12)
naidx[,1:3] <- FALSE; naidx[5,] <- FALSE; naidx[9,5] <- FALSE

mymat <- function(mode, nanOrNa, with42 = FALSE) {
    if (with42) {
        nanOrNa <- NA
        res <- do.call(paste("as", mode, sep = "."), list(rep(42, 19*12)))
        dim(res) <- c(19,12)
        res[naidx] <- nanOrNa
    } else {
        res <- cbind(a = do.call(paste("as", mode, sep = "."), list(1:4)),
                     b = do.call(paste("as", mode, sep = "."), list(12:15)))
        res[idx] <- nanOrNa
    }
    stopifnot(do.call(paste("is", mode, sep = "."), list(res)))  # afaik only numeric/character works with NaN
    res
}


### NA

test.NaN_NA.readWithNaStrings <- function() {
    rdata <- read.xls(rfile, naStrings = c("NO", "FALSE", "NEIN"), type = "double", sheet = "logSht", from = 3)
    checkTrue(all(is.na(rdata[c(6,8,11),2])))
    rdata <- read.xls(rfile, naStrings = c("NO", "FALSE", "NEIN"), type = "integer", sheet = "logSht", from = 3)
    checkTrue(all(is.na(rdata[c(6,8,11),2])))
    rdata <- read.xls(rfile, naStrings = c("NO", "FALSE", "NEIN"), type = "logical", sheet = "logSht", from = 3)
    checkTrue(all(is.na(rdata[c(6,8,11),2])))
    rdata <- read.xls(rfile, naStrings = c("NO", "FALSE", "NEIN"), type = "character", sheet = "logSht", from = 3)
    checkTrue(all(is.na(rdata[c(6,8,11),2])))
    rdata <- read.xls(rfile, naStrings = c("NO", "FALSE", "NEIN"), type = "data.frame", sheet = "logSht", from = 3)
    checkTrue(all(is.na(rdata[c(6,8,11),2])))
}

test.NaN_NA.writeWithNaStrings <- function() {
    write.xls(mymat("double", NA), wfile, naStrings = "hello", colNames = FALSE)
    checkIdentical(read.xls(wfile, type = "character", colNames = FALSE)[idx], c("hello", "hello"))
    write.xls(mymat("integer", NA), wfile, naStrings = "hello", colNames = FALSE)
    checkIdentical(read.xls(wfile, type = "character", colNames = FALSE)[idx], c("hello", "hello"))
    write.xls(mymat("logical", NA), wfile, naStrings = "hello", colNames = FALSE)
    checkIdentical(read.xls(wfile, type = "character", colNames = FALSE)[idx], c("hello", "hello"))
    write.xls(mymat("character", NA), wfile, naStrings = "hello", colNames = FALSE)
    checkIdentical(read.xls(wfile, type = "character", colNames = FALSE)[idx], c("hello", "hello"))
    write.xls(data.frame(mymat("double", NA)), wfile, naStrings = "hello", colNames = FALSE)
    checkIdentical(read.xls(wfile, type = "character", colNames = FALSE)[idx], c("hello", "hello"))
}    

test.NaN_NA.readNaDefault <- function() {
    checkTrue(all(is.na(read.xls(rfile, sheet = "intSht", type = "double")[naidx])))
    checkTrue(all(is.na(read.xls(rfile, sheet = "intSht", type = "integer")[naidx])))
    checkTrue(all(is.na(read.xls(rfile, sheet = "intSht", type = "logical")[naidx])))
    checkTrue(all(read.xls(rfile, sheet = "intSht", type = "character")[naidx] == ""))
    checkTrue(all(is.na(read.xls(rfile, sheet = "intSht", type = "data.frame")[naidx])))
}    
    
test.NaN_NA.writeNaDefault <- function() {
    write.xls(mymat("double", with42 = TRUE), wfile)
    checkTrue(all(read.xls(wfile, type = "character")[naidx] == ""))
    checkTrue(all(read.xls(wfile, type = "character")[!naidx] == "42"))
    write.xls(mymat("integer", with42 = TRUE), wfile)
    checkTrue(all(read.xls(wfile, type = "character")[naidx] == ""))
    checkTrue(all(read.xls(wfile, type = "character")[!naidx] == "42"))
    write.xls(mymat("logical", with42 = TRUE), wfile)
    checkTrue(all(read.xls(wfile, type = "character")[naidx] == ""))
    checkTrue(all(read.xls(wfile, type = "character")[!naidx] == "1"))
    write.xls(mymat("character", with42 = TRUE), wfile)
    checkTrue(all(read.xls(wfile, type = "character")[naidx] == ""))
    checkTrue(all(read.xls(wfile, type = "character")[!naidx] == "42"))
    write.xls(data.frame(mymat("double", with42 = TRUE)), wfile)
    checkTrue(all(read.xls(wfile, type = "character")[naidx] == ""))
    checkTrue(all(read.xls(wfile, type = "character")[!naidx] == "42"))
}


### NaN

test.NaN_NA.doubleNaN <- function() {
    write.xls(mymat("double", NaN), wfile, colNames = FALSE)
    wdata <- read.xls(wfile, type = "double", colNames = FALSE)
    checkIdentical(wdata[idx], c(NaN, NaN))
}

test.NaN_NA.characterNaN <- function() {
    write.xls(mymat("character", NaN), wfile, colNames = FALSE)
    wdata <- read.xls(wfile, type = "character", colNames = FALSE)
    checkIdentical(wdata[idx], c("NaN", "NaN"))
}

test.NaN_NA.frameNaN <- function() {
    write.xls(data.frame(mymat("double", NaN)), wfile, colNames = FALSE)
    wdata <- read.xls(wfile, colNames = FALSE)
    checkIdentical(wdata[idx], c(NaN, NaN))
}
