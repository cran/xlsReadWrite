### setup

strictNA <- function(x) is.na(x) && !is.nan(x)

naidx <- rbind(c(2,1))
nanidx <- rbind(c(4,2))

dataidx <- matrix(TRUE, nrow = 19, ncol = 12)
chidx <- !dataidx
dataidx[,1:3] <- FALSE; dataidx[5,] <- FALSE; dataidx[9,5:7] <- FALSE;
chidx[,5:6] <- TRUE; chidx[5,5:6] <- FALSE; chidx[9,5:6] <- FALSE;

databingo <- rbind(c(9, 5))
datana <- rbind(c(9, 6))
datanan <- rbind(c(9, 7))

mat <- function(mode) {
    res <- cbind(a = do.call(paste("as", mode, sep = "."), list(1:4)),
                 b = do.call(paste("as", mode, sep = "."), list(12:15)))
    res[naidx] <- NA
    res[nanidx] <- NaN
    res
}


### read write with custom naStrings

test.NaN_NA.readWithNaStrings <- function() {
    rdata <- read.xls(rfile, naStrings = c("NO", "FALSE", "NEIN"), type = "double", sheet = "logSht", from = 3)
    checkTrue(all(strictNA(rdata[c(6,8,11),2])))
    rdata <- read.xls(rfile, naStrings = c("NO", "FALSE", "NEIN"), type = "integer", sheet = "logSht", from = 3)
    checkTrue(all(strictNA(rdata[c(6,8,11),2])))
    rdata <- read.xls(rfile, naStrings = c("NO", "FALSE", "NEIN"), type = "logical", sheet = "logSht", from = 3)
    checkTrue(all(strictNA(rdata[c(6,8,11),2])))
    rdata <- read.xls(rfile, naStrings = c("NO", "FALSE", "NEIN"), type = "character", sheet = "logSht", from = 3)
    checkTrue(all(strictNA(rdata[c(6,8,11),2])))

    rdata <- read.xls(rfile, naStrings = c("NO", "FALSE", "NEIN"), type = "data.frame", sheet = "logSht", from = 3)
    checkTrue(all(strictNA(rdata[c(6,8,11),2])))
}

test.NaN_NA.writeWithNaStrings <- function() {
    write.xls(mat("double"), wfile, naStrings = "hello", colNames = FALSE)
    checkIdentical(read.xls(wfile, type = "character", colNames = FALSE)[naidx], "hello")
    write.xls(mat("integer"), wfile, naStrings = "hello", colNames = FALSE)
    checkIdentical(read.xls(wfile, type = "character", colNames = FALSE)[naidx], "hello")
    write.xls(mat("logical"), wfile, naStrings = "hello", colNames = FALSE)
    checkIdentical(read.xls(wfile, type = "character", colNames = FALSE)[naidx], "hello")
    write.xls(mat("character"), wfile, naStrings = "hello", colNames = FALSE)
    checkIdentical(read.xls(wfile, type = "character", colNames = FALSE)[naidx], "hello")

    write.xls(data.frame(mat("double")), wfile, naStrings = "hello", colNames = FALSE)
    checkIdentical(read.xls(wfile, type = "character", colNames = FALSE)[naidx], "hello")
}    

test.NaN_NA.writeWithMultipleNaStrings <- function() {
    checkException(write.xls(mat("double"), wfile, naStrings = c("hello", "world"), colNames = FALSE), silent = TRUE)
}


### read write (including cells picking)

test.NaN_NA.double <- function() {
    write.xls(mat("double"), wfile, colNames = FALSE)
    wdata <- read.xls(wfile, type = "double", colNames = FALSE)
    checkIdentical(wdata[naidx], as.double(NA))
    checkIdentical(wdata[nanidx], NaN)
    if (!isFreeVersion) {
        wdata <- read.xls(wfile, type = "double", cells = c(naidx, nanidx))
        checkIdentical(wdata[1], as.double(NA))
        checkIdentical(wdata[2], NaN)
    }
}

test.NaN_NA.integer <- function() {
    write.xls(mat("integer"), wfile, colNames = FALSE)
    wdata <- read.xls(wfile, type = "integer", colNames = FALSE)
    checkIdentical(wdata[naidx], as.integer(NA))
    checkIdentical(wdata[nanidx], as.integer(NA))
    if (!isFreeVersion) {
        wdata <- read.xls(wfile, type = "integer", cells = c(naidx, nanidx))
        checkIdentical(wdata[1], as.integer(NA))
        checkIdentical(wdata[2], as.integer(NA))
    }
}

test.NaN_NA.logical <- function() {
    write.xls(mat("logical"), wfile, colNames = FALSE)
    wdata <- read.xls(wfile, type = "logical", colNames = FALSE)
    checkIdentical(wdata[naidx], NA)
    checkIdentical(wdata[nanidx], NA)
    if (!isFreeVersion) {
        wdata <- read.xls(wfile, type = "logical", cells = c(naidx, nanidx))
        checkIdentical(wdata[1], NA)
        checkIdentical(wdata[2], NA)
    }
}

test.NaN_NA.character <- function() {
    write.xls(mat("character"), wfile, colNames = FALSE)
    wdata <- read.xls(wfile, type = "character", colNames = FALSE)
    checkIdentical(wdata[naidx], "")
    checkIdentical(wdata[nanidx], "NaN")
    if (!isFreeVersion) {
        wdata <- read.xls(wfile, type = "character", cells = c(naidx, nanidx))
        checkIdentical(wdata[1], "")
        checkIdentical(wdata[2], "NaN")
    }

    write.xls(mat("character"), wfile, colNames = FALSE, naStrings = "NA")
    wdata <- read.xls(wfile, type = "character", colNames = FALSE, naStrings = "NA")
    checkIdentical(wdata[naidx], as.character(NA))
    checkIdentical(wdata[nanidx], "NaN")
}

test.NaN_NA.frameNaN <- function() {
    write.xls(data.frame(mat("double")), wfile, colNames = FALSE)
    wdata <- read.xls(wfile, colNames = FALSE)
    checkIdentical(wdata[naidx], as.double(NA))
    checkIdentical(wdata[nanidx], NaN)
    if (!isFreeVersion) {
        suppressWarnings(wdata <- read.xls(wfile, cells = c(naidx, nanidx)))
        checkIdentical(wdata[[1]], NA)
        checkIdentical(wdata[[2]], NaN)
    }
}


### read data
# - we test NA, NaN, empty cells and wrong type
# - because of 'naString = NA', 'NA' values stays plain strings objects
# - 'NaN' become always NaN objects

test.NaN_NA.readMatrix <- function() {
    rdata <- read.xls(rfile, sheet = "intSht", type = "double")
    checkIdentical(rdata[dataidx], rep(as.numeric(NA), length(rdata[dataidx])))  # empty cells
    checkIdentical(rdata[databingo], as.double(NA))   # 'bingo' value
    checkIdentical(rdata[datana], as.double(NA))      # 'NA' value
    checkIdentical(rdata[datanan], NaN)               # 'NaN' value
    
    rdata <- read.xls(rfile, sheet = "intSht", type = "integer")
    checkIdentical(rdata[dataidx], rep(as.integer(NA), length(rdata[dataidx])))  # empty cells
    checkIdentical(rdata[databingo], as.integer(NA))  # 'bingo' value
    checkIdentical(rdata[datana], as.integer(NA))     # 'NA' value
    checkIdentical(rdata[datanan], as.integer(NA))    # 'NaN' value
    
    rdata <- read.xls(rfile, sheet = "intSht", type = "logical")
    checkIdentical(rdata[dataidx], rep(NA, length(rdata[dataidx])))  # empty cells
    checkIdentical(rdata[databingo], FALSE)           # 'bingo' value
    checkIdentical(rdata[datana], FALSE)              # 'NA' value
    checkIdentical(rdata[datanan], NA)                # 'NaN' value
    
    rdata <- read.xls(rfile, sheet = "intSht", type = "character")
    checkIdentical(rdata[dataidx], rep("", length(rdata[dataidx])))  # empty cells
    checkIdentical(rdata[databingo], "bingo")         # 'bingo' value
    checkIdentical(rdata[datana], "NA")               # 'NA' value
    checkIdentical(rdata[datanan], "NaN")             # 'NaN' value
}

test.NaN_NA.readFrame <- function() {
    tmpcls <- c("integer", "integer", "integer", "integer", "character", "character",
                "numeric", "integer", "integer", "integer", "integer", "integer")

    rdata <- read.xls(rfile, sheet = "intSht", type = "data.frame", stringsAsFactors = FALSE,
                      colClass = if (isFreeVersion) tmpcls else NA )

    checkIdentical(as.vector(sapply(rdata, class)), tmpcls)  

    checkIdentical(rdata[chidx], rep("", sum(chidx)))

    tmpidx <- !(chidx & dataidx) & dataidx            # empty cells
    checkIdentical(rdata[tmpidx], rep(as.character(NA), sum(tmpidx)))

    checkTrue(rdata[databingo] == "bingo")            # 'bingo' value
    checkTrue(rdata[datana] == "NA")                  # 'NA' value
    checkIdentical(rdata[datanan[1],datanan[2]], NaN) # NaN value (this is a double type column)
}
