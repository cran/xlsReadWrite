### vector

test.readWrite.vector <- function() {
    x <- c(42, 43)
    write.xls(x, wfile, colNames = FALSE)
    wdata <- read.xls(wfile, colNames = FALSE, type = "double")
    checkEquals(as.integer(wdata), x)
}


### matrices

test.readWrite.double <- function() {
    myval <- cbind(c(42.345, 2.1318527, 5.1960286, 4.0520327, 5.4554428, 0.9201211, 4.3205375, 9.2289868, 3.4951773, 7.162185, 0.0797354, 0.1904384, 6.9478864),
                   c(3.3950068, 3.688808, 1.1828297, 5.2862727, 2.7721154, 4.5476344, 9.7239585, 4.8111606, 6.2504654, 9.5798654, 6.6359867, 2.8083093, 6.8696548),
                   c(3.8005233, 3.1470417, 7.4015953, 4.7576348, 7.5309305, 2.1087203, 8.5166949, 2.070765, 9.2605556, 8.9307317, 6.0705168, 8.4730313, 4.9871624),
                   c(-20.4994424,	-22.6819437, -18.9855194, -27.6520526, -12.619529, -3.0658813, -15.4015045, 0.6767891, -10.1689136, -18.4513764, -32.4987449, -35.3584594, 4.4270298))
    rdata <- read.xls(rfile, sheet = "dSht", type = "double")
    checkEquals(rdata, myval, check.attributes = FALSE)
    checkIdentical(colnames(rdata), c("Kol1", "Kol2", "Kol3", "Kol4"))
    checkIdentical(rownames(rdata), as.character(1:13))

    write.xls(rdata, wfile)
    wdata <- read.xls(wfile, type = "double")
    checkIdentical(wdata, rdata)
}

test.readWrite.numeric <- function() {
    myval <- cbind(c(42.345, 2.1318527, 5.1960286, 4.0520327, 5.4554428, 0.9201211, 4.3205375, 9.2289868, 3.4951773, 7.162185, 0.0797354, 0.1904384, 6.9478864),
                   c(3.3950068, 3.688808, 1.1828297, 5.2862727, 2.7721154, 4.5476344, 9.7239585, 4.8111606, 6.2504654, 9.5798654, 6.6359867, 2.8083093, 6.8696548),
                   c(3.8005233, 3.1470417, 7.4015953, 4.7576348, 7.5309305, 2.1087203, 8.5166949, 2.070765, 9.2605556, 8.9307317, 6.0705168, 8.4730313, 4.9871624),
                   c(-20.4994424,	-22.6819437, -18.9855194, -27.6520526, -12.619529, -3.0658813, -15.4015045, 0.6767891, -10.1689136, -18.4513764, -32.4987449, -35.3584594, 4.4270298))
    rdata <- read.xls(rfile, sheet = "dSht", type = "numeric")
    checkEquals(rdata, myval, check.attributes = FALSE)
    checkIdentical(colnames(rdata), c("Kol1", "Kol2", "Kol3", "Kol4"))
    checkIdentical(rownames(rdata), as.character(1:13))

    write.xls(rdata, wfile)
    wdata <- read.xls(wfile, type = "numeric")
    checkIdentical(wdata, rdata)
}

test.readWrite.integer <- function() {
    myidx <- c(1, 5, 19)
    myval <- rbind(c(1L, 2L, 3L, NA, NA, NA, NA, NA, NA, NA, NA, NA),
                   c(1L, 2L, 3L, 4L, 5L, 6L, 7L, 8L, 9L, 10L, 11L, 12L),
                   c(0L, 22L, 33L, NA, NA, NA, NA, NA, NA, NA, NA, NA))
    rdata <- read.xls(rfile, sheet = "intSht", type = "integer")
    checkEquals(rdata[myidx,], myval, check.attributes = FALSE)
    checkIdentical(colnames(rdata), c("X1", "X2", "X3", paste("V", 4:12, sep = "")))
    checkIdentical(rownames(rdata), as.character(1:19))

    write.xls(rdata, wfile)
    wdata <- read.xls(wfile, type = "integer")
    checkIdentical(wdata, rdata)
}

test.readWrite.logical <- function() {
    myval <- cbind(c(T, T, T, T, F, F, T, T, T, F, F), c(F, F, NA, NA, rep(F, 7)))
    rdata <- read.xls(rfile, colNames = FALSE, sheet = "logSht", type = "logical", from = 4)
    checkEquals(rdata, myval, check.attributes = FALSE)
    checkIdentical(colnames(rdata), c("V1", "V2"))
    checkIdentical(rownames(rdata), as.character(1:11))

    write.xls(rdata, wfile, colNames = FALSE)
    wdata <- read.xls(wfile, colNames = FALSE, type = "logical")
    checkIdentical(wdata, rdata)
}

test.readWrite.character <- function() {
    myval <- matrix(c(
        "Sind hierorts H\u00e4user gr\u00fcn, tret ich noch in ein Haus.", "Sind hier die Br\u00fccken heil, geh ich auf gutem Grund.", 
        "I'd agree with that,' said Arkady.", "The world, if it has a future, has an ascetic future."), ncol = 2)
    rdata <- read.xls(rfile, colNames = TRUE, "charSht", "character")
    checkEquals(rdata, myval, check.attributes = FALSE)
    checkIdentical(colnames(rdata), c("Bachmann", "Chatwin"))
    checkIdentical(rownames(rdata), c("1", "2"))

    write.xls(rdata, wfile, colNames = TRUE)
    wdata <- read.xls(wfile, TRUE, 1, "character")
    checkIdentical(wdata, rdata)
}


### data.frame

test.readWrite.dataFrame.1 <- function() {
    mylogical <- if (isFreeVersion) "logical" else "numeric"
    myval <- data.frame(Fertility = c(80.2, 83.1, 92.5, 85.8, 76.9, 76.1, 83.8, 92.4, 82.4, 82.9, 87.1, 64.1), 
        Agriculture = c(17, 45.1, 39.7, 36.5, 43.5, 35.3, 70.2, 67.8, 53.3, 45.2, 64.5, 62), 
        Testlogical = if (isFreeVersion) c(T, T, F, T, T, F, T, T, F, T, T, F) else 
                                  as.integer(c(NA, NA, NA, 1, 1, 0, rep(NA, 6))),
        Education = as.numeric(c(12, 9, 5, 7, 15, 7, 7, 8, 7, 13, 6, 12)), 
        Catholic = c(9.96, 84.84, 93.4, 33.77, 5.16, 90.57, 92.85, 97.16, 97.67, 91.38, 98.61, 8.52), 
        Infant.Mortality = c(22.2, 22.2, 20.2, 20.3, 20.6, 26.6, 23.6, 24.9, 21, 24.4, 24.5, 16.5), 
        Testcharacter = c("Co", "De", "Fr", "Mo", "Ne", "Po", "Br", "Gl", "Gr", "Sa", "Ve", "Ai"), stringsAsFactors = TRUE)
    mycls <- c("numeric", "numeric", mylogical, "numeric", "numeric", "numeric", "factor")

    rdata <- read.xls(rfile, colNames = TRUE, "dfSht", from = 5)
    checkEquals(rdata, myval, check.attributes = FALSE)
    checkIdentical(colnames(rdata), colnames(myval))
    checkIdentical(rownames(rdata), c("Courtelary", "Delemont", "Franches-Mnt", "Moutier", "Neuveville", "Porrentruy", "Broye", "Glane", "Gruyere", "Sarine", "Veveyse", "Aigle"))
    checkIdentical(as.vector(sapply(rdata, class)), mycls)

    write.xls(rdata, wfile)
    wdata <- read.xls(wfile)
    if (isFreeVersion) wdata[,3] <- as.logical(wdata[,3])
    checkIdentical(wdata, rdata)
}

test.readWrite.dataFrame.2 <- function() {
    myclsp <- list(IntDate = "numeric", AsDate = "character", AsIsoDate = c("isodate", "isodatetime"),
                  Hour = "numeric", Minute = "numeric", Sec = "numeric", IntTime = "numeric", AsTime = "character",
                  AsIsoTime = c("isotime", "isodatetime"), IntDateTime = "numeric",
                  AsDateTime = "isodatetime", AsIsoDateTime = "isodatetime")
    myclsf <- structure(c("numeric", "character", "character", "numeric", "numeric", "numeric", "numeric",
                          "character", "character", "numeric", "character", "character"), 
                          .Names = c("IntDate", "AsDate", "AsIsoDate", "Hour", "Minute", "Sec", "IntTime", 
                          "AsTime", "AsIsoTime", "IntDateTime", "AsDateTime", "AsIsoDateTime"))
    mycls <- if (isFreeVersion) myclsf else myclsp

    rdata <- read.xls(rfile, sheet = "dateTime", dateTime = "isodatetime", stringsAsFactors = FALSE)
    checkIdentical(rdata[3,3], "2005-03-22")
    checkIdentical(rdata[12,7], 0.42353010416666664)
    checkIdentical(dateTimeToStr(rdata[12,10], "hh:mm:ss"), "10:09:53")
    checkIdentical(rdata[18,10], 38457.45491900463)
    checkIdentical(colnames(rdata), c("IntDate", "AsDate", "AsIsoDate", "Hour", "Minute", "Sec", "IntTime", "AsTime", "AsIsoTime", "IntDateTime", "AsDateTime", "AsIsoDateTime"))
    checkIdentical(rownames(rdata), as.character(1:19))
    checkIdentical(sapply(rdata, class), mycls)

    write.xls(rdata, wfile)
    wdata <- read.xls(wfile, dateTime = "isodatetime", stringsAsFactors = FALSE)
    checkIdentical(wdata, rdata)
}
