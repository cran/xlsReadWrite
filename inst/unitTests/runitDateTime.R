### setup

myd <- c(38429, 38432, 38433, 38434, 38435, 38440, 38441, 38442, 38443)


### basic datetime functions

test.dateTime.conversion <- function() {
    s <- dateTimeToStr(myd)
    res <- strToDateTime(s)
    if (isFreeVersion) checkEquals(res, myd) else checkEquals(unclass(res), myd)
}

test.dateTime.conversionFmt <- function() {
    res <- dateTimeToStr(myd, 'mm-dd-yyyy')
    checkEquals(res[c(3, 7)], c("03-22-2005", "03-30-2005"))
}

test.dateTime.isoConversion <- function() {
    res <- dateTimeToIsoStr(myd)
    res <- isoStrToDateTime(res)
    if (isFreeVersion) checkEquals(res, myd) else checkEquals(unclass(res), myd)
}

test.dateTime.isoConversionFmt <- function() {
    res <- isoStrToDateTime(c("20070319", "2007-03-19", "20070319233112", "2007-03-19 23:31:12" )) 
    if (isFreeVersion) {
        checkEquals(res, c(39160, 39160, 39160.98, 39160.98))
    } else {
        checkEquals(unclass(res), c(39160, 39160, 39160.98, 39160.98))
    }
}
