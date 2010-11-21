dateTimeToStr <- function(odate, format = "") {
    .Call("DateTimeXls", "xheDateTimeToStr", odate, format)
}

strToDateTime <- function(sdate) {
    .Call("DateTimeXls", "xheStrToDateTime", sdate, "")
}

dateTimeToIsoStr <- function(odate, isoformat = "YYYY-MM-DD hh:mm:ss") {
    isofmt <- as.integer(switch(isoformat,
                                "YYYYMMDD" = 1,
                                "YYYY-MM-DD" = 2,
                                "YYYYMMDDhhmmss" = 3,
                                "YYYY-MM-DD hh:mm:ss" = 4,
                                "YYYY-MM-DD hh:mm:ss.f" = 5,
                                stop("wrong ISO-8601 format. Allowed are:\n'YYYYMMDD', 'YYYY-MM-DD', 'YYYYMMDDhhmmss', 'YYYY-MM-DD hh:mm:ss' and 'YYYY-MM-DD hh:mm:ss.f'")))
    .Call("DateTimeXls", "xheDateTimeToIsoStr", odate, isofmt)
}

isoStrToDateTime <- function(sdate) {
    .Call("DateTimeXls", "xheIsoStrToDateTime", sdate, "")
}
