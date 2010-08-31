\name{xls.datetime}
\alias{dateTimeToStr}
\alias{dateTimeToIsoStr}
\alias{strToDateTime}
\alias{isoStrToDateTime}
\title{DateTime conversion}

\description{
    Several functions to convert Excel datetime values to and from strings.
}

\usage{
dateTimeToStr(odate, format = "")
strToDateTime(sdate)
dateTimeToIsoStr(odate, isoformat = "YYYY-MM-DD hh:mm:ss")
isoStrToDateTime(sdate)
}

\arguments{
    \item{odate}{a numeric (double) datetime value from Excel.}
    \item{format}{formatting string, see list in details. With an empty string the
        system defaults settings for shortDateFormat and longTimeFormat will be used.}
    \item{isoformat}{one of the following character strings: \code{YYYYMMDD} (basic date),
        \code{YYYY-MM-DD} (extended date), \code{YYYYMMDDhhmmss} (basic date/time),
        \code{YYYY-MM-DD hh:mm:ss} (extended date/time) or \code{YYYY-MM-DD hh:mm:ss.f}
        (extended date/time including fractions (with 1, 2 or 3 decimal places)}
    \item{sdate}{a date as a string.}
}

\details{
    \code{dateTimeToStr} converts a given double value to a string representation
    optionally using the formatting string.

    \code{strToDateTime} converts a given character string to a double value.
    The string must contain a valid date and/or time value (in respect to the
    active locale). Names (e.g. days and months) are not supported.

    \code{isoStrToDateTime} and \code{dateTimeToIsoStr} follow the ISO-8601 standard.

    The following table lists the supported formatting strings (adapted from Delphi help):

  \tabular{ll}{
    c   \tab ShortDateFormat followed by LongTimeFormat. Time is not displayed if midnight precisely.\cr
    d   \tab day as a number without a leading zero (1-31).\cr
    dd  \tab day as a number with a leading zero (01-31).\cr
    ddd \tab day as an abbreviation (Sun-Sat) using the ShortDayNames global variable.\cr
    dddd    \tab day as a full name (Sunday-Saturday) using LongDayNames.\cr
    ddddd   \tab date using ShortDateFormat variable.\cr
    dddddd  \tab date using LongDateFormat.\cr
    e   \tab year in the current period/era as a number without a leading zero.\cr
    ee  \tab year in the current period/era as a number with a leading zero.\cr
    g   \tab period/era as an abbreviation.\cr
    gg  \tab period/era as a full name.\cr
    m   \tab month as a number without a leading zero (1-12). If m immediately follows an h or hh, the minute rather than the month is displayed.\cr
    mm  \tab month as a number with a leading zero (01-12).\cr
    mmm \tab month as an abbreviation (Jan-Dec) using the ShortMonthNames global variable.\cr
    mmmm    \tab month as a full name (January-December) using LongMonthNames.\cr
    yy  \tab year as a two-digit number (00-99).\cr
    yyyy    \tab year as a four-digit number (0000-9999).\cr
    h   \tab hour without a leading zero (0-23).\cr
    hh  \tab hour with a leading zero (00-23).\cr
    n   \tab minute without a leading zero (0-59).\cr
    nn  \tab minute with a leading zero (00-59).\cr
    s   \tab second without a leading zero (0-59).\cr
    ss  \tab second with a leading zero (00-59).\cr
    z   \tab millisecond without a leading zero (0-999).\cr
    zzz \tab millisecond with a leading zero (000-999).\cr
    t   \tab time using the format given by ShortTimeFormat.\cr
    tt  \tab time using the format given by LongTimeFormat.\cr
    am/pm   \tab uses the 12-hour clock for the preceding h or hh specifier, and displays 'am' for any hour before noon, and 'pm' for any hour after noon. The am/pm specifier can use lower, upper, or mixed case, and the result is displayed accordingly.\cr
    a/p \tab uses the 12-hour clock for the preceding h or hh specifier, and displays 'a' for any hour before noon, and 'p' for any hour after noon. The a/p specifier can use lower, upper, or mixed case, and the result is displayed accordingly.\cr
    ampm    \tab uses the 12-hour clock for the preceding h or hh specifier, and displays the contents of the TimeAMString global variable for any hour before noon, and the contents of the TimePMString global variable for any hour after noon.\cr
    /   \tab separator character given by the DateSeparator variable.\cr
    :   \tab time separator character given by TimeSeparator.\cr
    'xx'/"xx"\tab characters enclosed in single or double quotes are displayed as-is, and do not affect formatting.
  }
}

\section{Pro version}{
    See section in \code{\link{read.xls}}.
}

\seealso{
    \code{\link{read.xls}}, \code{\link{write.xls}}
}

\examples{
# --- only run for regular, i.e. non-cran version
shlib <- system.file("libs", if (nzchar(arch <- .Platform$r_arch)) arch else "",
                     paste("xlsReadWrite", .Platform$dynlib.ext, sep = ""), package = "xlsReadWrite")
if (file.exists(shlib) && (file.info(shlib)$size > 20000)) {
# -----------------------------------------------

    # convert iso datetime character to numeric datetime
(idt <- isoStrToDateTime("2010-08-14 09:23:13"))

    # convert numeric datetime to string
dateTimeToStr(idt)
(sdt <- dateTimeToStr(idt, format = "c"))      # same as above (default)

dateTimeToStr(idt, format = "dddddd")          # long date format
(sd <- dateTimeToStr(idt, format = "ddddd"))   # short date format
(st <- dateTimeToStr(idt, format = "t"))       # short time format

    # convert character datetime to numeric
(dt <- strToDateTime(sdt))
(dd <- strToDateTime(sd))
(tt <- strToDateTime(st))

stopifnot(isTRUE(all.equal(dt, (dd + tt))))
}}

\keyword{chron}
