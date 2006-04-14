\name{xls.datetime}

\alias{dateTimeToStr}
\alias{dateTimeToIsoStr}
\alias{strToDateTime}
\alias{isoStrToDateTime}

\title{DateTime conversion}

\description{
  Misc. functions to convert Excel datetime values to and from strings.
}

\usage{
dateTimeToStr( odate, format = "" )
strToDateTime( sdate )
dateTimeToIsoStr( odate, isoformat = "YYYY-MM-DD hh:mm:ss" )
isoStrToDateTime( sdate )
}

\arguments{
  \item{odate}{a numeric (double) datetime value from Excel.}
  \item{format}{formatting string. See list in details.}
  \item{isoformat}{one of the following character strings: \code{YYYYMMDD} (basic date),
    \code{YYYY-MM-DD} (extended date), \code{YYYYMMDDhhmmss} (basic date/time),
    \code{YYYY-MM-DD hh:mm:ss} (extended date/time) or \code{YYYY-MM-DD hh:mm:ss.f}
    (extended date/time including fractions (with 1, 2 or 3 decimal places)}
  \item{sdate}{a date as a string.}
}

\details{
  \code{dateTimeToStr} converts a given double value to a string representation. 
    Optionally a formatting string can be used. By default the shortDateFormat 
    and longTimeFormat settings of the system will be used.

  \code{strToDateTime} converts a given character string to a double value. 
    The string must contain a valid date and/or time value (in respect of the 
    current locale).

  \code{dateTimeToIsoStr} and \code{isoStrToDateTime} do the same, but work with 
  date strings that are formatted according to the isoformat.

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
  While the free version should be sufficient for most day-to-day duties, you have 
  the opportunity to choose a more powerful pro version and at the same time support 
  our development and maintenance effort. Additional features of xlsReadWritePro: \cr
  - work with data area subsets, - use (named) ranges and pick individual cells, 
  - work with sheets (copy, rename, delete), - read/write formula values, 
  - insert images, - use oleDateTime classes and many datetime helper functions, 
  - append data to existing files and - work with in-memory objects (e.g. to compose Excel reports). 
  
  The trial (http://treetron.googlepages.com) is free and the two versions can 
  easily be switched and coexist just fine. So, why not check it out for yourself? 
}

\author{Hans-Peter Suter}

\seealso{
  \code{\link{read.xls}}, \code{\link{write.xls}}
}

\keyword{chron}
