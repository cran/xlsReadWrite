\name{read.xls}
\alias{read.xls}
\title{Data Input}

\description{
  Natively reads an Excelsheet and creates a data.frame or a matrix from it.
  The supported format is BIFF8, i.e. Excel v97-2003.
}

\usage{
  read.xls( file, colNames = TRUE, sheet = 1, type = "data.frame",
          from = 1, colClasses = NA )
}

\arguments{
  \item{file}{the name of the local file to read from.}
  \item{colNames}{indicates if there is a header row in Excel. Either a logical
    value or a character vector with user defined column names. With \code{TRUE}
    the argument \code{from} denotes the header row otherwise the first data row.}
  \item{sheet}{to read from. Either a number (starting from 1) or a character
    string naming the sheet.}
  \item{type}{specifies the type of the return value. Either \code{data.frame}
    or the entries \code{double}, \code{integer}, \code{logical} or
    \code{character} which all give back a matrix.}
  \item{from}{specifies the first row to start reading. Depending on the value
    of \code{colNames} this is the header row or the first data row. (Slightly
    different from the pro version where the upper left corner (row/col) can be
    indicated).}
  \item{colClasses}{A scalar or a vector of classes to be assumed for the
    columns. Possible values are \code{double}, \code{integer}, \code{logical},
    \code{character}and \code{factor}. With \code{NA} the default conversion
    will be used. \code{rowname} can be written for the first column to use
    it for the rownames (note: this is only possible for the first column).}
}

\details{
  For data.frames the first column is taken for the rownames either if the first
  entry in \code{colClasses} is \code{rowname} or if all of the following
  conditions are met: - \code{colClasses} is NULL, - colNames is TRUE
  or a vector where the first entry is empty and - the first entry of the
  header row in the Excel table is empty.
}

\value{
  A data.frama or a matrix containing the data.
}

\section{Pro version}{
  There is/will be a Pro version which is formally supported and has some more
  functionality, e.g. sheet handling (insert, copyAndInsert, select, delete and
  rename), select Excel regions (columns, rows, cells) directly, append data,
  work with multiple files at the same time, information about an Excelfile and
  last but not least an oledatetime class and lots of datetime functions (BETA).

  Why not check it out? xlsReadWritePro is free to try and most probably will
  serve you well (it surely does so here for our work!). Further infos, updates,
  bugfixes and support, please see: http://treetron.googlepages.com.
}

\author{Hans-Peter Suter}

\seealso{
  \code{\link{write.xls}},
  \code{\link{read.table}}, \code{\link{matrix}}, \code{\link{data.frame}}
 }

\keyword{file}
