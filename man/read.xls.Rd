\name{read.xls}
\alias{read.xls}
\title{Data Input}

\description{
  Natively reads an Excelsheet and creates a data.frame or a matrix from it.
  The supported format is BIFF8, i.e. Excel v97-2003.
}

\usage{
read.xls( file, colNames = TRUE, sheet = 1, type = "data.frame", from = 1 )
}

\arguments{
  \item{file}{the name of the file to read from.}
  \item{colNames}{indicates if there is a header row in Excel. Either a logical
    value or a character vector with user defined column names.}
  \item{sheet}{to read from. Either a number (starting from 1) or a character
    string naming the sheet.}
  \item{type}{specifies the type of the return value. Either \code{data.frame}
    or the entries \code{double}, \code{integer}, \code{logical} or
    \code{character} which give back a matrix.}
  \item{from}{specifies the first row to start reading. Depending on the value
    of \code{colNames} this is the header row or the first data row. (Slightly
    different from the pro version where the upper left corner (row/col) can be
    indicated).}
}

\details{
  For data.frames there is an auto-row feature. Provided the following two
  conditions are met, a column is taken as the rownames: 1) colNames is TRUE or
  has values where the first entry is empty and 2) in the Excel table the
  first entry of the header is empty.
}

\value{
  A data.frama or a matrix containing the data.
}

\section{Pro version}{
  Besides the free open source version there is a shareware version which
  is formally supported and offers some more features, e.g.:\code{xls.sheet}:
  insert, delete, rename and select sheets, \bold{directly address} sectors (columns,
  rows or \bold{cells}) in Excel, \code{xls.open} and \code{xls.close} keep the
  connection in memory and avoid the overhead to re-read the whole file for
  subsequent actions. \code{colClasses} supports classes and rownames.
  
  Why not check out xlsReadWritePro? It's free to try and - besides
  donations (hint, hint...) - also a good way to support my work.
  Infos, updates, bugfixes etc, see: http://treetron.googlepages.com
}

\author{Hans-Peter Suter}

\seealso{
  \code{\link{write.xls}},
  \code{\link{read.table}}, \code{\link{matrix}}, \code{\link{data.frame}}
 }

\keyword{file}
