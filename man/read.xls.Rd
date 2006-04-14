\name{read.xls}
\alias{read.xls}
\title{Read Excelfiles}

\description{
  Reads an Excelsheet and returns a data.frame or a matrix. 
  Supported is the native BIFF8 file format, i.e. Excel 97-2003.
}

\usage{
read.xls( file, 
          colNames = TRUE, 
          sheet = 1, 
          type = "data.frame",
          from = 1, 
          rowNames = NA, colClasses = NA, checkNames = TRUE, 
          dateTimeAs = "numeric", 
          stringsAsFactors = default.stringsAsFactors() )
}

\arguments{
  \item{file}{the name of a file. If it does not contain an absolute path, the 
    file name is relative to the current working directory.}
  \item{colNames}{a character vector or a logical. \code{TRUE} takes a title row
    from the spreadsheet. The default is "V" followed by the column number.}
  \item{sheet}{to read from. A case sensitive character string or a number.}
  \item{type}{the type of the returned value. Either \code{data.frame} or - 
    for matrices - \code{double}, \code{integer}, \code{logical} or \code{character}.}
  \item{from}{the starting row to read from.}
  \item{rowNames}{a character vector with the actual row names or \code{TRUE} to 
    use the first Excel column. 
    
    With \code{NA} the first column will be considered to contain row names 
    under two conditions: 1) \code{colNames} is \code{TRUE} or contains a character 
    vector and 2) the first column name thereof is an empty string.}
  \item{colClasses}{the class of the data.frame columns. Possible entries are: 
    \code{double}, \code{numeric}, \code{integer}, \code{logical}, \code{character}, 
    \code{factor}, \code{NA},  \code{isodate}, \code{isotime} and \code{isodatetime}. 
    A scalar entry will be recycled if necessary.

    \code{isodate} retrieves a "YYYY-MM-DD", \code{isotime} a "hh:mm:ss" and 
    \code{isodatetime} a "YYYY-MM-DD hh:mm:ss" character string.
    
    With \code{NA} a suitable column type will be determined based on the 
    \emph{first non-empty entry} of the first 16 rows of the Excelsheet.}
  \item{checkNames}{with \code{TRUE} the colNames are guaranteed to be 
    syntactically valid variable names (by \code{make.names}).}
  \item{dateTimeAs}{controls how to determine a column class for columns which are 
    formatted as date, time or datetime. Possible entries are: \code{numeric}
    or \code{isodatetime}. Will be overruled by explicit \code{colClasses} arguments.}
  \item{stringsAsFactors}{set to \code{FALSE} to retrieve plain characters. The
    default follows the global option \code{default.stringsAsFactors()}.}
}

\details{
  The arguments \code{colNames} and/or \code{colClasses} can optionally include 
  an entry for the rownames column (which will be discarded).
}

\value{
  A data.frame or a matrix of the specified type.
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
  \code{\link{write.xls}},
  \code{\link{read.table}}, \code{\link{matrix}}, \code{\link{data.frame}}
 }

\keyword{file}
