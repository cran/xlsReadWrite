\name{read.xls}
\alias{read.xls}
\title{Read Excelfiles}

\description{
  Reads an Excelsheet and returns a data.frame or a matrix. 
	Currently supported is the .xls (BIFF8) file format.
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
  \item{rowNames}{a character vector with the actual row names or a logical. 
    \code{TRUE} uses the first Excel column for the rownames, with \code{FALSE} 
    no rownames will be taken from Excel.
    
    With \code{NA} the first column will be \emph{considered} to contain row names 
    under the following three conditions: (1) \code{colNames} is \code{TRUE} or  
    contains a character vector, (2) the first column name thereof is an empty 
    string and (3) the first data cell contains a \emph{string} which is not "1".}
  \item{colClasses}{the class of the data.frame columns. Possible entries are: 
    \code{double}, \code{numeric}, \code{integer}, \code{logical}, \code{character}, 
    \code{factor}, \code{NA},  \code{isodate}, \code{isotime} and \code{isodatetime}. 
    A scalar entry will be recycled if necessary.

    \code{isodate} retrieves a "YYYY-MM-DD", \code{isotime} a "hh:mm:ss" and 
    \code{isodatetime} a "YYYY-MM-DD hh:mm:ss" character string.
    
    With \code{NA} a suitable column type will be determined based on the 
    \emph{first} non-empty entry of the 16 rows following \code{from} (starting row).}
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
  By purchasing xlsReadWritePro you help support our development and maintenance effort. 
  We are grateful that quite some people and companies did, and both, the free and pro
  version shall benefit (e.g. crossplatform, which is not done on a lazy afternoon...).
  
  Features of xlsReadWritePro: + append data to existing files, + work with in-memory objects
  (e.g. to compose Excel reports), + work with data area subsets, + use (named) ranges and pick 
  individual cells, + work with sheets (copy, rename, delete), + read/write formula values, 
  + insert images, + oleDateTime classes and more datetime helper functions, + formal support.
  
  The trial (http://www.swissr.org) is free and the two versions can easily be switched 
  and coexist just fine. So, why not check out xlsReadWritePro and see for yourself? 
  
  !! NOTE: xlsReadWritePro works well up to R 2.9.x but not yet on R 2.10 version.
  The reason is, that we are currently reorganizing internal things (git/redmine/web/distr/make).
  We started with the free version and the pro version will follow.
}

\seealso{
  \code{\link{write.xls}},
  \code{\link{read.table}}, \code{\link{matrix}}, \code{\link{data.frame}}
 }

\examples{
\dontrun{
rdata <- read.xls( rfile, colNames = myc( "one", "two", "three" ), 
                   from = 3, rowNames = FALSE, type = "double" )

## there are many more examples under 'RUnitTests/test<xy>.R'

}
}

\keyword{file}
