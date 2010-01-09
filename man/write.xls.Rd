\name{write.xls}
\alias{write.xls}
\title{Write Excelfiles}

\description{
  Saves a data.frame, matrix or vector as an Excelfile.
  Currently supported is the .xls (BIFF8) file format.

}

\usage{
write.xls( x, file, 
           colNames = TRUE, 
           sheet = 1, 
           from = 1, 
           rowNames = NA )
}

\arguments{
  \item{x}{data to be written. A data.frame or a matrix/vector of type double, 
    integer, logical or character.}
  \item{file}{the name of the file. If it does not contain an absolute path, the 
    file name is relative to the current working directory.}
  \item{colNames}{a character vector or \code{TRUE} to write a title row in the
    spreadsheet.}
  \item{sheet}{to write to. A case sensitive character string or a number.}
  \item{from}{specifies the starting row to write to.}
  \item{rowNames}{a character vector with the actual row names or a logical. 
    \code{TRUE} writes the rownames in the first Excel column, with \code{FALSE} 
    the rownames won't be written. With \code{NA} the first column will be considered
    to receive the row names under the following three conditions: 1) \code{colNames} 
    is \code{TRUE} or contains a character vector, 2) there are character rownames 
    (in the data.frame/matrix) and 3) the first entry thereof is not \code{"1"}.}
}

\details{
  New files are based on the template \code{TemplateNew.xls} which is located in
  a subfolder of the DLL (\code{R\_HOME/library/xlsReadWrite/libs/xlsReadWrite.dll/template})
  or in the application data folder \code{<APPDATA>/Treetron/xlsReadWrite}.
  If both templates exists, the one in the application data folder has priority.
  
  The arguments \code{colNames} and/or \code{colClasses} can optionally include 
  an entry for the rownames column (which will be discarded).
}

\section{Pro version}{
  See section under \link{read.xls}.
}

\seealso{
  \code{\link{read.xls}}, \code{\link{write.table}}
}

\examples{
\dontrun{
myval <- data.frame( 
		Fertility = c(80.2, 83.1, 92.5), 
  	Agriculture = c(17, 45.1, 39.7), 
  	Testlogical = c(T,T,F), 
  	Education = as.integer( c(12, 9, 5) ), 
  	Catholic = c(9.96, 84.84, 93.4), 
  	Infant.Mortality = c(22.2, 22.2, 20.2), 
  	Testcharacter = c("Co", "De", "Fr" ), stringsAsFactors = FALSE )
write.xls( myval, "mytest.xls" )

mycls <-  c( "double", "double", "logical", "integer", "double", "double", "character" )
wdata <- read.xls( "mytest.xls", colClasses = mycls )
rownames(myval) <- as.character(rownames(myval))
stopifnot( identical( wdata, myval ) )

## there are many more examples under 'inst/RUnitTests/test<xy>.R'

}
}

\keyword{file}
