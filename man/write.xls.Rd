\name{write.xls}
\alias{write.xls}
\title{Write Excelfiles}

\description{
  Saves a data.frame, matrix or vector as an Excelfile. The supported file format
  is BIFF8, i.e. Excel 97-2003.
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
  \item{rowNames}{a character vector with the actual row names or \code{TRUE} to 
    to write row names into the first Excel column.
    
    With \code{NA}) the first column will be considered to receive the row names 
    if the first entry in the character vector or the \code{rownames()} of the data 
    is \emph{not} \code{1}.}
}

\details{
  New files are based on the template \code{TemplateNew.xls} which is located in
  the same folder as the DLL (R\_HOME/library/xlsReadWrite/libs/xlsReadWrite.dll)
  or in the application data folder \code{<APPDATA>\\Treetron\\xlsReadWrite}.
  If both templates exists, the one in the application data folder has priority.
  
  The arguments \code{colNames} and/or \code{colClasses} can optionally include 
  an entry for the rownames column (which will be discarded).
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
  \code{\link{read.xls}}, \code{\link{write.table}}}
 }

\keyword{file}
