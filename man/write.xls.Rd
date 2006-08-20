\name{write.xls}
\alias{write.xls}
\title{Data Output}

\description{
  Saves a data.frame, a matrix or a vector as a native Excelfile.
  The supported format is BIFF8, i.e. Excel v97-2003.
}

\usage{
write.xls( x, file, colNames = TRUE, sheet = 1, from = 1 )
}

\arguments{
  \item{x}{a data.frame, a matrix or a vector of the types double, integer,
    logical or character.}
  \item{file}{the name of the file}
  \item{sheet}{to write into. Either a number (starting from 1) or a character
    string naming the sheet. If no sheet with this name is found, a new one will
    be inserted at the first position.}
  \item{colNames}{indicates if a header row should be written in Excel. Either
    a logical value (the columnnames or default values are taken) or or a character
    vector with userdefined column names.}
  \item{from}{specifies the first row to write to in Excel (slightly different
    from the pro version where the upper left corner (row/col) can be indicated).}
}

\details{
  Any new files are based on the template \code{TemplateNew.xls} which can be
  adapted to your specific needs. The file has to be in the same folder as the DLL
  (normally R\_HOME/library/xlsReadWrite/libs).

  For data.frames there is an auto-row feature: if the first entry in rownames
  is \emph{not} "1" a rowname-column will be written in Excel.
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
  \code{\link{read.xls}}, \code{\link{write.table}}}
 }

\keyword{file}
