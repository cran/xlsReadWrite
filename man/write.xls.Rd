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
  \item{file}{the name of a local file}
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
  is \emph{not} "1", a rowname-column will be written in Excel.
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
  \code{\link{read.xls}}, \code{\link{write.table}}}
 }

\keyword{file}
