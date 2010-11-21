\name{write.xls}
\alias{write.xls}
\title{Write Excel files}

\description{
    Saves a data.frame, matrix or vector as an Excel file.
    Currently supported is the .xls (BIFF8) file format.
}

\usage{
write.xls(x, file,
          colNames = TRUE,
          sheet = 1,
          from = 1,
          rowNames = NA,
          naStrings = "")
}

\arguments{
    \item{x}{data to be written. A data.frame or else a matrix or vector
        of the type double, integer, logical or character. Vectors will
        be written in columns.}
    \item{file}{name of an Excel file. Path may be absolute or relative
        to the current working directory.}
    \item{colNames}{with \code{TRUE} the data colnames will be written into
        the first row of the sheet or the 'from'-starting area. Provide a character
        vector to write custom colnames. With \code{FALSE} no colnames will be written.}
    \item{sheet}{case sensitive character string or one-based number indicating
        the Excel sheet to write to.}
    \item{from}{row to start writing from.}
    \item{rowNames}{with \code{TRUE} the data rownames will be written into
        the first column of the sheet. Provide a character vector to write
        custom rownames. With \code{FALSE} no rownames will be written.

        With \code{NA} the first column will be considered to receive the
        rownames under the following conditions: 1) \code{colNames} is \code{TRUE}
        or contains a character vector, 2) there are character rownames (supplied or
        in the data) and 3) the first entry thereof is not \code{"1"}.}
    \item{naStrings}{the string to be used for \code{NA} values. An empty string
        clears (blanks) the cell, with \code{NA} the cell value will not be changed
        (only relevant for an eventual future append capability).}
}

\details{
    New files are based on the template \code{TemplateNew.xls}. It can be modified and
    is located at \code{R\_HOME/library/xlsReadWrite/template} (internal search is relative
    to the shlib located at R\_HOME/library/xlsReadWrite/libs/[<arch>/]).

    The character string arguments for \code{colNames} can optionally include an entry
    for the column used for the rownames (entry will be discarded). For \code{rowNames},
    the character vector length must be the same as the number of data rows.
}

\section{Pro version}{
    See section in \code{\link{read.xls}}.
}

\seealso{
    \code{\link{read.xls}}, \code{\link{write.table}}
}

\examples{
# --- only run for regular, i.e. non-cran version
shlib <- system.file("libs", if (nzchar(arch <- .Platform$r_arch)) arch else "",
  paste("xlsReadWrite", .Platform$dynlib.ext, sep = ""), package = "xlsReadWrite")
if (file.exists(shlib) && (file.info(shlib)$size > 20000)) {
# -----------------------------------------------

myval <- data.frame(
    Fertility = c(80.2, 83.1, 92.5),
    Agriculture = c(17, 45.1, 39.7),
    Testlogical = c(TRUE, TRUE, FALSE),
    Education = as.integer(c(12, 9, 5)),
    Catholic = c(9.96, 84.84, 93.4),
    Infant.Mortality = c(22.2, 22.2, 20.2),
    Testcharacter = c("Co", "De", "Fr"), stringsAsFactors = FALSE)

# write the data.frame...

write.xls(myval, "mytest.xls")

# ..then read and check (colClasses because we want logical and integer (for
# double) and character (for factor; could also have given stringsAsFactors))

mycls <- c("double", "double", "logical", "integer", "double", "double", "character")
wdata <- read.xls("mytest.xls", colClasses = mycls)
stopifnot(identical(wdata, myval))
}}

\keyword{file}
