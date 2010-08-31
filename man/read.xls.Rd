\name{read.xls}
\alias{read.xls}
\title{Read Excel files}

\description{
    Reads an Excel file into a data.frame or matrix.
    Supported is the .xls (BIFF8) file format.
}

\usage{
read.xls(file, 
         colNames = TRUE, 
         sheet = 1,
         type = "data.frame",
         from = 1,
         rowNames = NA, colClasses = NA, checkNames = TRUE,
         dateTimeAs = "numeric",
         naStrings = NA,
         stringsAsFactors = default.stringsAsFactors())
}

\arguments{
    \item{file}{name of an Excel file. Path may be absolute or relative
        to the current working directory.}
    \item{colNames}{with \code{TRUE} the first row of the sheet or the
        'from'-starting area will be used for the colnames. Provide a
        character vector to use custom colnames. The above colname values may
        possibly get modified depending on the \code{checkNames} argument.
        With \code{FALSE} and for missing values, defaults will be used, i.e.
        a 'V' followed by the column number.}
    \item{sheet}{case sensitive character string or one-based number indicating
        the Excel sheet to read from.}
    \item{type}{a character string indicating the type of the returned object.
        Either \code{data.frame} or else \code{double}, \code{integer}, 
        \code{logical} or \code{character} to get a matrix.}
    \item{from}{row to start reading from.}
    \item{rowNames}{with \code{TRUE} the first column of the sheet will be used
        for the rownames. Provide a character vector to use custom rownames. The
        above rowname values may possibly get modified depending on the
        \code{checkNames} argument. With \code{FALSE} defaults will be used, i.e.
        numbers starting from one.

        With \code{NA} the first sheet column will be considered to 
        contain rownames under the following conditions: 1) \code{colNames}
        is \code{TRUE} or contains a character vector, 2) there are at least
        two columns and the column name for the potential rowname-column is empty, 
        3) the first data cell contains a string which 4) is not "1".}
    \item{colClasses}{specify the column types of the resulting data.frame.
        Possible entries are: \code{double}, \code{numeric}, \code{integer},
        \code{logical}, \code{character}, \code{factor}, \code{NA},
        \code{isodate}, \code{isotime} and \code{isodatetime}. Either supply
        an entry for each column or else a scalar which will be recycled.

        With \code{NA} suitable column types will be determined based on the
        \emph{first} non-empty cell of the 16 rows following the starting (from)
        row. An integer will be recognised as 'numeric' (this is different to
        the pro version where we examine \emph{all} 16 values and keep integers
        as 'integers'). \code{NA} can be used together with other entries, e.g.
        \code{c(NA, "double", NA, "isodate")}.

        If nothing can be determined, all values become NA (logical) and a warning
        message will be printed.}
    \item{checkNames}{logical. With \code{TRUE} col- and rownames are guaranteed
        to be syntactically valid variable names, e.g. by prepending an "X" or
        translating invalid characters to "." (\code{\link{make.names}} is being used).}
    \item{dateTimeAs}{scalar character controlling how Excel date values will be
        interpreted when no explicit colClasses have been given: \code{isodatetime} or
        \code{numeric}. 'isodatetime'  will possibly be shortened to 'isodate' or 'isotime'
        and 'numeric' to 'integer'.}
    \item{naStrings}{a character vector of strings which are to be interpreted as
        NA-values. \code{NA} disables this interpretation, i.e. an Excel 'NA' string
        becomes a character or factor.}
    \item{stringsAsFactors}{logical, the default follows the global option
        \code{\link{default.stringsAsFactors}}.}
}

\details{
    The character string arguments for \code{colNames} and \code{colClasses}
    can optionally include an entry for the column used for the rownames (entry
    will be discarded). For \code{rowNames}, the character vector length must be
    the same as the number of data rows.
}

\value{
    A data.frame or a matrix of the specified \code{type}. \code{NULL} if the
    sheet is empty.
}

\section{Pro version}{
    The pro version is our second attempt and goes further than the free version:
    + append data to existing files, + work with in-memory objects, e.g. to compose
    Excel reports with multiple sheets, + address confined data areas: rows, cols
    and from, to; (named) ranges and picking from individual cells, + manage sheets
    (select, insert, copy, rename, delete, hide), + read/write formula values,
    + support images (add, delete, count, list and export), + additional oleDateTime
    helper functions, + last but not least: formal support contact.

    By purchasing the \code{xlsReadWritePro} shareware version you help support our
    development and maintenance effort. We are grateful that quite some people
    and companies did and both, the free and the pro version shall benefit. The trial
    (http://www.swissr.org) is free and the two versions do coexist just fine. So, why
    not check out \code{xlsReadWritePro} now and see for yourself?
}

\seealso{
    \code{\link{write.xls}},
    \code{\link{read.table}}, \code{\link{matrix}}, \code{\link{data.frame}}
}

\examples{
# --- only run for regular, i.e. non-cran version
shlib <- system.file("libs", if (nzchar(arch <- .Platform$r_arch)) arch else "",
                     paste("xlsReadWrite", .Platform$dynlib.ext, sep = ""), package = "xlsReadWrite")
if (file.exists(shlib) && (file.info(shlib)$size > 20000)) {
# -----------------------------------------------

    # path to Excel file
rfile <- system.file("unitTests/data/origData.xls", package = "xlsReadWrite")

# read as data.frame and as double

read.xls(rfile)
read.xls(rfile, type = "double")

# read as integer with custom col-/rownames 

(rdata <- read.xls(rfile, from = 3, type = "integer",
                   colNames = c("one", "one", "three", "four"), 
                   rowNames = paste("r", 1:12, sep = "")))
    # 2nd 'one' ends as 'one.1' due to 'checkNames'
stopifnot(colnames(rdata)[2] == "one.1")
}}

\keyword{file}
