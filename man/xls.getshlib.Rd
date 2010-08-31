\name{xls.getshlib}
\alias{xls.getshlib}
\title{Download regular shlib (dll/so)}

\description{
    Downloads the regular shlib from our swissr dropbox account and replaces
    the currently used placeholder shlib with it. The size is about ~420 KB.

    Using the default command, \code{xls.getshlib()}, should work just
    fine. Every step is documented, i.e. will be printed on the console.

    There are quite a lot of different versions (shlib, R, (in future) platforms).
    If the correct shlib cannot be found and/or if there are permission issues,
    \code{xls.getshlib} will stop. But you can always download the full regular
    package or shlib manually from the URLs indicated in the details section below.
    Other than being less convenient it is the same. In case of any issues we are
    happy to hear about them (bug tracker/forum/email), thank you.
}

\usage{
xls.getshlib(pkgvers = NA, url = NA, md5 = TRUE,
             reload.shlib = TRUE, tmpdir = tempdir())
}

\arguments{
    \item{pkgvers}{optional character string to override the default (which is:
        (\code{packageDescription("xlsReadWrite")$Version})). \option{pkgvers} is 
        used to replace placeholder in dropbox download url.}
    \item{url}{optional character string to indicate a custom url to the zipped file 
        containing the shlib. Local, i.e. \code{file://<my url>} urls are possible. If 
        an url string has been given, the 'pkgvers' argument will not be used.}
    \item{md5}{with \code{TRUE} an md5 value will be retrieved from a file (with 
        '.md5.txt' suffix) located at the url indicated. A character string is assumed
        to be the md5 value to be used for checking. \code{FALSE} disables the check.}
    \item{reload.shlib}{\code{TRUE} to replace the loaded (placeholder) library. With
        \code{FALSE} the shlib will only be downloaded (path will be displayed).}
    \item{tmpdir}{temporary folder to downloaded file.}
}

\details{
    All shlibs, the regular packages and the source files are available at our swissr
    dropbox account, most recent versions also at the main site and the source at github. URLs:\cr
    - \url{http://dl.dropbox.com/u/2602516/swissrpkg/index.html}\cr
    - \url{http://www.swissr.org}\cr
    - \url{http://github.com/swissr/xlsreadwrite}).

    Why do you need to download the shlib from an external source? Two (three)
    reasons:\enumerate{
        \item Our own xlsReadWrite code is free (GPL-2), but we also use
            proprietary code (Flexcel, tmssoftware.com) which can only be distributed
            legally in precompiled, i.e. binary form.
        \item As CRAN \emph{'generally does not accept submissions of precompiled
            binaries due to security reasons'} we only provide a placeholder and you
            have to download the binary shlib separately.
        \item xlsReadWrite is written in Pascal and CRAN most probably would not
            be able/interested to support suitable compilers, i.e. Delphi and/or FPC.
            (We did not ask about compilers though, because the first point already
            rules out a normal open source distribution of the package).
    }

    There have been thorough tests initially but we do not give ANY GUARANTEES AT ALL.
    Eventually inspect the \code{xls.getshlib} code and/or type \code{?xlsReadWrite} to
    find instructions about how to compile the package for yourself (Delphi/Flexcel needed).
}

\section{Pro version}{
    See section in \code{\link{read.xls}}.
}

\seealso{
    \code{\link{read.xls}}, \code{\link{write.xls}}
}

\examples{\dontrun{
xls.getshlib()
}}

\keyword{file}
