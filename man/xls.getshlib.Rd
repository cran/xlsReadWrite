\name{xls.getshlib}
\alias{xls.getshlib}
\title{Download regular shlib (dll/so)}

\description{
  Downloads regular shlib from our swissr dropbox account and replaces
  cran placeholder version with it. For background info see details.
}

\usage{
xls.getshlib( pkgvers = NA, url = NA, md5 = TRUE, reload.shlib = TRUE, 
							lib.loc = .libPaths(), tmpdir = tempdir() )
}

\arguments{
  \item{pkgvers}{optional character string to override the default (which is:
  	(\code{packageDescription( "xlsReadWrite" )$Version})). \option{pkgvers} is 
  	used to replace placeholder in dropbox download url.}
  \item{url}{optional character string to indicate a custom url to the zipped file 
  	containing the shlib. Local, i.e. \code{file://<my url>} urls are possible. If 
  	an url string has been given, the 'pkgvers' argument will not be used.}
  \item{md5}{with \code{TRUE} an md5 value will be retrieved from a file (with 
  	'.md5.txt' suffix) located at the url indicated. A character string is assumed
  	to be the md5 value to be used for checking. \code{FALSE} disables the check.}
  \item{reload.shlib}{\code{TRUE} to replace the loaded (placeholder) library. With
    \code{FALSE} the shlib will only be downloaded (path will be displayed).}
  \item{lib.loc}{library tree which contains the xlsReadWrite package.}
  \item{tmpdir}{temporary folder to downloaded file.}
}

\details{
  \code{xls.getshlib()} will download, check and install the regular shlib just
  fine. You shouldn't need to provide additional arguments.
  
  The shlib file is hosted on our swissr dropbox account with the following url:
  \code{http://dl.dropbox.com/u/2602516/swissrpkg/bin/<os>/shlib/xlsReadWrite_<pkgvers>_dll.zip}
  
  There have been thorough tests initially but we do not give ANY GUARANTEES AT ALL. You
  might want to revise the 'xls.getshlib' code and/or checkout the package help. In the
  README file you'll find instructions how to compile the package for yourself.
  
  Why is such an 'install-procedure' necessary? Two reasons:
  \enumerate{
  \item Our own xlsReadWrite code is free, but we also use a proprietary code library (Flexcel, 
  	tmssoftware.com) which can only be distributed legally in precompiled, i.e. binary form. 
  	However cran 'generally does not accept submissions of precompiled binaries due to security 
  	reasons' (see R-exts and R-FAQ).
  \item xlsReadWrite is written in Pascal and cran most probably wouldn't be able/interested to
    support suitable compilers, i.e. Delphi and/or FPC. (We didn't ask about compilers though, 
    because the first point already rules out a normal open source distribution of the package.}
}

\section{Pro version}{
  See section under \link{read.xls}.
}

\seealso{
  \code{\link{read.xls}}, \code{\link{write.xls}}
}

\examples{
\dontrun{
xls.getshlib()
}
}

\keyword{file}
