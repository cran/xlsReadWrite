\name{xlsReadWrite-package}
\docType{package}
\alias{xlsReadWrite-package}
\alias{xlsReadWrite}
\title{Read and write Excel files}

\description{
    \enumerate{
        \item Overview
        \item CRAN/pro versions
        \item RUnit tests
	\item Future plans
        \item Download and compilation
        \item Links and contact
    }
}

\section{1. Overview}{
    \code{xlsReadWrite} is a packages which allows to read and write Excel files.
    Currently supported is the .xls (BIFF8) file format.
    \cr

    \bold{Getting started}\itemize{
        \item check out \code{\link{read.xls}}, \code{\link{write.xls}} and
        \code{\link{xls.getshlib}}
	\item type: \code{help(package = "xlsReadWrite")} for more information\cr
    }

    \bold{Why xlsReadWrite?}\itemize{
        \item it is fast
        \item well documented and (RUnit) tested
        \item the package does \emph{not} have any external dependencies
        \item it has a nice simple interface and \emph{just works fine} since ~4 years
    }

    On the less rosy side: xlsReadWrite still is windows only and uses proprietary
    3rd party code (which means only our own code is, and can be, free). The package
    is written in Pascal, a very nice but rather obscure language in the R world.\cr

    \bold{Other solutions}

    xlsReadWrite is not the only game in town, we know of the following packages
    (brackets indicate dependencies): \emph{RODBC} (drivers), \emph{gdata} (Perl),
    \emph{RExcel} or \emph{COM} (Excel), \emph{dataframes2xls} (Python), \emph{xlsx}
    (Apache POI, Java). Last but not least, you may use plain old \emph{.csv} files
    (none) or consider our own \emph{xlsReadWritePro} (none) shareware package.
}

\section{2. CRAN/pro versions}{

    Besides the regular xlsReadWrite package there is a special \code{CRAN}
    version.\cr
    Reason (as copied from the startup message):\cr\cr
\emph{BACKGROUND: Our own xlsReadWrite code is free, but we also use proprietary code}\cr
\emph{(Flexcel, tmssoftware.com) which can only be distributed legally in precompiled,}\cr
\emph{i.e. binary form. As CRAN 'generally does not accept submissions of precompiled}\cr
\emph{binaries due to security reasons' we only provide a placeholder and you can}\cr
\emph{download the binary shlib separately.}\cr

    The \code{xlsReadWritePro} version is our second attempt and goes further than
    the free version:\itemize{
        \item append data to existing files
        \item work with in-memory Excel 'files' (called xls-obj)\cr
        \item manage sheets (select, insert, copy, rename, delete, hide, info, list)
        \item support images (add, delete, count, list, export, info)
        \item address confined data areas (rows, cols, from, to, (named) ranges, cells)\cr
        \item read and write formula values
        \item get file and sheet related info
        \item more date time helper functions\cr
        \item formal support
    }

    By purchasing the \code{xlsReadWritePro} shareware version you help support
    our development and maintenance effort. We are grateful that quite some people
    and companies did and both, the free and the pro version shall benefit. The
    trial (http://www.swissr.org) is free and the two versions do coexist just fine.
    So, why not check out \code{xlsReadWritePro} now and see for yourself?
}

\section{3. RUnit tests}{
    There are more than 50 unit tests which should cover every aspect of the
    package and ensure and maintain code quality. The tests also serve as an
    extended reference going beyond the examples.

    You can easily run the tests on your system, the RUnit test files are located
    at:\itemize{
        \item \code{R_HOME/library/xlsReadWrite/unitTest/runit*.R}
        \item \code{R_HOME/library/xlsReadWrite/unitTest/execManually.R} (exec details)
    }
}

\section{4. Future plans}{
    xlsReadWrite is mature and no big changes are planned. Certainly some internal
    refactoring would be beneficial and some features could be backported from
    the pro version.
    
    In 2011 Delphi should be released for Mac, Linux and Win64, and we plan to
    support these platforms eventually. The underlying library is being developed
    for .xlsx and we are looking forward to it (probably as a separate package).
}

\section{5. Download and compilation}{

    \tabular{rll}{
        Current version:  \tab \tab  \url{http://www.swissr.org/download}\cr
        All resources:    \tab \tab  \url{http://dl.dropbox.com/u/2602516/swissrpkg/index.html}\cr
        Github repo:      \tab \tab  \url{http://github.com/swissr/xlsReadWrite}\cr
        CRAN:             \tab \tab  \url{http://cran.r-project.org/web/packages/xlsReadWrite/index.html}
    }

    \bold{Compilation}

    Prerequisites for compiling: Delphi 2007 and Flexcel (http://www.tmssoftware.com,
    the core library doing the Excel-work). Then follow these steps:\enumerate{
        \item Clone the github repo and initialize the RPascal submodule
        \item Adapt paths in 'include.mk'
        \item Regarding Flexcel code:\itemize{
           \item in FLXCOMPILER.INC: disable the FLEXCEL define'
           \item 'src/pas/vFlexcelCommits.txt' has the currently used Flexcel version
        }
        \item Run \code{make <target>} in cmd.exe (system path will be modified temporarily)
        \item Alternatively compile/run project in the Delphi IDE.
    }
}

\section{6. Links and contact}{

    \tabular{rll}{
        General info:  \tab \tab  \url{http://www.swissr.org}\cr
        Bugtracker:    \tab \tab  \url{http://dev.swissr.org/projects/xlsReadWrite}\cr
        Forum:         \tab \tab  \url{http://dev.swissr.org/projects/xlsReadWrite/boards}\cr
        Email:         \tab \tab  support at swissr.org\cr
    }
    November 21, 2010 / Hans-Peter
}

\keyword{package}
