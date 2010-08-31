\name{xlsReadWrite-package}
\docType{package}
\alias{xlsReadWrite-package}
\alias{xlsReadWrite}
\title{Read and write Excel files}


\description{
    \enumerate{
        \item Overview
        \item Links and downloads
        \item Tests and compiling
        \item CRAN version and xlsReadWritePro
        \item Contact/Feedback
    }
}

\section{1. Overview}{
    \code{xlsReadWrite} is a packages which allows to read and write Excel files.
    Currently supported is the .xls (BIFF8) file format.

    Please type \code{help(package = "xlsReadWrite")}, \code{?read.xls}, \code{?write.xls}
    and \code{?xls.getshlib} to learn about the package and the actual commands and usage.

    The main plus points:\itemize{
        \item it is fast
        \item well documented and (RUnit) tested
        \item the package does \emph{not} have any external dependencies
        \item it has a nice simple interface and \emph{just works fine} (since ~4 years)
    }

    Regarding the less rosy side: \code{xlsReadWrite} - still - is windows
    only and uses proprietary 3rd party code (which means only our own code is,
    and can be, free). The package is written in Pascal, a very nice language,
    but a bit an obscure thing in the R world.

    In 2011 Delphi should be released for Mac/Linux incl. a 64 bit preview and we
    plan to support those platforms then (depending on job requirements, maybe
    already earlier through FreePascal). The underlying library is being developed
    for .xlsx (quite a long time already) and so there is hope to support the .xlsx
    format sometime in future. Probably as a separate package, but we will see.

    xlsReadWrite is not the only possibility to interface Excel. Besides the
    \emph{pro} shareware version (see below) there are the following packages:
    \emph{RODBC} (drivers), \emph{gdata} (Perl), \emph{RExcel} or \emph{COM} (Excel),
    \emph{dataframes2xls} (Python), \emph{xlsx} (Apache POI, Java) and, of course,
    using plain old \emph{.csv} files. Note, dependencies have been mentioned in brackets.
}

\section{2. Links and downloads}{
    \tabular{rll}{
        General info:   \tab\tab		\url{http://www.swissr.org} \cr
        Downloads:	    \tab\tab	    \url{http://dl.dropbox.com/u/2602516/swissrpkg/index.html}\cr
        .\tab\tab.\cr
        Bugtracker:  	\tab\tab		\url{http://dev.swissr.org/projects/xlsReadWrite}\cr
        Forum:          \tab\tab        \url{http://dev.swissr.org/projects/xlsReadWrite/boards}\cr
        .\tab\tab.\cr
        CRAN version:	\tab\tab		\url{http://cran.r-project.org/web/packages/xlsReadWrite/index.html}\cr
        Source code: 	\tab\tab		\url{http://github.com/swissr/xlsReadWrite}\cr
    }
}

\section{3. Tests and compiling}{
    TESTS

    The enclosed RUnit tests help us maintaining the quality and kind of secure what happens in certain
    circumstances, e.g. when column types for data.frames will be determined. Or with the, sometimes,
    quite complicated col-/rownames behaviour. Otherwise they provide additional user examples and give
    assurance that things work correctly on a local system. 

    You may run the approx. 50 unit tests with the following code:

    \preformatted{
            # adapt path
        ruDir <- "C:/Programme/R/R-2.11.0/library/xlsReadWrite/unitTests" 

            # note: RUnit package needed
        library(RUnit); library(xlsReadWrite)

            # source and run (output goes to current working dir) 
        source(file.path(ruDir, "loadRUnit.R"))
        .setup(ruDir, getwd())
        execTestSuite(ruDir, getwd())
    }

    COMPILING

    Prerequisites: Delphi 2007 and Flexcel (http://www.tmssoftware.com, a high-quality
    code library for the core Excel-work).

    Then follow these steps:\enumerate{
        \item Clone the github repo and initialize the RPascal submodule
        \item Adapt paths in 'include.mk'
        \item Make sure that the RTools are in your path (as first entry!)
        \item Regarding Flexcel code:\itemize{
           \item in FLXCOMPILER.INC: disable the FLEXCEL define, e.g. '{.$DEFINE FLEXCEL}'
           \item 'src/pas/vFlexcelCommits.txt' has the currently used Flexcel version
        }
        \item Compile:\itemize{
           \item open a console in the root folder and execute 'make build-reg' or
           \item open the Delphi project file, adapt paths, host application, and press F9...
        }
    }
}

\section{4. CRAN version and xlsReadWritePro}{
    CRAN

    Besides the \emph{regular} \code{xlsReadWrite} package there is a special \emph{cran} version.
    Reason (copied from the cran version startup message):\preformatted{
        "BACKGROUND: Our own xlsReadWrite code is free, but we also use proprietary code\n"
        "(Flexcel, tmssoftware.com) which can only be distributed legally in precompiled,\n"
        "i.e. binary form. As CRAN 'generally does not accept submissions of precompiled\n"
        "binaries due to security reasons' we only provide a placeholder and you can\n"
        "download the binary shlib separately.\n\n";
    }

    XLSREADWRITEPRO

    The pro version is our second attempt and goes further than the free version:\itemize{
        \item append data to existing files, 
        \item work with in-memory objects, e.g. to compose Excel reports with multiple sheets,
        \item address confined data areas: rows, cols and from, to; (named) ranges and picking from individual cells, 
        \item manage sheets (select, insert, copy, rename, delete, hide),
        \item read/write formula values,
        \item support images (add, delete, count, list and export),
        \item additional oleDateTime helper functions,
        \item last but not least: formal support contact.
    }

    By purchasing the \code{xlsReadWritePro} shareware version you help support our
    development and maintenance effort. We are grateful that quite some people
    and companies did and both, the free and the pro version shall benefit. The trial
    (http://www.swissr.org) is free and the two versions do coexist just fine. So, why
    not check out \code{xlsReadWritePro} now and see for yourself?
}

\section{5. Contact/feedback}{
    The preferred communication channel is the forum
    (\url{http://dev.swissr.org/projects/xlsreadwrite/boards}). Otherwise by email,
    \emph{support at swissr.org} or - for pro version users -
    \emph{hans-peter.suter at treetron.ch}.

    August 31, 2010 / Hans-Peter
}

\keyword{package}
