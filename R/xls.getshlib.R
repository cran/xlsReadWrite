xls.getshlib <- function(
    pkgvers = NA,
    url = NA,
    md5 = TRUE,
    reload.shlib = TRUE,
    tmpdir = tempdir())
{
    require(tools)
    printmsg <- function(x) {cat(x, "\n"); flush.console()}
    printmsg("--- xls.getshlib running... ---")

    # url template, settings and paths

    urltmpl <- "http://dl.dropbox.com/u/2602516/swissrpkg/bin/<os>/shlib/xlsReadWrite_<pkgvers>_dll.zip"
    if (is.na(pkgvers)) pkgvers <- packageDescription("xlsReadWrite")$Version
    os <- if (R.version$platform == "i386-pc-mingw32") "win32" else stop("currently only windows (32 bit) supported")

    fp <- getLoadedDLLs()$xlsReadWrite[["path"]]
    fn <- basename(fp)
    fp.backup <- paste(fp, "~", sep = "")
    fp.temp <- file.path(tmpdir, fn)
    fpzip.temp <- file.path(tmpdir, "xlsReadWrite.zip")

    # download shlib

    copyOrDownload <- function(url) {
            # local 'download.file' ignores wb mode argument, we use 'file.copy')
        if (length(grep("^file://", url))){
            if (!file.copy(sub("^file://", "", url), fpzip.temp, overwrite = TRUE))
                stop("copying '", url, "'\nto '", fpzip.temp, "' failed")
        } else {
            res <- try(download.file(url, fpzip.temp, method = "internal", quiet = TRUE, mode = "wb"), silent = TRUE)
            if (inherits(res, "try-error") || res != 0)
                stop("downloading '", url, "'\nto '", fpzip.temp, "' failed")
        }
        printmsg(paste("  - zipped shlib has been downloaded from '", url, "' to '", fpzip.temp, "'", sep = ""))
    }
    if (is.na(url)) {
        url <- sub("<os>", os, urltmpl, fixed = TRUE)
        url <- sub("<pkgvers>", pkgvers, url, fixed = TRUE)
        copyOrDownload(url)
    } else {
        copyOrDownload(url)
    }

    # check md5

    if (is.character(md5) || md5) {
        if (is.logical(md5)) {
            url <- paste(url, ".md5.txt", sep = "")
            res <- try(readLines(url), silent = TRUE)
            if (inherits(res, "try-error") || length(res) < 1)
                stop("reading '", url, "' failed")
        md5 <- res[1]
        }
        printmsg(paste("  - md5 hash value has been read from '", url, "'", sep = ""))

        stopifnot(file.exists(fpzip.temp)); stopifnot(is.character(md5))
        if (!(md5cal <- md5sum(fpzip.temp) == md5))
            stop("downloaded shlib has wrong md5 hash (", md5cal, " instead of ", md5, ")")
        printmsg("  - zipped shlib has correct md5 hash")
    } else {
        printmsg("  - WARNING: md5 check has been skipped")
    }

    # unzip shlib

    unzip(fpzip.temp, exdir = dirname(fp.temp))
    if (file.exists(fp.temp))
        printmsg(paste("  - zipped shlib has been extracted to'", fp.temp, "'", sep = ""))
    else
        stop("unzipping '", fpzip.temp, "' failed")

    # reload shlib

    if (reload.shlib) {
        manrepl <- paste("Please replace the existing library '", fp, "' with the downloaded shlib '", fp.temp, "' manually", sep = "")

            # unload dummy shared library
        if (!file.exists(fp)) stop("existing (cran) shlib could not be found at '", fp, "'")
        printmsg(paste("  - try to unload existing shlib '", fp, "'", sep = ''))
        dyn.unload(fp)

            # replace with downloaded shlib (copy because R couldn't file.rename from C: to V: drive)
        if (file.exists(fp.backup)) file.remove(fp.backup)
        if (!file.rename(fp, fp.backup))
            stop("Existing shlib has been unloaded but could not be backuped (as '", fp.backup, "')\n", manrepl)
        printmsg(paste("  - replace existing library with downloaded shlib", sep = ""))
        if (!file.copy(fp.temp, fp)) 
            stop("Could not move downloaded shlib '", fp.temp, "' to the correct place,\n existing shlib has been ",
                  "unloaded and renamed to '", fp.backup, "'")
        if (!file.remove(fp.backup)) 
            warning("Existing shlib has been unloaded, renamed and replaced with the downloaded shlib.\n",
                  "However the backuped file (", fp.backup, ") could not be deleted ",
                  "(do this manually). After restarting R everything should be ok.")

            # load again
        printmsg("  - try loading new shlib:\n")
        dyn.load(fp)
        printmsg("--- Done (shlib successfully updated) ---")

    } else {
        printmsg("--- Done (shlib successfully downloaded) ---")
        printmsg(manrepl)
    }
    invisible()
}
