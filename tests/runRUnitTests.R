#
# execute RUnit tests
#

pkg <- "xlsReadWrite"
shlib <- system.file("libs", if (nzchar(arch <- .Platform$r_arch)) arch else "",
                     paste(pkg, .Platform$dynlib.ext, sep = ""), package = pkg)
stopifnot(file.exists(shlib))

if (file.info(shlib)$size < 20000)  {
        # cran version with dummy shlib (msg appears in the log)
    message("tests not executed (cran placeholder shlib)")
} else if (require("RUnit", quietly = TRUE)) {
    rutdir <- system.file("unitTests", package = pkg)
    stopifnot(file.exists(rutdir), file.info(rutdir)$isdir)
    source(system.file("unitTests", "loadRUnit.R", package = pkg))
    execTestSuiteCHECK(rutdir)
} else {
    stop("tests not executed ('RUnit' package is required)")
}