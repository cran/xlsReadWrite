read.xls <- function(
    file,
    colNames = TRUE,
    sheet = 1,
    type = "data.frame",
    from = 1,
    rowNames = NA, colClasses = NA, checkNames = TRUE,
    dateTime = "numeric",
    naStrings = NA,
    stringsAsFactors = default.stringsAsFactors())
{
    res <- .Call("ReadXls", file, colNames, sheet, type, from,
                 rowNames, colClasses, checkNames,
                 dateTime, naStrings, stringsAsFactors)
    res
}