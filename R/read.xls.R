"read.xls" <-
function( file, colNames = TRUE, sheet = 1, type = "data.frame", from = 1,
    rowNames = NA, colClasses = NA, checkNames = TRUE,
    dateTimeAs = "numeric", stringsAsFactors = default.stringsAsFactors() )
{
    res <- .Call( "ReadXls", file, colNames, sheet, type, from, 
                  rowNames, colClasses, checkNames,
                  dateTimeAs, stringsAsFactors )
    res
}