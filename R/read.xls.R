"read.xls" <-
function( file, colNames = TRUE, sheet = 1, type = "data.frame", from = 1, colClasses = NA) {
    res <- .Call( "ReadXls", file, sheet, type, colNames, from - 1, colClasses )
    if (!is.null( colClasses ) && ("factor" %in% colClasses)) {
        if ((length( colClasses ) == 1) || (length( colClasses ) == length( res ))) {
            hasRowname <- FALSE
        } else if (length( colClasses ) == (length( res ) + 1)) {
            hasRowname <- TRUE
        } else {
            stop( "length of colClasses must be 1, equal or 1 more than column count in data" )
        }
        for (i in 1:length( colClasses ))
            if (colClasses[i] == "factor") res[[i - hasRowname]] <- as.factor(res[[i - hasRowname]])
    }
    res
}