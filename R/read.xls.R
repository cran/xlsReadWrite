"read.xls" <-
function( file, colNames = TRUE, sheet = 1, type = "data.frame", from = 1 ) {
  res <- .Call( "ReadXls", file, sheet, type, colNames, from - 1 )
  res
}
