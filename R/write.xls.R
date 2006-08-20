"write.xls" <-
function( x, file, colNames = TRUE, sheet = 1, from = 1 ) {
  invisible( .Call( "WriteXls", x, file, sheet, colNames, from - 1 ) )
}
