"write.xls" <-
function( x, file, colNames = TRUE, sheet = 1, from = 1, rowNames = NA ) 
{
    invisible( .Call( "WriteXls", x, file, colNames, sheet, from - 1, rowNames ) )
}
