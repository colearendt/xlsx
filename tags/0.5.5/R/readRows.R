# Useful when you have a "fat" spreadsheet (few rows and many columns.) 
#
# Only reads Strings for now.
#
#

readRows <- function(sheet, startRow, endRow, startColumn,
  endColumn=NULL)
{
  row1 <- sheet$getRow(as.integer(startRow-1))   # first row
  if (is.null(row1))
    stop("First row, with index ", startRow, " is empty.  Please check!")
  
  trueEndColumn <- row1$getLastCellNum()
  if (is.null(endColumn))     # get it from the first row 
    endColumn <- trueEndColumn
  
  noRows <- endRow - startRow + 1

  Rintf <- .jnew("org/cran/rexcel/RInterface")  # create an interface object 
  
  res <- matrix(NA, nrow=noRows, ncol=endColumn-startColumn+1)
  for (i in seq_len(noRows)) {
    aux <- .jcall(Rintf, "[S", "readRowStrings",
      .jcast(sheet, "org/apache/poi/ss/usermodel/Sheet"),
      as.integer(startColumn-1), as.integer(endColumn-1), 
      as.integer(startRow-1+i-1))

    res[i,] <- aux
  }
  
  res
}


  ## if (endColumn > trueEndColumn) {
  ##   warning(paste("First row requested has only", trueEndColumn, "columns."))
  ##   endColumn <- min(endColumn, trueEndColumn)
  ## }
  
