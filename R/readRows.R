# Useful when you have a "fat" spreadsheet (few rows and many columns.)
#
# Only reads Strings for now.
#
#

#' Read a contiguous set of rows into an R matrix
#'
#' Read a contiguous set of rows into an R character matrix.  Uses the
#' \code{RInterface} for speed.
#'
#'
#' Use the \code{readRows} function when you want to read a row or a block
#' block of data from an Excel worksheet.  Internally, the loop over rows is
#' done in R, and the loop over columns is done in Java, so this function
#' achieves good performance when number of rows << number of columns.
#'
#' In general, you should prefer the function \code{\link{readColumns}} over
#' this one.
#'
#' @param sheet a \code{\link{Worksheet}} object.
#' @param startRow a numeric value for the starting row.
#' @param endRow a numeric value for the ending row.  If \code{NULL} it reads
#' all the rows in the sheet.
#' @param startColumn a numeric value for the starting column.
#' @param endColumn a numeric value for the ending column.  Empty cells will be
#' returned as "".
#' @return A character matrix.
#' @author Adrian Dragulescu
#' @seealso \code{\link{read.xlsx2}} for reading entire sheets.  See also
#' \code{\link{addDataFrame}} for writing a \code{data.frame} to a sheet.
#' @examples
#'
#' \dontrun{
#'
#'   file <- system.file("tests", "test_import.xlsx", package = "xlsx")
#'
#'   wb     <- loadWorkbook(file)
#'   sheets <- getSheets(wb)
#'
#'   sheet <- sheets[["all"]]
#'   res <- readRows(sheet, startRow=3, endRow=7, startColumn=3, endColumn=10)
#'
#'
#'
#' }
#'
#' @export
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

