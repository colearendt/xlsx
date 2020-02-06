#
#
# For CellType: FORMULA, ERROR or blanks, it defaults to "character"
#
#

#' Read a contiguous set of columns from sheet into an R data.frame
#'
#' Read a contiguous set of columns from sheet into an R data.frame.  Uses the
#' \code{RInterface} for speed.
#'
#'
#' Use the \code{readColumns} function when you want to read a rectangular
#' block of data from an Excel worksheet.  If you request columns which are
#' blank, these will be read in as empty character "" columns. Internally, the
#' loop over columns is done in R, the loop over rows is done in Java, so this
#' function achieves good performance when number of rows >> number of columns.
#'
#' Excel internally stores dates and datetimes as numeric values, and does not
#' keep track of time zones and DST.  When a numeric column is formatted as a
#' datetime, it will be converted into \code{POSIXct} class with a \emph{GMT}
#' timezone.  If you need a \code{Date} column, you need to specify explicitly
#' using \code{colClasses} argument.
#'
#' For a numeric column Excels's errors and blank cells will be returned as NaN
#' values.  Excel's \code{#N/A} will be returned as NA.  Formulas will be
#' evaluated. For a chracter column, blank cells will be returned as "".
#'
#' @param sheet a \code{\link{Worksheet}} object.
#' @param startColumn a numeric value for the starting column.
#' @param endColumn a numeric value for the ending column.
#' @param startRow a numeric value for the starting row.
#' @param endRow a numeric value for the ending row.  If \code{NULL} it reads
#' all the rows in the sheet.  If you request more than the existing rows in
#' the sheet, the result will be truncated by the actual row number.
#' @param as.data.frame a logical value indicating if the result should be
#' coerced into a \code{data.frame}.  If \code{FALSE}, the result is a list
#' with one element for each column.
#' @param header a logical value indicating whether the first row corresponding
#' to the first element of the \code{rowIndex} vector contains the names of the
#' variables.
#' @param colClasses a character vector that represents the class of each
#' column.  Recycled as necessary, or if \code{NA} an attempt is made to guess
#' the type of each column by reading the first row of data.  Only
#' \code{numeric}, \code{character}, \code{Date}, \code{POSIXct}, column types
#' are accepted.  Anything else will be coverted to a \code{character} type.
#' If the length is less than the number of columns requested, replicate it.
#' @param ... other arguments to \code{data.frame}, for example
#' \code{stringsAsFactors}
#' @return A data.frame or a list, depending on the \code{as.data.frame}
#' argument.
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
#'   res <- readColumns(sheet, startColumn=3, endColumn=10, startRow=3,
#'     endRow=7)
#'
#'   sheet <- sheets[["NAs"]]
#'   res <- readColumns(sheet, 1, 6, 1,  colClasses=c("Date", "character",
#'     "integer", rep("numeric", 2),  "POSIXct"))
#'
#'
#' }
#'
#' @export
readColumns <- function(sheet, startColumn, endColumn, startRow,
  endRow=NULL, as.data.frame=TRUE, header=TRUE, colClasses=NA, ...)
{
  trueEndRow <- sheet$getLastRowNum() + 1
  if (is.null(endRow))     # get it from the sheet
    endRow <- trueEndRow

  if (endRow > trueEndRow) {
    warning(paste("This sheet has only", trueEndRow, "rows"))
    endRow <- min(endRow, trueEndRow)
  }

  noColumns <- endColumn - startColumn + 1

  Rintf <- .jnew("org/cran/rexcel/RInterface")  # create an interface object

  if (header) {
    cnames <- try(.jcall(Rintf, "[S", "readRowStrings",
      .jcast(sheet, "org/apache/poi/ss/usermodel/Sheet"),
      as.integer(startColumn-1), as.integer(endColumn-1),
      as.integer(startRow-1)))
    if (class(cnames) == "try-error")
      stop(paste("Cannot read the header. ",
        "The header row doesn't have all requested cells non-blank!"))
    startRow <- startRow + 1
  } else {
    cnames <- as.character(1:noColumns)
  }

  # guess or expand colClasses
  if (is.na(colClasses)[1]) {
    row <- getRows(sheet, rowIndex=startRow)
    cells <- getCells(row, colIndex=startColumn:endColumn)
    if (length(cells) != noColumns)
      warning("Not enough columns in the first row of data to correctly guess colClasses!")
    colClasses <- .guess_cell_type(cells)
  }
  colClasses <- rep(colClasses, length.out=noColumns)  # extend colClasses

  res <- vector("list", length=noColumns)
  for (i in seq_len(noColumns)) {
    if (any(c("numeric", "POSIXct", "Date") %in% colClasses[i])) {
      aux <- .jcall(Rintf, "[D", "readColDoubles",
        .jcast(sheet, "org/apache/poi/ss/usermodel/Sheet"),
        as.integer(startRow-1), as.integer(endRow-1),
        as.integer(startColumn-1+i-1))
      if (colClasses[i]=="Date")
        aux <- as.Date(aux-25569, origin="1970-01-01")
      if (colClasses[i]=="POSIXct")
        aux <- as.POSIXct((aux-25569)*86400, tz="GMT", origin="1970-01-01")
    } else {
      aux <- .jcall(Rintf, "[S", "readColStrings",
        .jcast(sheet, "org/apache/poi/ss/usermodel/Sheet"),
        as.integer(startRow-1), as.integer(endRow-1),
        as.integer(startColumn-1+i-1))
    }
#browser()
    if (!is.na(colClasses[i]))
      suppressWarnings(class(aux) <- colClasses[i])  # if it gets specified
    res[[i]] <- aux
  }

  if (as.data.frame) {
    cnames[cnames == ""] <- " " # remove some silly c.......... colnames!
    names(res) <- cnames
    res <- data.frame(res, ...)
  }

  res
}
