# Write a data.frame to a new xlsx file.
# A convenience function.
#

# System.gc()   # to call the garbage collection in Java

#######################################################################
# NO rownames for this function.  Just the contents of the data.frame!
# have showNA=TRUE for legacy reason
#
.write_block <- function(wb, sheet, y, rowIndex=seq_len(nrow(y)),
   colIndex=seq_len(ncol(y)), showNA=TRUE)
{
  rows  <- createRow(sheet, rowIndex)      # create rows
  cells <- createCell(rows, colIndex)      # create cells

  for (ic in seq_len(ncol(y)))
    mapply(setCellValue, cells[seq_len(nrow(cells)), colIndex[ic]], y[,ic], FALSE, showNA)

  # Date and POSIXct classes need to be formatted
  indDT <- which(sapply(y, function(x) inherits(x, "Date")))
  if (length(indDT) > 0) {
    dateFormat <- CellStyle(wb) + DataFormat(getOption("xlsx.date.format"))
    for (ic in indDT){
      lapply(cells[seq_len(nrow(cells)),colIndex[ic]], setCellStyle, dateFormat)
    }
  }

  indDT <- which(sapply(y, function(x) inherits(x, "POSIXct")))
  if (length(indDT) > 0) {
    datetimeFormat <- CellStyle(wb) + DataFormat(getOption("xlsx.datetime.format"))
    for (ic in indDT){
      lapply(cells[seq_len(nrow(cells)),colIndex[ic]], setCellStyle, datetimeFormat)
    }
  }

}


#######################################################################
# High-level API
#
#' Write a data.frame to an Excel workbook.
#'
#' Write a \code{data.frame} to an Excel workbook.
#'
#'
#' This function provides a high level API for writing a \code{data.frame} to
#' an Excel 2007 worksheet.  It calls several low level functions in the
#' process.  Its goal is to provide the conveniency of
#' \code{\link[utils]{write.csv}} by borrowing from its signature.
#'
#' Internally, \code{write.xlsx} uses a double loop in over all the elements of
#' the \code{data.frame} so performance for very large \code{data.frame} may be
#' an issue.  Please report if you experience slow performance.  Dates and
#' POSIXct classes are formatted separately after the insertion.  This also
#' adds to processing time.
#'
#' If \code{x} is not a \code{data.frame} it will be converted to one.
#'
#' Function \code{write.xlsx2} uses \code{addDataFrame} which speeds up the
#' execution compared to \code{write.xlsx} by an order of magnitude for large
#' spreadsheets (with more than 100,000 cells).
#'
#' The default formats for Date and DateTime columns can be changed via the two
#' package options \code{xlsx.date.format} and \code{xlsx.datetime.format}.
#' They need to be specified in Java date format
#' \url{https://docs.oracle.com/javase/7/docs/api/java/text/SimpleDateFormat.html}.
#'
#' Writing of password protected workbooks is supported for Excel 2007 OOXML
#' format only.  Note that in Linux, LibreOffice is not able to read password
#' protected spreadsheets.
#'
#' @param x a \code{data.frame} to write to the workbook.
#' @param file the path to the output file.
#' @param sheetName a character string with the sheet name.
#' @param col.names a logical value indicating if the column names of \code{x}
#' are to be written along with \code{x} to the file.
#' @param row.names a logical value indicating whether the row names of
#' \code{x} are to be written along with \code{x} to the file.
#' @param append a logical value indicating if \code{x} should be appended to
#' an existing file.  If \code{TRUE} the file is read from disk.
#' @param showNA a logical value.  If set to \code{FALSE}, NA values will be
#' left as empty cells.
#' @param ... other arguments to \code{addDataFrame} in the case of
#' \code{read.xlsx2}.
#' @param password a String with the password.
#' @author Adrian Dragulescu
#' @seealso \code{\link{read.xlsx}} for reading \code{xlsx} documents.  See
#' also \code{\link{addDataFrame}} for writing a \code{data.frame} to a sheet.
#' @examples
#'
#' \dontrun{
#'
#' file <- paste(tempdir(), "/usarrests.xlsx", sep="")
#' res <- write.xlsx(USArrests, file)
#'
#' # to change the default date format
#' oldOpt <- options()
#' options(xlsx.date.format="dd MMM, yyyy")
#' write.xlsx(x, sheet) # where x is a data.frame with a Date column.
#' options(oldOpt)      # revert back to defaults
#'
#' }
#'
#' @export
write.xlsx <- function(x, file, sheetName="Sheet1",
  col.names=TRUE, row.names=TRUE, append=FALSE, showNA=TRUE,
  password=NULL)
{
  if (!is.data.frame(x))
    x <- data.frame(x)    # just because the error message is too ugly

  iOffset <- jOffset <- 0
  if (col.names)
    iOffset <- 1
  if (row.names)
    jOffset <- 1

  if (append && file.exists(file)){
    wb <- loadWorkbook(file, password=password)
  } else {
    ext <- gsub(".*\\.(.*)$", "\\1", basename(file))
    wb  <- createWorkbook(type=ext)
  }
  sheet <- createSheet(wb, sheetName)

  noRows <- nrow(x) + iOffset
  noCols <- ncol(x) + jOffset
  if (col.names){
    rows  <- createRow(sheet, 1)                  # create top row
    cells <- createCell(rows, colIndex=1:noCols)  # create cells
    mapply(setCellValue, cells[1,(1+jOffset):noCols], colnames(x))
  }
  if (row.names)             # add rownames to data x
    x <- cbind(rownames=rownames(x), x)

  colIndex <- seq_len(ncol(x))
  rowIndex <- seq_len(nrow(x)) + iOffset

  .write_block(wb, sheet, x, rowIndex, colIndex, showNA)
  saveWorkbook(wb, file, password=password)

  invisible()
}


#  .jcall("java/lang/System", "V", "gc")  # doesn't do anything!


