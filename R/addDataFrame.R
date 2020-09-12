#' Add a \code{data.frame} to a sheet.
#'
#' Add a \code{data.frame} to a sheet, allowing for different column styles.
#' Useful when constructing the spreadsheet from scratch.
#'
#' Starting with version 0.5.0 this function uses the functionality provided by
#' \code{CellBlock} which results in a significant improvement in performance
#' compared with a cell by cell application of \code{\link{setCellValue}} and
#' with other previous atempts.
#'
#' It is difficult to treat \code{NA}'s consistently between R and Excel via
#' Java.  Most likely, users of Excel will want to see \code{NA}'s as blank
#' cells.  In R character \code{NA}'s are simply characters, which for Excel
#' means "NA".
#'
#' The default formats for Date and DateTime columns can be changed via the two
#' package options \code{xlsx.date.format} and \code{xlsx.datetime.format}.
#' They need to be specified in Java date format
#' \url{https://docs.oracle.com/javase/7/docs/api/java/text/SimpleDateFormat.html}.
#'
#' @param x a \code{data.frame}.
#' @param sheet a \code{\link{Sheet}} object.
#' @param col.names a logical value indicating if the column names of \code{x}
#' are to be written along with \code{x} to the file.
#' @param row.names a logical value indicating whether the row names of
#' \code{x} are to be written along with \code{x} to the file.
#' @param startRow a numeric value for the starting row.
#' @param startColumn a numeric value for the starting column.
#' @param colStyle a list of \code{\link{CellStyle}}.  If the name of the list
#' element is the column number, it will be used to set the style of the
#' column.  Columns of type \code{Date} and \code{POSIXct} are styled
#' automatically even if \code{colSyle=NULL}.
#' @param colnamesStyle a \code{\link{CellStyle}} object to customize the table
#' header.
#' @param rownamesStyle a \code{\link{CellStyle}} object to customize the row
#' names (if \code{row.names=TRUE}).
#' @param showNA a boolean value to control how NA's are displayed on the
#' sheet.  If \code{FALSE}, NA values will be represented as blank cells.
#' @param characterNA a string value to control how character NA will be shown
#' in the spreadsheet.
#' @param byrow a logical value indicating if the data.frame should be added to
#' the sheet in row wise fashion.
#' @return None.  The modification to the workbook is done in place.
#' @author Adrian Dragulescu
#' @examples
#'
#'
#'   wb <- createWorkbook()
#'   sheet  <- createSheet(wb, sheetName="addDataFrame1")
#'   data <- data.frame(mon=month.abb[1:10], day=1:10, year=2000:2009,
#'     date=seq(as.Date("1999-01-01"), by="1 year", length.out=10),
#'     bool=c(TRUE, FALSE), log=log(1:10),
#'     rnorm=10000*rnorm(10),
#'     datetime=seq(as.POSIXct("2011-11-06 00:00:00", tz="GMT"), by="1 hour",
#'       length.out=10))
#'   cs1 <- CellStyle(wb) + Font(wb, isItalic=TRUE)           # rowcolumns
#'   cs2 <- CellStyle(wb) + Font(wb, color="blue")
#'   cs3 <- CellStyle(wb) + Font(wb, isBold=TRUE) + Border()  # header
#'   addDataFrame(data, sheet, startRow=3, startColumn=2, colnamesStyle=cs3,
#'     rownamesStyle=cs1, colStyle=list(`2`=cs2, `3`=cs2))
#'
#'   # to change the default date format use something like this
#'   # options(xlsx.date.format="dd MMM, yyyy")
#'
#'
#'   # Don't forget to save the workbook ...
#'   # saveWorkbook(wb, file)
#'
#' @export
addDataFrame <- function(x, sheet, col.names=TRUE, row.names=TRUE,
  startRow=1, startColumn=1, colStyle=NULL, colnamesStyle=NULL,
  rownamesStyle=NULL, showNA=FALSE, characterNA="", byrow=FALSE)
{
  if (!is.data.frame(x))
    x <- data.frame(x)    # just because the error message is too ugly

  if (row.names) {        # add rownames to data x as the first column
    x <- cbind(rownames=rownames(x), x)
    if (!is.null(colStyle))
      names(colStyle) <- as.numeric(names(colStyle)) + 1
  }

  wb <- sheet$getWorkbook()
  classes <- unlist(sapply(x, class))
  if ("Date" %in% classes)
    csDate <- CellStyle(wb) + DataFormat(getOption("xlsx.date.format"))
  if ("POSIXct" %in% classes)
    csDateTime <- CellStyle(wb) + DataFormat(getOption("xlsx.datetime.format"))

  # offset required to give space for column names
  # (either excel columns if byrow=TRUE or rows if byrow=FALSE)
  iOffset <- if (col.names) 1L else 0L
  # offset required to give space for row names
  # (either excel rows if byrow=TRUE or columns if byrow=FALSE)
  jOffset <- if (row.names) 1L else 0L

  if ( byrow ) {
      # write data.frame columns data row-wise
      setDataMethod   <- "setRowData"
      setHeaderMethod <- "setColData"
      blockRows <- ncol(x)
      blockCols <- nrow(x) + iOffset # row-wise data + column names
  } else {
      # write data.frame columns data column-wise, DEFAULT
      setDataMethod   <- "setColData"
      setHeaderMethod <- "setRowData"
      blockRows <- nrow(x) + iOffset # column-wise data + column names
      blockCols <- ncol(x)
  }

  # create a CellBlock, not sure why the usual .jnew doesn't work
  cellBlock <- CellBlock( sheet,
           as.integer(startRow), as.integer(startColumn),
           as.integer(blockRows), as.integer(blockCols),
           TRUE)

  # insert colnames
  if (col.names) {
    if (!(ncol(x) == 1 && colnames(x)=="rownames"))
      .jcall( cellBlock$ref, "V", setHeaderMethod, 0L, jOffset,
         .jarray(colnames(x)[(1+jOffset):ncol(x)]), showNA,
         if ( !is.null(colnamesStyle) ) colnamesStyle$ref else
             .jnull('org/apache/poi/ss/usermodel/CellStyle') )
  }

  # write one data.frame column at a time, and style it if it has style
  # Dates and POSIXct columns get styled if not overridden.
  for (j in seq_along(x)) {
    thisColStyle <-
      if ((j==1) && (row.names) && (!is.null(rownamesStyle))) {
        rownamesStyle
      } else if (as.character(j) %in% names(colStyle)) {
        colStyle[[as.character(j)]]
      } else if ("Date" %in% class(x[,j])) {
        csDate
      } else if ("POSIXt" %in% class(x[,j])) {
        csDateTime
      } else {
        NULL
      }

    xj <- x[,j]
    if ("integer" %in% class(xj)) {
      aux <- xj
    } else if (any(c("numeric", "Date", "POSIXt") %in% class(xj))) {
      aux <- if ("Date" %in% class(xj)) {
          as.numeric(xj)+25569
        } else if ("POSIXt" %in% class(x[,j])) {
          as.numeric(xj)/86400 + 25569
        } else {
          xj
        }
      haveNA <- is.na(aux)
      if (any(haveNA))
        aux[haveNA] <- NaN          # encode the numeric NAs as NaN for java
    } else {
      aux <- as.character(x[,j])
      haveNA <- is.na(aux)
      if (any(haveNA))
        aux[haveNA] <- characterNA

      # Excel max cell size limit
      if (!all(is.na(aux)) && max(nchar(aux), na.rm = TRUE) > .EXCEL_LIMIT_MAX_CHARS_IN_CELL) {
          warning(sprintf("Some cells exceed Excel's limit of %d characters and they will be truncated",
                          .EXCEL_LIMIT_MAX_CHARS_IN_CELL))
          aux <- strtrim(aux, .EXCEL_LIMIT_MAX_CHARS_IN_CELL)
      }
    }

   .jcall( cellBlock$ref, "V", setDataMethod,
      as.integer(j-1L),   #  -1L for Java index
      iOffset,            # does not need -1L
      .jarray(aux), showNA,
      if ( !is.null(thisColStyle) ) thisColStyle$ref else
        .jnull('org/apache/poi/ss/usermodel/CellStyle') )
  }

  # return the cellBlock occupied by the generated data frame
  invisible(cellBlock)
}

