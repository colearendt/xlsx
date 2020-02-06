# One sheet extraction.  Similar to read.csv.
#
#
#' Read the contents of a worksheet into an R \code{data.frame}.
#'
#'
#' The \code{read.xlsx} function provides a high level API for reading data
#' from an Excel worksheet.  It calls several low level functions in the
#' process.  Its goal is to provide the conveniency of
#' \code{\link[utils]{read.table}} by borrowing from its signature.
#'
#' The function pulls the value of each non empty cell in the worksheet into a
#' vector of type \code{list} by preserving the data type.  If
#' \code{as.data.frame=TRUE}, this vector of lists is then formatted into a
#' rectangular shape.  Special care is needed for worksheets with ragged data.
#'
#' An attempt is made to guess the class type of the variable corresponding to
#' each column in the worksheet from the type of the first non empty cell in
#' that column.  If you need to impose a specific class type on a variable, use
#' the \code{colClasses} argument.  It is recommended to specify the column
#' classes and not rely on \code{R} to guess them, unless in very simple cases.
#'
#' Excel internally stores dates and datetimes as numeric values, and does not
#' keep track of time zones and DST.  When a datetime column is brought into ,
#' it is converted to \code{POSIXct} class with a \emph{GMT} timezone.
#' Occasional rounding errors may appear and the and Excel string
#' representation my differ by one second.  For \code{read.xlsx2} bring in a
#' datetime column as a numeric one and then convert to class \code{POSIXct} or
#' \code{Date}.  Also rounding the \code{POSIXct} column in R usually does the
#' trick too.
#'
#' The \code{read.xlsx2} function does more work in Java so it achieves better
#' performance (an order of magnitude faster on sheets with 100,000 cells or
#' more).  The result of \code{read.xlsx2} will in general be different from
#' \code{read.xlsx}, because internally \code{read.xlsx2} uses
#' \code{readColumns} which is tailored for tabular data.
#'
#' Reading of password protected workbooks is supported for Excel 2007 OOXML
#' format only.
#'
#' @param file the path to the file to read.
#' @param sheetIndex a number representing the sheet index in the workbook.
#' @param sheetName a character string with the sheet name.
#' @param rowIndex a numeric vector indicating the rows you want to extract.
#' If \code{NULL}, all rows found will be extracted, unless \code{startRow} or
#' \code{endRow} are specified.
#' @param colIndex a numeric vector indicating the cols you want to extract.
#' If \code{NULL}, all columns found will be extracted.
#' @param as.data.frame a logical value indicating if the result should be
#' coerced into a \code{data.frame}.  If \code{FALSE}, the result is a list
#' with one element for each column.
#' @param header a logical value indicating whether the first row corresponding
#' to the first element of the \code{rowIndex} vector contains the names of the
#' variables.
#' @param colClasses For \code{read.xlsx} a character vector that represent the
#' class of each column.  Recycled as necessary, or if the character vector is
#' named, unspecified values are taken to be \code{NA}.  For \code{read.xlsx2}
#' see \code{\link{readColumns}}.
#' @param keepFormulas a logical value indicating if Excel formulas should be
#' shown as text in and not evaluated before bringing them in.
#' @param encoding encoding to be assumed for input strings.  See
#' \code{\link[utils]{read.table}}.
#' @param startRow a number specifying the index of starting row.  For
#' \code{read.xlsx} this argument is active only if \code{rowIndex} is
#' \code{NULL}.
#' @param endRow a number specifying the index of the last row to pull.  If
#' \code{NULL}, read all the rows in the sheet.  For \code{read.xlsx} this
#' argument is active only if \code{rowIndex} is \code{NULL}.
#' @param password a String with the password.
#' @param ... other arguments to \code{data.frame}, for example
#' \code{stringsAsFactors}
#' @return A data.frame or a list, depending on the \code{as.data.frame}
#' argument.  If some of the columns are read as NA's it's an indication that
#' the \code{colClasses} argument has not been set properly.
#'
#' If the sheet is empty, return \code{NULL}.  If the sheet does not exist,
#' return an error.
#' @author Adrian Dragulescu
#' @seealso \code{\link{write.xlsx}} for writing \code{xlsx} documents.  See
#' also \code{\link{readColumns}} for reading only a set of columns into R.
#' @examples
#'
#' \dontrun{
#'
#' file <- system.file("tests", "test_import.xlsx", package = "xlsx")
#' res <- read.xlsx(file, 1)  # read first sheet
#' head(res)
#' #          NA. Population Income Illiteracy Life.Exp Murder HS.Grad Frost   Area
#' # 1    Alabama       3615   3624        2.1    69.05   15.1    41.3    20  50708
#' # 2     Alaska        365   6315        1.5    69.31   11.3    66.7   152 566432
#' # 3    Arizona       2212   4530        1.8    70.55    7.8    58.1    15 113417
#' # 4   Arkansas       2110   3378        1.9    70.66   10.1    39.9    65  51945
#' # 5 California      21198   5114        1.1    71.71   10.3    62.6    20 156361
#' # 6   Colorado       2541   4884        0.7    72.06    6.8    63.9   166 103766
#' # >
#'
#'
#' # To convert an Excel datetime colum to POSIXct, do something like:
#' #   as.POSIXct((x-25569)*86400, tz="GMT", origin="1970-01-01")
#' # For Dates, use a conversion like:
#' #   as.Date(x-25569, origin="1970-01-01")
#'
#' res2 <- read.xlsx2(file, 1)
#'
#' }
#'
#' @export
read.xlsx <- function(file, sheetIndex, sheetName=NULL,
  rowIndex=NULL, startRow=NULL, endRow=NULL, colIndex=NULL,
  as.data.frame=TRUE, header=TRUE, colClasses=NA,
  keepFormulas=FALSE, encoding="unknown", password=NULL, ...)
{
  if (is.null(sheetName) & missing(sheetIndex))
    stop("Please provide a sheet name OR a sheet index.")

  wb <- loadWorkbook(file, password=password)
  sheets <- getSheets(wb)
  sheet  <- if (is.null(sheetName)) {
    sheets[[sheetIndex]]
  } else {
    sheets[[sheetName]]
  }

  if (is.null(sheet))
    stop("Cannot find the sheet you requested in the file!")

  rowIndex <- if (is.null(rowIndex)) {
    if (is.null(startRow))
      startRow <- .jcall(sheet, "I", "getFirstRowNum") + 1
    if (is.null(endRow))
      endRow <- .jcall(sheet, "I", "getLastRowNum") + 1
    startRow:endRow
  } else rowIndex

  rows  <- getRows(sheet, rowIndex)
  if (length(rows)==0)
      return(NULL)             # exit early

  cells <- getCells(rows, colIndex)
  res <- lapply(cells, getCellValue, keepFormulas=keepFormulas,
                encoding=encoding)

  if (as.data.frame) {
    # need to use the index from the names because of empty cells
    ind <- lapply(strsplit(names(res), "\\."), as.numeric)
    namesIndM <- do.call(rbind, ind)

    row.names <- sort(as.numeric(unique(namesIndM[,1])))
    col.names <- paste("V", sort(unique(namesIndM[,2])), sep="")
    col.names <- sort(unique(namesIndM[,2]))
    cols <- length(col.names)

    VV <- matrix(list(NA), nrow=length(row.names), ncol=cols,
      dimnames=list(row.names, col.names))
    # you need indM for empty rows/columns when indM != namesIndM
    indM <- apply(namesIndM, 2, function(x){as.numeric(as.factor(x))})
    VV[indM] <- res

    if (header){  # first row of cells that you want
      colnames(VV) <- VV[1,]
      VV <- VV[-1,,drop=FALSE]
    }

    res <- vector("list", length=cols)
    names(res) <- colnames(VV)
    for (ic in seq_len(cols)) {
      aux   <- unlist(VV[,ic], use.names=FALSE)
      nonNA <- which(!is.na(aux))
      if (length(nonNA)>0) {  # not a NA column in the middle of data
        ind <- min(nonNA)
        if (class(aux[ind])=="numeric") {
          # test first not NA cell if it's a date/datetime
          dateUtil <- .jnew("org/apache/poi/ss/usermodel/DateUtil")
          cell <- cells[[paste(row.names[ind + header], ".", col.names[ic], sep = "")]]
          isDatetime <- dateUtil$isCellDateFormatted(cell)
          if (isDatetime){
            if (identical(aux, round(aux))){ # you have dates
              aux <- as.Date(aux-25569, origin="1970-01-01")
            } else {  # Excel does not know timezones?!
              aux <- as.POSIXct((aux-25569)*86400, tz="GMT",
                                origin="1970-01-01")
            }
          }
        }
      }
      if (!is.na(colClasses[ic]))
        suppressWarnings(class(aux) <- colClasses[ic])  # if it gets specified

      res[[ic]] <- aux
    }

    res <- data.frame(res, ...)
  }

  res
}



