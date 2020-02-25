# One sheet extraction.  Similar to read.csv.
#
#
#' @export
#' @rdname read.xlsx
read.xlsx2 <- function(file, sheetIndex, sheetName=NULL, startRow=1,
  colIndex=NULL, endRow=NULL, as.data.frame=TRUE, header=TRUE,
  colClasses="character", password=NULL, ...)
{
  if (is.null(sheetName) & missing(sheetIndex))
    stop("Please provide a sheet name OR a sheet index.")

  wb <- loadWorkbook(file, password=password)
  sheets <- getSheets(wb)

  if (is.null(sheetName)){
    sheet <- sheets[[sheetIndex]]
  } else {
      sheet <- sheets[[sheetName]]
  }
  if (is.null(sheet))
      stop("Cannot find the sheet you requested in the file!")

  if (is.null(endRow)) {  # get it from the sheet
      endRow <- sheet$getLastRowNum() + 1
  }

  if (is.null(colIndex)){  # get it from the startRow
      row <- sheet$getRow(as.integer(startRow-1))
      if (is.null(row)) {
          if (endRow == startRow)
              return(NULL)          # sheet is empty
          stop(paste("Row with index startRow is EMPTY! ",
                     "Specify a different startRow value."))
      }
      startColumn  <- .jcall(row, "T", "getFirstCellNum") + 1
      endColumn    <- .jcall(row, "T", "getLastCellNum")
      colIndex     <- startColumn:endColumn
      listColIndex <- list(startColumn:endColumn)
  } else {
      listColIndex <- .splitBlocks(sort(colIndex))
  }

  # if the colIndex is not contiguous, split into contiguous blocks
  # and then cbind the blocks together.
  res <- NULL
  for (b in seq_along(listColIndex)) {
    startColumn <- listColIndex[[b]][1]
    endColumn <- tail(listColIndex[[b]],1)
    res[[b]] <- readColumns(sheet, startColumn, endColumn, startRow,
      endRow=endRow, as.data.frame=as.data.frame, header=header,
      colClasses=colClasses[which(colIndex %in% listColIndex[[b]])], ...)
  }

  do.call(cbind, res)
}


