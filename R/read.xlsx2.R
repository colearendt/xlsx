# One sheet extraction.  Similar to read.csv.
#
# TODO:  what's the best colClasses default?
# 
read.xlsx2 <- function(file, sheetIndex, sheetName=NULL, startRow=1,
  colIndex=NULL, endRow=NULL, as.data.frame=TRUE, header=TRUE,
  colClasses="character", ...)
{
  if (is.null(sheetName) & missing(sheetIndex))
    stop("Please provide a sheet name OR a sheet index.")

  wb <- loadWorkbook(file)
  sheets <- getSheets(wb)

  if (is.null(sheetName)){
    sheet <- sheets[[sheetIndex]]
  } else {
    sheet <- sheets[[sheetName]]
  }

  if (is.null(endRow)) {  # get it from the sheet 
    endRow <- sheet$getLastRowNum() + 1
  }
  
  if (is.null(colIndex)){  # get it from the startRow
    row <- sheet$getRow(as.integer(startRow-1))
    if (is.null(row)) 
      stop(paste("Row with index startRow is EMPTY! ",
           "Specify a different startRow value."))
    startColumn <- .jcall(row, "T", "getFirstCellNum") + 1   
    endColumn   <- .jcall(row, "T", "getLastCellNum") + 1    
    colIndex <- list(startColumn:endColumn)
  } else {
    colIndex <- xlsx:::.splitBlocks(sort(colIndex))
  }

  # if the colIndex is not contiguous, split into contiguous blocks
  # and then cbind the blocks together. 
  res <- NULL
  for (b in seq_along(colIndex)) {
    startColumn <- colIndex[[b]][1]
    endColumn <- tail(colIndex[[b]],1)
    res[[b]] <- readColumns(sheet, startColumn, endColumn, startRow,
      endRow=endRow, as.data.frame=as.data.frame, header=header,
      colClasses=colClasses, ...)
  }
  
  do.call(cbind, res)
}

  
