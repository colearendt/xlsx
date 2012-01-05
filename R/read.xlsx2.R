# One sheet extraction.  Similar to read.csv.
#
# TODO:  what's the best colClasses default?
# 
read.xlsx2 <- function(file, sheetIndex, sheetName=NULL, startRow=1,
  colIndex=NULL, noRows=NULL, as.data.frame=TRUE,
  header=TRUE, colClasses="character", ...)
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

  if (is.null(colIndex)){  # get it from the startRow
    row <- sheet$getRow(as.integer(startRow-1))
    startColumn <- .jcall(row, "T", "getFirstCellNum")   # 0-based
    endColumn   <- .jcall(row, "T", "getLastCellNum")-1  # 0-based
  }

  if (is.null(noRows)) {  # get it from the sheet 
    endRow <- sheet$getLastRowNum() + 1
  } else {
    endRow <- startRow + noRows
  }
 
  readColumns(sheet, startColumn, endColumn, startRow, endRow=endRow,
    as.data.frame=as.data.frame, header=header, colClasses=colClasses,
    ...)

}

  
