# Write a data.frame to a new xlsx file. 
# A convenience function.
#

# System.gc()   # to call the garbage collection in Java

#######################################################################
# NO rownames for this function.  Just the contents of the data.frame!
#
.write_block <- function(wb, sheet, y, rowIndex=1:nrow(y), colIndex=1:ncol(y))
{
  rows  <- createRow(sheet, rowIndex)      # create rows 
  cells <- createCell(rows, colIndex)      # create cells
  
  for (ic in seq_len(ncol(y)))
    mapply(setCellValue, cells[1:nrow(cells), colIndex[ic]], y[,ic])

  # Date and POSIXct classes need to be formatted
  indDT <- which(sapply(y, class) == "Date")
  dateFormat <- createCellStyle(wb, dataFormat="m/d/yyyy")
  for (ic in indDT){
    lapply(cells[1:nrow(cells),colIndex[ic]], setCellStyle, dateFormat)
  }
  indDT <- which(sapply(y, class) == "POSIXct")
  datetimeFormat <- createCellStyle(wb, dataFormat="m/d/yyyy h:mm:ss;@")
  for (ic in indDT){
    lapply(cells[1:nrow(cells),colIndex[ic]], setCellStyle, datetimeFormat)
  }

}


write.xlsx <- function(x, file, sheetName="Sheet1", formatTemplate=NULL,
  col.names=TRUE, row.names=TRUE, append=FALSE)
{
  if (!is.data.frame(x))
    x <- data.frame(x)    # just because the error message is too ugly
    
  iOffset <- jOffset <- 0
  if (col.names)
    iOffset <- 1
  if (row.names)
    jOffset <- 1

  if (append){
    wb <- loadWorkbook(file)
  } else {
    wb <- createWorkbook()
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
  
  colIndex <- 1:ncol(x)
  rowIndex <- 1:nrow(x) + iOffset
  
  .write_block(wb, sheet, x, rowIndex, colIndex)
  saveWorkbook(wb, file)

  
##   # add a general format template
##   if (!is.null(formatTemplate)){
##     cat("formatTemplate is not yet supported.  Patches welcome.")
##     wb <- wb  # call the custom function
##   }
    
  
  invisible()
}


#  .jcall("java/lang/System", "V", "gc")  # doesn't do anything!

##  don't do it with blocks, it's silly
##   noBlocks <- trunc(nrow(x)/.BLOCK_SIZE)+1
##   for (b in seq_len(noBlocks)){
##     if (b > 1){
##       wb <- loadWorkbook(file)
##       sheet <- getSheets(wb)[[sheetName]]
##     }
##     indR <- ((b-1)*.BLOCK_SIZE+1):min(nrow(x), b*.BLOCK_SIZE)
##     rowIndex <- indR + jOffset
##     .write_block(sheet, x[indR,], rowIndex, colIndex)
##     saveWorkbook(wb, file)
##   }  
