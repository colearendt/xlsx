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
write.xlsx <- function(x, file, sheetName="Sheet1",
  col.names=TRUE, row.names=TRUE, append=FALSE, showNA=TRUE)
{
  if (!is.data.frame(x))
    x <- data.frame(x)    # just because the error message is too ugly
    
  iOffset <- jOffset <- 0
  if (col.names)
    iOffset <- 1
  if (row.names)
    jOffset <- 1

  if (append && file.exists(file)){
    wb <- loadWorkbook(file)
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
  saveWorkbook(wb, file)

  invisible()
}


#  .jcall("java/lang/System", "V", "gc")  # doesn't do anything!


