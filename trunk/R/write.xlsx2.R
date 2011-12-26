# Write a data.frame to a new xlsx file. 
# with a java back-end
#

write.xlsx2 <- function(x, file, sheetName="Sheet 1", formatTemplate=NULL,
  col.names=TRUE, row.names=TRUE, append=FALSE)
{
  if (!is.data.frame(x))
    x <- data.frame(x)    # just because the error message is too ugly
    
  iOffset <- jOffset <- 0L
  if (col.names)
    iOffset <- 1L

  if (append){
    wb <- loadWorkbook(file)
  } else {
    wb <- createWorkbook()
  }  
  sheet <- createSheet(wb, sheetName)

  if (row.names)             # add rownames to data x                   
    x <- cbind(rownames=rownames(x), x)

  # create a new interface object 
  Rintf <- .jnew("dev/RInterface")
  Rintf$NCOLS <- ncol(x)             # set the number of columns
  Rintf$NROWS <- nrow(x) + iOffset   # set the number of rows
  # create the cells in Rint$CELL_ARRAY
  .jcall(Rintf, "V", "createCells", sheet, 0L, 0L)

  if (col.names){                    # insert colnames
    aux <- .jarray(names(x))
    .jcall(Rintf, "V", "writeRowStrings", sheet, 0L, 0L, aux)
  }

  # insert one column at a time
  for (j in 1:ncol(x)){
    if (class(x[,j]) == "integer") {
      aux <- .jarray(x[,j])
      .jcall(Rintf, "V", "writeColInts", sheet, iOffset, as.integer(j-1), aux)
    } else if (class(x[,j]) == "numeric") {
      aux <- .jarray(x[,j])
      .jcall(Rintf, "V", "writeColDoubles", sheet, iOffset, as.integer(j-1), aux)
    } else {
      aux <- .jarray(as.character(x[,j]))
      .jcall(Rintf, "V", "writeColStrings", sheet, iOffset, as.integer(j-1), aux)
    }
  }

  fh <- .jnew("java/io/FileOutputStream", file)  # open the file
  wb$write(fh)                                   # write the spreadsheet
  .jcall(fh, "V", "close")                       # close the file
  
  invisible()
}


