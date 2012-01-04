# Add a data.frame to a sheet
#
# colStyle can be a structure of CellStyle with names representing column
#  index.  
#
addDataFrame <- function(x, sheet, col.names=TRUE, row.names=TRUE,
  startRow=1, startColumn=1, colStyle=NULL, colnamesStyle=NULL,
  rownamesStyle=NULL)
{
  if (!is.data.frame(x))
    x <- data.frame(x)    # just because the error message is too ugly
  
  if (row.names) {        # add rownames to data x                   
    x <- cbind(rownames=rownames(x), x)
    if (!is.null(colStyle))
      names(colStyle) <- as.numeric(names(colStyle)) + 1
  }
  
  wb <- sheet$getWorkbook()
  classes <- unlist(sapply(x, class))
  if ("Date" %in% classes) 
   csDate <- CellStyle(wb) + DataFormat("m/d/yyyy")
  if ("POSIXct" %in% classes) 
   csDateTime <- CellStyle(wb) + DataFormat("m/d/yyyy h:mm:ss;@")

  iOffset <- if (col.names) 1L else 0L
  jOffset <- if (row.names) 1L else 0L
  indX1   <- as.integer(startRow-1)        # index of top row
  indY1   <- as.integer(startColumn-1)     # index of top column
  
  # create a new interface object 
  Rintf <- .jnew("dev/RInterface")
  Rintf$NCOLS <- ncol(x) + jOffset   # set the number of columns
  Rintf$NROWS <- nrow(x) + iOffset   # set the number of rows
  # create the cells in Rint$CELL_ARRAY
  .jcall(Rintf, "V", "createCells", sheet, indX1, indY1)
  
  if (col.names) {                   # insert colnames
    aux <- .jarray(names(x))
    if (!is.null(colnamesStyle)) {   
      .jcall(Rintf, "V", "writeRowStrings", sheet, indX1, indY1, aux,
             colnamesStyle)
    } else {
      .jcall(Rintf, "V", "writeRowStrings", sheet, indX1, indY1, aux)
    }
  }

  # insert one column at a time, and style it if it has style
  # Dates and POSIXct columns get styled if not overridden. 
  for (j in 1:ncol(x)){
    thisColStyle <-
      if ((j==1) && (row.names) && (!is.null(rownamesStyle))) {
        rownamesStyle
      } else if (class) {

      
    } else {
      NULL
    }
    
    if (class(x[,j]) == "integer") {
      aux <- .jarray(x[,j])
      .jcall(Rintf, "V", "writeColInts", sheet, iOffset, as.integer(j-1), aux)
    } else if (class(x[,j]) == "numeric") {
      aux <- .jarray(x[,j])
      .jcall(Rintf, "V", "writeColDoubles", sheet, iOffset, as.integer(j-1), aux)
    } else {
      aux <- .jarray(as.character(x[,j]))
      if ((j==1) && (row.names) && (!is.null(rownamesStyle))) {
        .jcall(Rintf, "V", "writeColStrings", sheet, iOffset, as.integer(j-1),
           aux, rownamesStyle)
      } else {
        .jcall(Rintf, "V", "writeColStrings", sheet, iOffset, as.integer(j-1), aux)
      }
      
    }
  }

  
  invisible()
}
