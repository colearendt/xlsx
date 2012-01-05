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
    aux <- .jarray(names(x)[-1])
    if (!is.null(colnamesStyle)) {   
      .jcall(Rintf, "V", "writeRowStrings", sheet, 0L, 1L, aux,
             colnamesStyle$ref) 
    } else {
      .jcall(Rintf, "V", "writeRowStrings", sheet, 0L, 1L, aux)
    }
  }
  # insert one column at a time, and style it if it has style
  # Dates and POSIXct columns get styled if not overridden. 
  for (j in 1:ncol(x)) {
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
#browser()
    if ("integer" %in% class(x[,j])) {
      if (is.null(thisColStyle)) {
        .jcall(Rintf, "V", "writeColInts", sheet, iOffset, as.integer(j-1),
          .jarray(x[,j]))
      } else {
        .jcall(Rintf, "V", "writeColInts", sheet, iOffset, as.integer(j-1),
          .jarray(x[,j]), thisColStyle$ref)
      }
      
    } else if (any(c("numeric", "Date", "POSIXt") %in% class(x[,j]))) {
      aux <- if ("Date" %in% class(x[,j])) {
          as.numeric(x[,j])+25569
        } else if ("POSIXt" %in% class(x[,j])) {
          as.numeric(x[,j])/86400 + 25569
        } else {
          x[,j]
        } 
      if (is.null(thisColStyle)) {
        .jcall(Rintf, "V", "writeColDoubles", sheet, iOffset, as.integer(j-1),
           .jarray(aux))
      } else {
        .jcall(Rintf, "V", "writeColDoubles", sheet, iOffset, as.integer(j-1),
           .jarray(aux), thisColStyle$ref)
      }
      
    } else {
      if (is.null(thisColStyle)) {
        .jcall(Rintf, "V", "writeColStrings", sheet, iOffset, as.integer(j-1),
           .jarray(as.character(x[,j])))
      } else {
        .jcall(Rintf, "V", "writeColStrings", sheet, iOffset, as.integer(j-1),
           .jarray(as.character(x[,j])), thisColStyle$ref)
      }      
    }
  }
  
  invisible()
}
