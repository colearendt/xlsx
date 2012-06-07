# Add a data.frame to a sheet
#
# colStyle can be a structure of CellStyle with names representing column
#  index.  
#
# I shouldn't offer to change the data for the user, with characterNA="",
# they should do it themselves.  It's not that hard!
#
addDataFrame <- function(x, sheet, col.names=TRUE, row.names=TRUE,
  startRow=1, startColumn=1, colStyle=NULL, colnamesStyle=NULL,
  rownamesStyle=NULL, showNA=FALSE, characterNA="")
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


  # create a CellBlock, not sure why the usual .jnew doesn't work 
  cellBlock <- CellBlock(sheet, indX1, indY1,
    as.integer(nrow(x) + iOffset), as.integer(ncol(x)), TRUE)

  # insert colnames
  if (col.names) {                   
    .jcall( cellBlock$ref, "V", "setRowData", 0L, jOffset,
       .jarray(if (row.names) names(x)[-1] else names(x)), showNA,
       if ( !is.null(colnamesStyle) ) colnamesStyle$ref else
           .jnull('org/apache/poi/ss/usermodel/CellStyle') )
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
    }
   .jcall( cellBlock$ref, "V", "setColData", as.integer(j+jOffset-1L), iOffset,
     .jarray(aux), showNA, if ( !is.null(colStyle) ) colStyle$ref else
        .jnull('org/apache/poi/ss/usermodel/CellStyle') )
  }

  return ( invisible(cellBlock) )
}

