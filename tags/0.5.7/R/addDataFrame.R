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
  rownamesStyle=NULL, showNA=FALSE, characterNA="", byrow=FALSE)
{
  if (!is.data.frame(x))
    x <- data.frame(x)    # just because the error message is too ugly

  if (row.names) {        # add rownames to data x as the first column
    x <- cbind(rownames=rownames(x), x)
    if (!is.null(colStyle))
      names(colStyle) <- as.numeric(names(colStyle)) + 1
  }
  
  wb <- sheet$getWorkbook()
  classes <- unlist(sapply(x, class))
  if ("Date" %in% classes) 
    csDate <- CellStyle(wb) + DataFormat(getOption("xlsx.date.format"))
  if ("POSIXct" %in% classes) 
    csDateTime <- CellStyle(wb) + DataFormat(getOption("xlsx.datetime.format"))

  # offset required to give space for column names
  # (either excel columns if byrow=TRUE or rows if byrow=FALSE)
  iOffset <- if (col.names) 1L else 0L
  # offset required to give space for row names
  # (either excel rows if byrow=TRUE or columns if byrow=FALSE)
  jOffset <- if (row.names) 1L else 0L

  if ( byrow ) {
      # write data.frame columns data row-wise
      setDataMethod   <- "setRowData"
      setHeaderMethod <- "setColData"
      blockRows <- ncol(x)
      blockCols <- nrow(x) + iOffset # row-wise data + column names
  } else {
      # write data.frame columns data column-wise, DEFAULT
      setDataMethod   <- "setColData"
      setHeaderMethod <- "setRowData"
      blockRows <- nrow(x) + iOffset # column-wise data + column names
      blockCols <- ncol(x)
  }

  # create a CellBlock, not sure why the usual .jnew doesn't work
  cellBlock <- CellBlock( sheet,
           as.integer(startRow), as.integer(startColumn),
           as.integer(blockRows), as.integer(blockCols),
           TRUE)

  # insert colnames
  if (col.names) {
    if (!(ncol(x) == 1 && colnames(x)=="rownames")) 
      .jcall( cellBlock$ref, "V", setHeaderMethod, 0L, jOffset,
         .jarray(colnames(x)[(1+jOffset):ncol(x)]), showNA,
         if ( !is.null(colnamesStyle) ) colnamesStyle$ref else
             .jnull('org/apache/poi/ss/usermodel/CellStyle') )
  }

  # write one data.frame column at a time, and style it if it has style
  # Dates and POSIXct columns get styled if not overridden. 
  for (j in seq_along(x)) {
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
#browser()
   .jcall( cellBlock$ref, "V", setDataMethod,
      as.integer(j-1L),   #  -1L for Java index 
      iOffset,            # does not need -1L
      .jarray(aux), showNA, 
      if ( !is.null(thisColStyle) ) thisColStyle$ref else
        .jnull('org/apache/poi/ss/usermodel/CellStyle') )
  }

  # return the cellBlock occupied by the generated data frame
  invisible(cellBlock)
}

