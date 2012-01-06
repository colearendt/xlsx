# 
#
# For CellType: FORMULA, ERROR or blanks, it defaults to "character"
#
#

readColumns <- function(sheet, startColumn, endColumn, startRow,
  endRow=NULL, as.data.frame=TRUE, header=TRUE, colClasses=NA, ...)
{

  if (is.null(endRow))    # get it from the sheet 
    endRow <- sheet$getLastRowNum() + 1

  noColumns <- endColumn - startColumn

  Rintf <- .jnew("dev/RInterface")  # create an interface object 
  
  if (header){
    cnames <- .jcall(Rintf, "[S", "readRowStrings",
      .jcast(sheet, "org/apache/poi/ss/usermodel/Sheet"),
      as.integer(startColumn-1), as.integer(endColumn-1),
      as.integer(startRow-1))
    startRow <- startRow + 1
  } else {
    cnames <- as.character(1:noColumns)
  }
  
  # guess or expand colClasses
  if (length(colClasses) < noColumns) {
    colClasses <- rep(colClasses, noColumns)   
  } else {
    row <- getRows(sheet, rowIndex=startRow) 
    cells <- getCells(row, colIndex=startColumn:endColumn)
    colClasses <- .guess_cell_type(cells)
  }
  
  res <- vector("list", length=noColumns)
  for (i in seq_len(noColumns)) {
    if (any(c("numeric", "POSIXct", "Date") %in% colClasses[i])) {
      aux <- .jcall(Rintf, "[D", "readColDoubles",
        .jcast(sheet, "org/apache/poi/ss/usermodel/Sheet"),
        as.integer(startRow-1), as.integer(endRow-1), 
        as.integer(startColumn-1+i-1))
      if (colClasses[i]=="Date") 
        aux <- as.Date(aux-25569, origin="1970-01-01")
      if (colClasses[i]=="POSIXct")
        aux <- as.POSIXct((aux-25569)*86400, tz="GMT", origin="1970-01-01")
    } else {
      aux <- .jcall(Rintf, "[S", "readColStrings",
        .jcast(sheet, "org/apache/poi/ss/usermodel/Sheet"),
        as.integer(startRow-1), as.integer(endRow-1), 
        as.integer(startColumn-1+i-1))
      }
    if (!is.na(colClasses[ic]))
      suppressWarnings(class(aux) <- colClasses[ic])  # if it gets specified
    res[[ic]] <- aux
  }
  
  if (as.data.frame){
    names(res) <- cnames
    res <- data.frame(res, ...)
  }

  res
}
