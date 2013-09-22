# 
#
# For CellType: FORMULA, ERROR or blanks, it defaults to "character"
#
#

readColumns <- function(sheet, startColumn, endColumn, startRow,
  endRow=NULL, as.data.frame=TRUE, header=TRUE, colClasses=NA, ...)
{
  trueEndRow <- sheet$getLastRowNum() + 1
  if (is.null(endRow))     # get it from the sheet 
    endRow <- trueEndRow
  
  if (endRow > trueEndRow) {
    warning(paste("This sheet has only", trueEndRow, "rows"))
    endRow <- min(endRow, trueEndRow)
  }
  
  noColumns <- endColumn - startColumn + 1

  Rintf <- .jnew("org/cran/rexcel/RInterface")  # create an interface object 

  if (header) {
    cnames <- try(.jcall(Rintf, "[S", "readRowStrings",
      .jcast(sheet, "org/apache/poi/ss/usermodel/Sheet"),
      as.integer(startColumn-1), as.integer(endColumn-1),
      as.integer(startRow-1)))
    if (class(cnames) == "try-error") 
      stop(paste("Cannot read the header. ",
        "The header row doesn't have all requested cells non-blank!"))
    startRow <- startRow + 1
  } else {
    cnames <- as.character(1:noColumns)
  }

  # guess or expand colClasses
  if (is.na(colClasses)[1]) {
    row <- getRows(sheet, rowIndex=startRow) 
    cells <- getCells(row, colIndex=startColumn:endColumn)
    if (length(cells) != noColumns) 
      warning("Not enough columns in the first row of data to correctly guess colClasses!")
    colClasses <- .guess_cell_type(cells)
  }
  colClasses <- rep(colClasses, length.out=noColumns)  # extend colClasses
  
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
#browser()
    if (!is.na(colClasses[i]))
      suppressWarnings(class(aux) <- colClasses[i])  # if it gets specified
    res[[i]] <- aux
  }
  
  if (as.data.frame) {
    cnames[cnames == ""] <- " " # remove some silly c.......... colnames!
    names(res) <- cnames  
    res <- data.frame(res, ...)
  }

  res
}
