#
#
#
#



.guess_cell_type <- function(cell)
{ 
  
}

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

  if (length(colClasses) < noColumns) {
    row <- getRows(sheet, rowIndex=startRow) 
    cells <- getCells(row, colIndex=startColumn:endColumn)
    cellType <- sapply(cells, function(x){x$getCellType()})
    # what happens with FORMULAS, ERROR or blanks? should default to "character"
    colClasses <- rep(colClasses, noColumns)
  }
  
  
  
  
  res <- vector("list", length=noColumns)
  for (i in seq_len(noColumns)) {
    
    res[[i]] <- switch(colClasses[i],
      numeric = .jcall(Rintf, "[D", "readColDoubles",
        .jcast(sheet, "org/apache/poi/ss/usermodel/Sheet"),
        as.integer(startRow-1), as.integer(noRows-1), 
        as.integer(startColumn-1+i-1)),
      character = .jcall(Rintf, "[S", "readColStrings",
        .jcast(sheet, "org/apache/poi/ss/usermodel/Sheet"),
        as.integer(startRow-1), as.integer(noRows-1), 
        as.integer(startColumn-1+i-1))                       
      )
  }
  
  if (as.data.frame){
    names(res) <- cnames
    res <- data.frame(res, ...)
  }

  res
}
