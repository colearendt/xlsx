# One sheet extraction.  Similar to read.csv. 
#
#
read.xlsx2 <- function(file, sheetIndex, sheetName=NULL, startRow=1,
  startColumn=1, noRows=NULL, noColumns=NULL, as.data.frame=TRUE,
  header=TRUE, colClasses="character", ...)
{
  if (is.null(sheetName) & missing(sheetIndex))
    stop("Please provide a sheet name OR a sheet index.")

  wb <- loadWorkbook(file)
  sheets <- getSheets(wb)

  if (is.null(sheetName)){
    sheet <- sheets[[sheetIndex]]
  } else {
    sheet <- sheets[[sheetName]]
  }

  if (is.null(noRows)){  # get it from the sheet 
    noRows <- sheet$getLastRowNum() + 1 
  }
  
  if (is.null(noColumns)){  # get it from the startRow
    row <- sheet$getRow(as.integer(startRow-1))
    noColumns <- row$getLastCellNum()
  }

  if (length(colClasses) < noColumns)
    colClasses <- rep(colClasses, noColumns)
  
  Rintf <- .jnew("dev/RInterface")  # create an interface object 
  
  if (header){
    cnames <- .jcall(Rintf, "[S", "readRowStrings",
      .jcast(sheet, "org/apache/poi/ss/usermodel/Sheet"),
      as.integer(startColumn-1), as.integer(startColumn-1+noColumns-1),
      as.integer(startRow-1))
    startRow <- startRow + 1
    noRows <- noRows - 1 
  } else {
    cnames <- as.character(1:noColumns)
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

  
