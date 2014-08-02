# Functions to deal with ranges.
#
# createRange
# getRanges
# readRange
# 

#############################################################################
# Set an area of a sheet to a contiguous range
#
createRange <- function(rangeName, firstCell, lastCell)
{
  sheet <- firstCell$getSheet()
  sheetName <- sheet$getSheetName()
  firstCellRef <- .jnew("org/apache/poi/ss/util/CellReference",
    as.integer(firstCell$getRowIndex()), as.integer(firstCell$getColumnIndex()))
  lastCellRef <- .jnew("org/apache/poi/ss/util/CellReference",
    as.integer(lastCell$getRowIndex()), as.integer(lastCell$getColumnIndex()))
  
  nameFormula <- paste(sheetName, "!", firstCellRef$formatAsString(), ":",
    lastCellRef$formatAsString(), sep="")

  wb <- sheet$getWorkbook()
  range <- wb$createName()
  range$setNameName(rangeName)
  range$setRefersToFormula( nameFormula )

  range
}

#############################################################################
# get info about ranges in the spreadsheet, similar to getSheets
#
getRanges <- function(wb)
{
  noRanges <- .jcall(wb, "I", "getNumberOfNames")
  if (noRanges == 0){
    cat("Workbook has no ranges!\n")
    return()
  }

  res <- list()
  for (i in 1:noRanges) {
    aRange <- .jcall(wb, "Lorg/apache/poi/ss/usermodel/Name;", "getNameAt",
      as.integer(i-1))
    if (.jcall(aRange, "Z", "isDeleted")) {
      aRangeRef <- "No cell reference"
    } else {
      aRangeRef <- tryCatch(.jcall(aRange, "S", "getRefersToFormula"),
        error = function(e) "Invalid reference")
#      one <- list(ref=aRange, name=aRange$getNameName())
      res[[i]] <- aRange
#      names(res)[i] <- aRangeRef
    }
  }    
    
  res
}


#############################################################################
# Read one range.
#  range is a Java Object reference as returned by getRanges
#
readRange <- function(range, sheet, colClasses = "character")
{
  ## getRefersToFormula throws an error if external reference is broken
  aRangeRef <- tryCatch(.jcall(range, "S", "getRefersToFormula"),
    error = function(e) stop("Invalid range address", call. = FALSE))
  
  ## A broken reference starts with #REF!
  if (grepl("#REF!", aRangeRef, fixed=TRUE) > 0)
    stop(paste(range, aRangeRef, "is broken"), call. = FALSE)
  
  ## A linked range looks like '[path\path\file.xls]Sheet'!A1:B1
  if (grepl("[", aRangeRef, fixed=TRUE) > 0)
    stop(paste(range, aRangeRef, "linked range can't be read"), call. = FALSE)
  
  areaRef <- .jnew("org/apache/poi/ss/util/AreaReference", aRangeRef)
  if (!areaRef$isContiguous(aRangeRef))
    stop(paste(range, "is not contiguous.  Not supported yet!"), call. = FALSE)
 
  firstCell   <- areaRef$getFirstCell()
  startRow    <- firstCell$getRow()+1
  startColumn <- firstCell$getCol()+1

  lastCell  <- areaRef$getLastCell()
  endRow    <- lastCell$getRow()+1
  endColumn <- lastCell$getCol()+1

  noRows    <- endRow - startRow + 1
  noColumns <- endColumn - startColumn + 1

  if (length(colClasses) < noColumns)
    colClasses <- rep(colClasses, noColumns)
  
  Rintf <- .jnew("org/cran/rexcel/RInterface")  # create an interface object 
  res <- vector("list", length=noColumns)
  for (i in seq_len(noColumns)){
    res[[i]] <- switch(colClasses[i],
      numeric = .jcall(Rintf, "[D", "readColDoubles",
        .jcast(sheet, "org/apache/poi/ss/usermodel/Sheet"),
        as.integer(startRow-1), as.integer(startRow-1+noRows-1), 
        as.integer(startColumn-1+i-1)),
      character = .jcall(Rintf, "[S", "readColStrings",
        .jcast(sheet, "org/apache/poi/ss/usermodel/Sheet"),
        as.integer(startRow-1), as.integer(startRow-1+noRows-1), 
        as.integer(startColumn-1+i-1))                       
      )
  }
  
  names(res) <- paste("C", startColumn:endColumn, sep="")
  data.frame(res)
}




  ## firstCellRef <- .jnew("org/apache/poi/hssf/util/CellReference", firstCell)
  ## endCellRef   <- .jnew("org/apache/poi/hssf/util/CellReference", endCell)
  ## areaRef <- .jnew("org/apache/poi/ss/util/AreaReference", firstCellRef, endCellRef)
  ## nameFormula <- areaRef$formatAsString()


