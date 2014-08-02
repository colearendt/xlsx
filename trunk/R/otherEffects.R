# addAutoFilter
# addHyperlink
# addMergedRegion
# autoSizeColumn
# createFreezePane
# createSplitPane
# removeMergedRegion
# setColumnWidth
# setPrintArea
# setZoom
#
#

######################################################################
# Add autofilter to a cell range 
# @param sheet a Sheet object
# @param cellRange a string, "C5:F200", or one row "C1:F1",
#   a column range "C:F", a singe cell "B5", or a row range "3:5"
#
addAutoFilter <- function(sheet, cellRange)
{
  cellRangeAddress <- J("org/apache/poi/ss/util/CellRangeAddress")$valueOf(cellRange) 
  invisible(sheet$setAutoFilter(cellRangeAddress))     
}


######################################################################
# Add hyperlink
#
addHyperlink <- function(cell, address, linkType=c("URL", "DOCUMENT",
  "EMAIL", "FILE"), hyperlinkStyle=NULL)
{
  linkType <- match.arg(linkType)
  
  sh <- .jcall(cell, "Lorg/apache/poi/ss/usermodel/Sheet;", "getSheet")
  wb <- .jcall(sh, "Lorg/apache/poi/ss/usermodel/Workbook;", "getWorkbook")

  if (is.null(hyperlinkStyle)) {
    # create a cell style for hyperlinks
    hyperFont <- Font(wb, color="blue", underline=1)
    hyperlinkStyle <- CellStyle(wb) +  hyperFont
  }

  # create the link
  creationHelper <- .jcall(wb, "Lorg/apache/poi/ss/usermodel/CreationHelper;",
                           "getCreationHelper")
  type <- switch(linkType, URL=1L, DOCUMENT=2L, EMAIL=3L, FILE=4L)
  link <- .jcall(creationHelper, "Lorg/apache/poi/ss/usermodel/Hyperlink;",
    "createHyperlink", type)
  .jcall(link, "V", "setAddress", address)

  # add the link to the cell and set the cell style
  .jcall(cell, "V", "setHyperlink", link)
  setCellStyle(cell, hyperlinkStyle)

  invisible()
}


######################################################################
# Merge regions 
#
addMergedRegion <- function(sheet, startRow, endRow, startColumn,
  endColumn)
{

  cellRangeAddress <- .jnew("org/apache/poi/ss/util/CellRangeAddress",
    as.integer(startRow-1), as.integer(endRow-1), as.integer(startColumn-1),
    as.integer(endColumn-1))
  ind <- .jcall(sheet, "I", "addMergedRegion", cellRangeAddress)
    
  invisible(ind)  
}
removeMergedRegion <- function(sheet, ind)
  .jcall(sheet, "V", as.integer(ind))



######################################################################
# fit contents to column. col can be a vector of column indices.
#
autoSizeColumn <- function(sheet, colIndex)
{
  for (ic in (colIndex-1))
    .jcall(sheet, "V", "autoSizeColumn", as.integer(ic), TRUE)
    
  invisible()  
}

######################################################################
# fit contents to column. col can be a vector of column indices.
#
createFreezePane <- function(sheet, rowSplit, colSplit, startRow=NULL,
  startColumn=NULL)
{
  if (is.null(startRow) & is.null(startColumn)){
    .jcall(sheet, "V", "createFreezePane", as.integer(colSplit-1),
      as.integer(rowSplit-1))
  } else {   
    .jcall(sheet, "V", "createFreezePane", as.integer(colSplit-1),
      as.integer(rowSplit-1), as.integer(startColumn-1),
      as.integer(startRow-1))
  }
  
  invisible()  
}

######################################################################
# fit contents to column. col can be a vector of column indices.
#
createSplitPane <- function(sheet, xSplitPos=2000, ySplitPos=2000,
  startRow=1, startColumn=1, position="PANE_LOWER_LEFT")
{
  ind <- .jfield(sheet, "B", position)
  
  .jcall(sheet, "V", "createSplitPane", as.integer(xSplitPos),
    as.integer(ySplitPos), as.integer(startRow-1),
    as.integer(startColumn-1), ind)
    
  invisible()  
}


######################################################################
# 
#
setColumnWidth <- function(sheet, colIndex, colWidth)
{
  for (ic in (colIndex - 1))
   .jcall(sheet, "V", "setColumnWidth", as.integer(ic),
          as.integer(colWidth*256))
   
 invisible()
}


######################################################################
# set the zoom
#
setPrintArea <- function(wb, sheetIndex, startColumn, endColumn, startRow,
  endRow)
{
   .jcall(wb, "V", "setPrintArea", as.integer(sheetIndex-1),
     as.integer(startColumn-1),  as.integer(endColumn-1),
     as.integer(startRow-1),  as.integer(endRow-1))
}



######################################################################
# set the zoom
#
setZoom <- function(sheet, numerator=100, denominator=100)
{
  .jcall(sheet, "V", "setZoom", as.integer(numerator),
         as.integer(denominator))
    
  invisible()  
}


