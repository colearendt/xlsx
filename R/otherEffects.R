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

#' Functions to do various spreadsheets effects.
#'
#' Function \code{autoSizeColumn} expands the column width to match the column
#' contents thus removing the ###### that you get when cell contents are larger
#' than cell width.
#'
#' You may need other functionality that is not exposed.  Take a look at the
#' java docs and the source code of these functions for how you can implement
#' it in R.
#'
#' @param cellRange a string specifying the cell range.  For example a standard
#' area ref (e.g. "B1:D8"). May be a single cell ref (e.g. "B5") in which case
#' the result is a 1 x 1 cell range. May also be a whole row range (e.g.
#' "3:5"), or a whole column range (e.g. "C:F")
#' @param colIndex a numeric vector specifiying the columns you want to auto
#' size.
#' @param colSplit a numeric value for the column to split.
#' @param colWidth a numeric value to specify the width of the column.  The
#' units are in 1/256ths of a character width.
#' @param denominator a numeric value representing the denomiator of the zoom
#' ratio.
#' @param endColumn a numeric value for the ending column.
#' @param endRow a numeric value for the ending row.
#' @param ind a numeric value indicating which merged region you want to
#' remove.
#' @param numerator a numeric value representing the numerator of the zoom
#' ratio.
#' @param position a character.  Valid value are "PANE_LOWER_LEFT",
#' "PANE_LOWER_RIGHT", "PANE_UPPER_LEFT", "PANE_UPPER_RIGHT".
#' @param rowSplit a numeric value for the row to split.
#' @param sheet a \code{\link{Worksheet}} object.
#' @param sheetIndex a numeric value for the worksheet index.
#' @param startColumn a numeric value for the starting column.
#' @param startRow a numeric value for the starting row.
#' @param xSplitPos a numeric value for the horizontal position of split in
#' 1/20 of a point.
#' @param ySplitPos a numeric value for the vertical position of split in 1/20
#' of a point.
#' @param wb a \code{\link{Workbook}} object.
#' @return \code{addMergedRegion} returns a numeric value to label the merged
#' region.  You should use this value as the \code{ind} if you want to
#' \code{removeMergedRegion}.
#' @author Adrian Dragulescu
#' @examples
#'
#'
#'   wb <- createWorkbook()
#'   sheet1 <- createSheet(wb, "Sheet1")
#'   rows   <- createRow(sheet1, 1:10)              # 10 rows
#'   cells  <- createCell(rows, colIndex=1:8)       # 8 columns
#'
#'   ## Merge cells
#'   setCellValue(cells[[1,1]], "A title that spans 3 columns")
#'   addMergedRegion(sheet1, 1, 1, 1, 3)
#'
#'   ## Set zoom 2:1
#'   setZoom(sheet1, 200, 100)
#'
#'   sheet2 <- createSheet(wb, "Sheet2")
#'   rows  <- createRow(sheet2, 1:10)              # 10 rows
#'   cells <- createCell(rows, colIndex=1:8)       # 8 columns
#'   #createFreezePane(sheet2, 1, 1, 1, 1)
#'   createFreezePane(sheet2, 5, 5, 8, 8)
#'
#'   sheet3 <- createSheet(wb, "Sheet3")
#'   rows  <- createRow(sheet3, 1:10)              # 10 rows
#'   cells <- createCell(rows, colIndex=1:8)       # 8 columns
#'   createSplitPane(sheet3, 2000, 2000, 1, 1, "PANE_LOWER_LEFT")
#'
#'   # set the column width of first column to 25 characters wide
#'   setColumnWidth(sheet1, 1, 25)
#'
#'   # add a filter on the 3rd row, columns C:E
#'   addAutoFilter(sheet1, "C3:E3")
#'
#'   # Don't forget to save the workbook ...
#'
#' @name OtherEffects
NULL

######################################################################
# Add autofilter to a cell range
# @param sheet a Sheet object
# @param cellRange a string, "C5:F200", or one row "C1:F1",
#   a column range "C:F", a singe cell "B5", or a row range "3:5"
#
#' @export
#' @rdname OtherEffects
addAutoFilter <- function(sheet, cellRange)
{
  cellRangeAddress <- J("org/apache/poi/ss/util/CellRangeAddress")$valueOf(cellRange)
  invisible(sheet$setAutoFilter(cellRangeAddress))
}


######################################################################
# Add hyperlink
#
#' Add a hyperlink to a cell.
#'
#' Add a hyperlink to a cell to point to an external resource.
#'
#' The cell needs to have content before you add a hyperlink to it.  The
#' contents of the cells don't need to be the same as the address of the
#' hyperlink.  See the examples.
#'
#' @param cell a \code{\link{Cell}} object.
#' @param address a string pointing to the resource.
#' @param linkType a the type of the resource.
#' @param hyperlinkStyle a \code{\link{CellStyle}} object.  If \code{NULL} a
#' default cell style is created, blue underlined font.
#' @return None.  The modification to the cell is done in place.
#' @author Adrian Dragulescu
#' @examples
#'
#'
#'   wb <- createWorkbook()
#'   sheet1 <- createSheet(wb, "Sheet1")
#'   rows   <- createRow(sheet1, 1:10)              # 10 rows
#'   cells  <- createCell(rows, colIndex=1:8)       # 8 columns
#'
#'   ## Add hyperlinks to a cell
#'   cell <- cells[[1,1]]
#'   address <- "https://poi.apache.org/"
#'   setCellValue(cell, "click me!")
#'   addHyperlink(cell, address)
#'
#'   # Don't forget to save the workbook ...
#'
#' @export
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
#' @export
#' @rdname OtherEffects
addMergedRegion <- function(sheet, startRow, endRow, startColumn,
  endColumn)
{

  cellRangeAddress <- .jnew("org/apache/poi/ss/util/CellRangeAddress",
    as.integer(startRow-1), as.integer(endRow-1), as.integer(startColumn-1),
    as.integer(endColumn-1))
  ind <- .jcall(sheet, "I", "addMergedRegion", cellRangeAddress)

  invisible(ind)
}

#' @export
#' @rdname OtherEffects
removeMergedRegion <- function(sheet, ind)
  .jcall(sheet, "V", as.integer(ind))



######################################################################
# fit contents to column. col can be a vector of column indices.
#
#' @export
#' @rdname OtherEffects
autoSizeColumn <- function(sheet, colIndex)
{
  for (ic in (colIndex-1))
    .jcall(sheet, "V", "autoSizeColumn", as.integer(ic), TRUE)

  invisible()
}

######################################################################
# fit contents to column. col can be a vector of column indices.
#
#' @export
#' @rdname OtherEffects
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
#' @export
#' @rdname OtherEffects
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
#' @export
#' @rdname OtherEffects
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
#' @export
#' @rdname OtherEffects
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
#' @export
#' @rdname OtherEffects
setZoom <- function(sheet, numerator=100, denominator=100)
{
  .jcall(sheet, "V", "setZoom", as.integer(numerator),
         as.integer(denominator))

  invisible()
}


