# Cell Block functionality.  This is to be used when writing to a spreadsheet.
#
#

#' Create and style a block of cells.
#'
#' Functions to create and style (not read) a block of cells.  Use it to
#' set/update cell values and cell styles in an efficient manner.
#'
#'
#' Introduced in version 0.5.0 of the package, these functions speed up the
#' creation and styling of cells that are part of a "cell block" (a rectangular
#' shaped group of cells).  Use the functions above if you want to create
#' efficiently a complex sheet with many styles.  A simple by-column styling
#' can be done by directly using \code{\link{addDataFrame}}.  With the
#' functionality provided here you can efficiently style individual cells, see
#' the example.
#'
#' It is difficult to treat \code{NA}'s consistently between R and Excel via
#' Java.  Most likely, users of Excel will want to see \code{NA}'s as blank
#' cells.  In R character \code{NA}'s are simply characters, which for Excel
#' means "NA".
#'
#' If you try to set more data to the block than you have cells in the block,
#' only the existing cells will be set.
#'
#' Note that when modifying the style of a group of cells, the changes are made
#' to the pairs defined by \code{(rowIndex, colIndex)}.  This implies that the
#' length of \code{rowIndex} and \code{colIndex} are the same value.  An
#' exception is made when either \code{rowIndex} or \code{colIndex} have length
#' one, when they will be expanded internally to match the length of the other
#' index.
#'
#' Function \code{CB.setMatrixData} works for numeric or character matrices.
#' If the matrix \code{x} is not of numeric type it will be converted to a
#' character matrix.
#'
#' @param sheet a \code{\link{Sheet}} object.
#' @param startRow a numeric value for the starting row.
#' @param startColumn a numeric value for the starting column.
#' @param rowOffset a numeric value for the starting row.
#' @param colOffset a numeric value for the starting column.
#' @param showNA a logical value.  If set to \code{FALSE}, NA values will be
#' left as empty cells.
#' @param noRows a numeric value to specify the number of rows for the block.
#' @param noColumns a numeric value to specify the number of columns for the
#' block.
#' @param create If \code{TRUE} cells will be created if they don't exist, if
#' \code{FALSE} only existing cells will be used.  If cells don't exist (on a
#' new sheet for example), you have to use \code{TRUE}.  On an existing sheet
#' with data, use \code{TRUE} if you want to blank out an existing cell block.
#' Use \code{FALSE} if you want to keep the styling of existing cells, but just
#' modify the value of the cell.
#' @param cellBlock a cell block object as returned by \code{\link{CellBlock}}.
#' @param rowStyle a \code{\link{CellStyle}} object used to style the row.
#' @param colStyle a \code{\link{CellStyle}} object used to style the column.
#' @param cellStyle a \code{\link{CellStyle}} object.
#' @param border a Border object, as returned by \code{\link{Border}}.
#' @param fill a Fill object, as returned by \code{\link{Fill}}.
#' @param font a Font object, as returned by \code{\link{Font}}.
#' @param colIndex a numeric vector specifiying the columns you want relative
#' to the \code{startColumn}.
#' @param rowIndex a numeric vector specifiying the rows you want relative to
#' the \code{startRow}.
#' @param x the data you want to add to the cell block, a vector or a matrix
#' depending on the function.
#' @return For \code{CellBlock} a cell block object.
#'
#' For \code{CB.setColData}, \code{CB.setRowData}, \code{CB.setMatrixData},
#' \code{CB.setFill}, \code{CB.setFont}, \code{CB.setBorder} nothing as he
#' modification to the workbook is done in place.
#' @author Adrian Dragulescu
#' @examples
#'
#' \donttest{
#'   wb <- createWorkbook()
#'   sheet  <- createSheet(wb, sheetName="CellBlock")
#'
#'   cb <- CellBlock(sheet, 7, 3, 1000, 60)
#'   CB.setColData(cb, 1:100, 1)    # set a column
#'   CB.setRowData(cb, 1:50, 1)     # set a row
#'
#'   # add a matrix, and style it
#'   cs <- CellStyle(wb) + DataFormat("#,##0.00")
#'   x  <- matrix(rnorm(900*45), nrow=900)
#'   CB.setMatrixData(cb, x, 10, 4, cellStyle=cs)
#'
#'   # highlight the negative numbers in red
#'   fill <- Fill(foregroundColor = "red", backgroundColor="red")
#'   ind  <- which(x < 0, arr.ind=TRUE)
#'   CB.setFill(cb, fill, ind[,1]+9, ind[,2]+3)  # note the indices offset
#'
#'   # set the border on the top row of the Cell Block
#'   border <-  Border(color="blue", position=c("TOP", "BOTTOM"),
#'     pen=c("BORDER_THIN", "BORDER_THICK"))
#'   CB.setBorder(cb, border, 1, 1:1000)
#'
#'
#'   # Don't forget to save the workbook ...
#'   # saveWorkbook(wb, file)
#' }
#'
#' @export
CellBlock <- function(sheet, startRow, startColumn, noRows, noColumns,
  create=TRUE)  UseMethod("CellBlock")

#' @export
#' @rdname CellBlock
is.CellBlock <- function(cellBlock) {inherits(cellBlock, "CellBlock")}

#########################################################################
# Create a cell block
#
#' @export
#' @rdname CellBlock
CellBlock.default <- function(sheet, startRow, startColumn, noRows, noColumns,
  create=TRUE)
{
  cb <- new(J("org/cran/rexcel/RCellBlock"), sheet, as.integer(startRow-1),
    as.integer(startColumn-1), as.integer(noRows), as.integer(noColumns),
    create)

  structure(list(ref=cb), class="CellBlock")
}


#########################################################################
# set the column data for a Cell Block
#
#' @export
#' @rdname CellBlock
CB.setColData <- function(cellBlock, x, colIndex, rowOffset=0, showNA=TRUE,
  colStyle=NULL)
{
  .jcall( cellBlock$ref, "V", "setColData", as.integer(colIndex-1),
    as.integer(rowOffset), .jarray(x), showNA,
    if ( !is.null(colStyle) ) colStyle$ref else
      .jnull('org/apache/poi/ss/usermodel/CellStyle') )
}

#########################################################################
# set the row data for a Cell Block
#
#' @export
#' @rdname CellBlock
CB.setRowData <- function(cellBlock, x, rowIndex, colOffset=0, showNA=TRUE,
  rowStyle=NULL)
{
  .jcall( cellBlock$ref, "V", "setRowData", as.integer(rowIndex-1),
    as.integer(colOffset), .jarray(as.character(x)), showNA,
    if ( !is.null(rowStyle) ) rowStyle$ref else
      .jnull('org/apache/poi/ss/usermodel/CellStyle') )
}

#########################################################################
# set a matrix of data for a Cell Block.
# if x is not a numeric matrix, coerce it a character matrix
#
#' @export
#' @rdname CellBlock
CB.setMatrixData <- function(cellBlock, x, startRow, startColumn,
  showNA=TRUE, cellStyle=NULL)
{
  endRow <- startRow + dim(x)[1] - 1
  endColumn <- startColumn + dim(x)[2] - 1

  if ( (endRow-startRow+1) > .jfield(cellBlock$ref, "I", "noRows") )
    stop("The rows of x don't fit in the cellBlock!")
  if ( (endColumn-startColumn+1) > .jfield(cellBlock$ref, "I", "noCols") )
    stop("The columns of x don't fit in the cellBlock!")

  if (class(x[1,1])[1] == "numeric") {
    .jcall( cellBlock$ref, "V", "setMatrixData", as.integer(startRow-1),
      as.integer(endRow-1), as.integer(startColumn-1), as.integer(endColumn-1),
      .jarray(x), showNA,
      if ( !is.null(cellStyle) ) cellStyle$ref else
        .jnull('org/apache/poi/ss/usermodel/CellStyle') )
  } else {
    .jcall( cellBlock$ref, "V", "setMatrixData", as.integer(startRow-1),
      as.integer(endRow-1), as.integer(startColumn-1), as.integer(endColumn-1),
      .jarray(as.character(x)), showNA,
      if ( !is.null(cellStyle) ) cellStyle$ref else
        .jnull('org/apache/poi/ss/usermodel/CellStyle') )
  }

}


#########################################################################
# set the Fill for an array of indices of a Cell Block
#
#' @export
#' @rdname CellBlock
CB.setFill <- function( cellBlock, fill, rowIndex, colIndex)
{
  if (length(colIndex)==1 & length(rowIndex) > 1)
    colIndex <- rep(colIndex, length(rowIndex))
  if (length(rowIndex)==1 & length(colIndex) > 1)
    rowIndex <- rep(rowIndex, length(colIndex))

  if (length(rowIndex) != length(colIndex))
    stop("rowIndex and colIndex arguments don't have the same length!")


  if ( cellBlock$ref$isXSSF() ) {
    .jcall( cellBlock$ref, 'V', 'setFill',
      .xssfcolor( fill$foregroundColor ),
      .xssfcolor( fill$backgroundColor ),
      .jshort(FILL_STYLES_[[fill$pattern]]),
      .jarray( as.integer( rowIndex-1 ) ),
      .jarray( as.integer( colIndex-1 ) ) )
  } else {
    .jcall( cellBlock$ref, 'V', 'setFill',
      .hssfcolor( fill$foregroundColor ),
      .hssfcolor( fill$backgroundColor ),
      .jshort(FILL_STYLES_[[fill$pattern]]),
      .jarray( as.integer( rowIndex-1 ) ),
      .jarray( as.integer( colIndex-1 ) ) )
  }

  invisible()
}


#########################################################################
# set the Font for an array of indices of a Cell Block
#
#' @export
#' @rdname CellBlock
CB.setFont <- function( cellBlock, font, rowIndex, colIndex )
{
  if (length(colIndex)==1 & length(rowIndex) > 1)
    colIndex <- rep(colIndex, length(rowIndex))
  if (length(rowIndex)==1 & length(colIndex) > 1)
    rowIndex <- rep(rowIndex, length(colIndex))

  if (length(rowIndex) != length(colIndex))
    stop("rowIndex and colIndex arguments don't have the same length!")

  .jcall( cellBlock$ref, 'V', 'setFont', font$ref,
    .jarray( as.integer( rowIndex-1 ) ),
    .jarray( as.integer( colIndex-1 ) ) )

  invisible()
}

#########################################################################
# set the Border for an array of indices of a Cell Block
#
#' @export
#' @rdname CellBlock
CB.setBorder <- function( cellBlock, border, rowIndex, colIndex)
{
  if (length(colIndex)==1 & length(rowIndex) > 1)
    colIndex <- rep(colIndex, length(rowIndex))
  if (length(rowIndex)==1 & length(colIndex) > 1)
    rowIndex <- rep(rowIndex, length(colIndex))

  if (length(rowIndex) != length(colIndex))
    stop("rowIndex and colIndex arguments don't have the same length!")

  isXSSF <- .jcall( cellBlock$ref, 'Z', 'isXSSF' )

  border_none <- BORDER_STYLES_[['BORDER_NONE']]
  borders <- c( TOP    = border_none,
                BOTTOM = border_none,
                LEFT   = border_none,
                RIGHT  = border_none )
  borders[ border$position ] <- sapply( border$pen,
    function( pen ) BORDER_STYLES_[pen] )

  null_color <- if (isXSSF) {
      .jnull('org/apache/poi/xssf/usermodel/XSSFColor')
    } else {
      0
    }
  border_colors <- c( TOP    = null_color,
                      BOTTOM = null_color,
                      LEFT   = null_color,
                      RIGHT  = null_color )
  border_colors[ border$position ] <- .Rcolor2XLcolor( border$color, isXSSF)

  if ( isXSSF ) {
    .jcall( cellBlock$ref, "V", "putBorder",
          .jshort(borders[['TOP']]),    border_colors[['TOP']],
          .jshort(borders[['BOTTOM']]), border_colors[['BOTTOM']],
          .jshort(borders[['LEFT']]),   border_colors[['LEFT']],
          .jshort(borders[['RIGHT']]),  border_colors[['RIGHT']],
          .jarray( as.integer( rowIndex-1L ) ),
          .jarray( as.integer( colIndex-1L ) ) )
  } else {
    .jcall( cellBlock$ref, "V", "putBorder",
          .jshort(borders[['TOP']]),    .jshort(border_colors[['TOP']]),
          .jshort(borders[['BOTTOM']]), .jshort(border_colors[['BOTTOM']]),
          .jshort(borders[['LEFT']]),   .jshort(border_colors[['LEFT']]),
          .jshort(borders[['RIGHT']]),  .jshort(border_colors[['RIGHT']]),
          .jarray( as.integer( rowIndex-1L ) ),
          .jarray( as.integer( colIndex-1L ) ) )
  }

  invisible()
}

############################################################################
# get reference to java cell object by its CellBlock row and column indices
# DON'T NEED TO EXPOSE THIS it's just: cb$ref$getCell(1L, 1L)
## CB.getCell <- function( cellBlock, rowIndex, colIndex )
## {
##     invisible(
##     .jcall( cellBlock$ref, 'Lorg/apache/poi/ss/usermodel/Cell;', 'getCell',
##          rowIndex - 1L, colIndex - 1L ) )
## }
