# Cell Block functionality.  This is to be used when writing to a spreadsheet. 
#
#

CellBlock <- function(sheet, startRow, startColumn, noRows, noColumns,
  create=TRUE)  UseMethod("CellBlock")

is.CellBlock <- function(x) {inherits(x, "CellBlock")}

#########################################################################     
# Create a cell block
#
CellBlock.default <- function(sheet, startRow, startColumn, noRows, noColumns,
  create=TRUE)
{
  cb <- new(J("org/cran/rexcel/RCellBlock"), sheet, as.integer(startRow-1),
    as.integer(startColumn-1), as.integer(noRows), as.integer(noColumns), create)

  structure(list(ref=cb), class="CellBlock")
}


#########################################################################     
# set the column data for a Cell Block
#
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
CB.setRowData <- function(cellBlock, x, rowIndex, colOffset=0, showNA=TRUE,
  rowStyle=NULL)
{
  .jcall( cellBlock$ref, "V", "setRowData", as.integer(rowIndex-1),
    as.integer(colOffset), .jarray(as.character(x)), showNA,
    if ( !is.null(rowStyle) ) rowStyle$ref else
      .jnull('org/apache/poi/ss/usermodel/CellStyle') )
}

#########################################################################     
# set a numeric matrix of data for a Cell Block
#
CB.setMatrixData <- function(cellBlock, x, startRow, startColumn,
  showNA=TRUE, cellStyle=NULL)
{
  endRow <- startRow + dim(x)[1] - 1
  endColumn <- startColumn + dim(x)[2] - 1

  if ( (endRow-startRow+1) > .jfield(cellBlock$ref, "I", "noRows") )
    stop("The rows of x don't fit in the cellBlock!")
  if ( (endColumn-startColumn+1) > .jfield(cellBlock$ref, "I", "noCols") )
    stop("The columns of x don't fit in the cellBlock!")
  
  .jcall( cellBlock$ref, "V", "setMatrixData", as.integer(startRow-1),
    as.integer(endRow-1), as.integer(startColumn-1), as.integer(endColumn-1),
    .jarray(as.numeric(x)), showNA,
    if ( !is.null(cellStyle) ) cellStyle$ref else
      .jnull('org/apache/poi/ss/usermodel/CellStyle') )
  
}


#########################################################################     
# set the Fill for an array of indices of a Cell Block
#
CB.setFill <- function( cellBlock, fill, rowIndex, colIndex)
{
  if (length(rowIndex) != length(colIndex))
    stop("Length of indRows should equal length of indCols!")
  
  if ( cellBlock$ref$isXSSF() ) {
    .jcall( cellBlock$ref, 'V', 'setFill',
      .xssfcolor( fill$foregroundColor ),
      .xssfcolor( fill$backgroundColor ),
      .jshort(FILL_STYLES_[[fill$pattern]]),
      .jarray( as.integer( rowIndex-1 ) ),
      .jarray( as.integer( colIndex-1 ) ) )
  } else {
    stop("Implement CB.setFill for HSSF!")
  }
  
  invisible()
}


#########################################################################     
# set the Font for an array of indices of a Cell Block
#
CB.setFont <- function( cellBlock, font, rowIndex, colIndex )
{
  cellBlock$setFont( font$ref, .jarray( as.integer( rowIndex-1 ) ),
    .jarray( as.integer( colIndex-1 ) ) )
    
  invisible()
}

#########################################################################     
# set the Border for an array of indices of a Cell Block
#
CB.setBorder <- function( cellBlock, border, rowIndex, colIndex)
{
  isXSSF <- .jcall( cellBlock$ref, 'Z', 'isXSSF' )

  border_none <- BORDER_STYLES_[['BORDER_NONE']]
  borders <- c( TOP    = border_none,
                BOTTOM = border_none,
                LEFT   = border_none,
                RIGHT  = border_none )
  
  null_color <- .jnull('org/apache/poi/xssf/usermodel/XSSFColor')
  border_colors <- c( TOP    = null_color,
                      BOTTOM = null_color,
                      LEFT   = null_color,
                      RIGHT  = null_color )
  borders[ border$position ] <- sapply( border$pen,
    function( pen ) BORDER_STYLES_[pen] )
  border_colors[ border$position ] <- .Rcolor2XLcolor( border$color, isXSSF)

  .jcall( cellBlock, "V", "putBorder",
          .jshort(borders[['TOP']]),    border_colors[['TOP']],
          .jshort(borders[['BOTTOM']]), border_colors[['BOTTOM']],
          .jshort(borders[['LEFT']]),   border_colors[['LEFT']],
          .jshort(borders[['RIGHT']]),  border_colors[['RIGHT']],
          .jarray( as.integer( rowIndex-1L ) ),
          .jarray( as.integer( colIndex-1L ) ) )
  
  invisible()
}

## CB.getCell <- function( cellBlock, rowIndex, colIndex )
## {
##   .jcall( cellBlock, 'Lorg/apache/poi/ss/usermodel/Cell;', 'getCell',
##     rowIndex - 1L, colIndex - 1L )
## }
