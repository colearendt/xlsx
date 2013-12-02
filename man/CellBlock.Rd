\name{CellBlock}
\alias{CellBlock}
\alias{CellBlock.default}
\alias{CB.setColData}
\alias{CB.setRowData}
\alias{CB.setMatrixData}
\alias{CB.setFill}
\alias{CB.setFont}
\alias{CB.setBorder}
\alias{is.CellBlock}

\title{Create and style a block of cells.}
\description{
  Functions to create and style (not read) a block of cells.  Use it to 
  set/update cell values and cell styles in an efficient manner.
}
\usage{

\method{CellBlock}{default}(sheet, startRow, startColumn, noRows, noColumns,
  create=TRUE)

is.CellBlock( cellBlock )

CB.setColData(cellBlock, x, colIndex, rowOffset=0, showNA=TRUE,
  colStyle=NULL)

CB.setRowData(cellBlock, x, rowIndex, colOffset=0, showNA=TRUE,
  rowStyle=NULL)

CB.setMatrixData(cellBlock, x, startRow, startColumn,
  showNA=TRUE, cellStyle=NULL)

CB.setFill( cellBlock, fill, rowIndex, colIndex)

CB.setFont( cellBlock, font, rowIndex, colIndex )

CB.setBorder( cellBlock, border, rowIndex, colIndex)

}
\arguments{
  \item{sheet}{a \code{\link{Sheet}} object.}
    
  \item{startRow}{a numeric value for the starting row.}
  
  \item{startColumn}{a numeric value for the starting column.}

  \item{rowOffset}{a numeric value for the starting row.}

  \item{colOffset}{a numeric value for the starting column.}

  \item{showNA}{a logical value.  If set to \code{FALSE}, NA values will
    be left as empty cells. }

  \item{noRows}{a numeric value to specify the number of rows for the block.}
  
  \item{noColumns}{a numeric value to specify the number of columns for
  the block.}
  
  \item{create}{ If \code{TRUE} cells will be created if they don't
    exist, if \code{FALSE} only existing cells will be used.  If cells
    don't exist, you have to use \code{TRUE}.  Use \code{TRUE} if you
    want to blank out an existing cell block.  Use \code{FALSE} if you
    want to reuse/restyle existing cells.}
  
  \item{cellBlock}{ a cell block object as returned by
    \code{\link{CellBlock}}. }

  \item{rowStyle}{a \code{\link{CellStyle}} object used to style the row.}

  \item{colStyle}{a \code{\link{CellStyle}} object used to style the column.}
  
  \item{cellStyle}{a \code{\link{CellStyle}} object.}
  
  \item{border}{a Border object, as returned by \code{\link{Border}}.}

  \item{fill}{a Fill object, as returned by \code{\link{Fill}}.}

  \item{font}{a Font object, as returned by \code{\link{Font}}.}

  \item{colIndex}{a numeric vector specifiying the columns you want.}

  \item{rowIndex}{a numeric vector specifiying the rows you want.}

  \item{x}{the data you want to add to the cell block, a vector or a
    matrix depending on the function.}
}

\details{
  
  Introduced in version 0.5.0 of the package, these functions speed up
  the creation and styling of cells that are part of a "cell block" (a
  rectangular shaped group of cells).  Use the functions above if you
  want to create efficiently a complex sheet with many styles.  A simple
  by-column styling can be done by directly using
  \code{\link{addDataFrame}}.  With the functionality provided here you
  can efficiently style individual cells, see the example.

  It is difficult to treat \code{NA}'s consistently between R
  and Excel via Java.  Most likely, users of Excel will want to see
  \code{NA}'s  as blank cells.  In R character \code{NA}'s are simply
  characters, which for Excel means "NA".

  If you try to set more data to the block than you have cells in the
  block, only the existing cells will be set.

  Note that when modifying the style of a group of cells, the changes
  are made to the pairs defined by \code{(rowIndex, colIndex)}.  This
  implies that the length of \code{rowIndex} and \code{colIndex} are the
  same value.  An exception is made when either \code{rowIndex} or
  \code{colIndex} have length one, when they will be expanded internally
  to match the length of the other index.

  Function \code{CB.setMatrixData} works for numeric or character
  matrices.  If the matrix \code{x} is not of numeric type it will be
  converted to a character matrix.
  
  
}

\value{
  For \code{CellBlock} a cell block object. 
  
  For \code{CB.setColData}, \code{CB.setRowData},
  \code{CB.setMatrixData}, \code{CB.setFill}, \code{CB.setFont},
  \code{CB.setBorder} nothing as he modification to the workbook is done
  in place.
  
}
\author{Adrian Dragulescu}
\examples{

  wb <- createWorkbook()
  sheet  <- createSheet(wb, sheetName="CellBlock")

  cb <- CellBlock(sheet, 7, 3, 1000, 60)
  CB.setColData(cb, 1:100, 1)    # set a column
  CB.setRowData(cb, 1:50, 1)     # set a row

  # add a matrix, and style it
  cs <- CellStyle(wb) + DataFormat("#,##0.00")
  x  <- matrix(rnorm(900*45), nrow=900)
  CB.setMatrixData(cb, x, 10, 4, cellStyle=cs)  

  # highlight the negative numbers in red 
  fill <- Fill(foregroundColor = "red", backgroundColor="red")
  ind  <- which(x < 0, arr.ind=TRUE)
  CB.setFill(cb, fill, ind[,1]+9, ind[,2]+3)  # note the indices offset

  # set the border on the top row of the Cell Block
  border <-  Border(color="blue", position=c("TOP", "BOTTOM"),
    pen=c("BORDER_THIN", "BORDER_THICK"))
  CB.setBorder(cb, border, 1:1000, 1)


  # Don't forget to save the workbook ...  
  # saveWorkbook(wb, file) 
}

