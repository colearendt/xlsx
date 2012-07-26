\name{CellBlock}
\alias{CellBlock}
\alias{CB.setColData}
\alias{CB.setRowData}
\alias{CB.setMatrixData}
\alias{CB.setFill}
\alias{CB.setFont}
\alias{CB.setBorder}
\alias{CB.getCell}
\alias{is.CellBlock}

\title{Create and manipulate a block of cells.}
\description{
  Functions to create and manipulate a block of cells.  Use it to 
  set/update cell values and cell styles in an efficient manner.
}
\usage{

CellBlock(sheet, startRow, startColumn, noRows, noColumns,
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

CB.getCel( cellBlock, rowIndex, colIndex )

}
\arguments{
  \item{sheet}{a \code{\link{Sheet}} object.}
    
  \item{startRow}{a numeric value for the starting row.}
  
  \item{startColumn}{a numeric value for the starting column.}

  \item{noRows}{a numeric value to specify the number of rows for the block.}
  
  \item{noColumns}{a numeric value to specify the number of columns for
  the block.}
  
  \item{create}{ If \code{TRUE} cells will be created if they don't
    exist,  if \code{FALSE} only existing cells will be used. }
  
  \item{cellBlock}{ a cell block object as returned by
    \code{\link{CellBlock}}. }

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
  Starting with version 0.5.0 this function uses the functionality
  provided by \code{CellBlock} which results in a significant
  improvement in performance compared with a cell by cell application of
  \code{\link{setCellValue}} and with other previous atempts.
  
  It is difficult to treat \code{NA}'s consistently between R
  and Excel via Java.  Most likely, users of Excel will want to see
  \code{NA}'s  as blank cells.  In R character \code{NA}'s are simply
  characters, which for Excel means "NA".

  If you try to set more data to the block than you have cells, only the
  available cells will be set. 
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

