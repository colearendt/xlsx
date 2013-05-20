\name{Comment}
\alias{Comment}
\alias{createCellComment}
\alias{getCellComment}
\alias{removeCellComment}
\title{Functions to manipulate cell comments.}
\description{
Functions to manipulate cell comments. 
}
\usage{

createCellComment(cell, string="", author=NULL, visible=TRUE)

getCellComment(cell)

removeCellComment(cell)

}
\arguments{
  \item{cell}{a \code{Cell} object.}
  \item{string}{a string for the comment.}
  \item{author}{a string with the author's name}
  \item{visible}{a logical value.  If \code{TRUE} the comment will be visible.}
}

\value{
  
  \code{createCellComment} creates a \code{Comment} object.
      
  \code{getCellComment} returns a the \code{Comment} object if it
  exists. 
  
  \code{removeCellComment} removes a comment from the given cell. 
  
}
\details{
  These functions are not vectorized. 
}

\author{Adrian Dragulescu}
\seealso{For cells, see \code{\link{Cell}}.  To format cells, see
  \code{\link{CellStyle}}.}
\examples{

  wb <- createWorkbook()
  sheet1 <- createSheet(wb, "Sheet1")
  rows   <- createRow(sheet1, rowIndex=1:10)     # 10 rows
  cells  <- createCell(rows, colIndex=1:8)       # 8 columns

  cell1 <- cells[[1,1]]
  setCellValue(cell1, 1)   # add value 1 to cell A1

  # create a cell comment
  createCellComment(cell1, "Cogito", author="Descartes")
  

  # extract the comments 
  comment <- getCellComment(cell1)
  stopifnot(comment$getAuthor()=="Descartes")
  stopifnot(comment$getString()$toString()=="Cogito")

  # don't forget to save your workbook!
}

