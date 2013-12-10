\name{addHyperlink}
\alias{addHyperlink}
\title{Add a hyperlink to a cell.}
\description{
Add a hyperlink to a cell to point to an external resource.
}
\usage{

addHyperlink(cell, address, linkType=c("URL", "DOCUMENT",
  "EMAIL", "FILE"), hyperlinkStyle=NULL)

}
\arguments{
  \item{cell}{a \code{\link{Cell}} object.}

  \item{address}{a string pointing to the resource.}
  
  \item{linkType}{a the type of the resource.}
  
  \item{hyperlinkStyle}{a \code{\link{CellStyle}} object.  If
    \code{NULL} a default cell style is created, blue underlined font.}
   
}

\details{
  The cell needs to have content before you add a hyperlink to it.  The
  contents of the cells don't need to be the same as the address of the
  hyperlink.  See the examples. 
}

\value{
  None.  The modification to the cell is done in place. 
}
\author{Adrian Dragulescu}
\examples{

  wb <- createWorkbook()
  sheet1 <- createSheet(wb, "Sheet1")
  rows   <- createRow(sheet1, 1:10)              # 10 rows
  cells  <- createCell(rows, colIndex=1:8)       # 8 columns

  cat("Add hyperlinks to a cell")
  cell <- cells[[1,1]]
  address <- "http://poi.apache.org/"
  setCellValue(cell, "click me!")  
  addHyperlink(cell, address)
  
  # Don't forget to save the workbook ...  
 
}

