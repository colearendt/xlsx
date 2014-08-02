\name{OtherEffects}
\alias{OtherEffects}
\alias{addMergedRegion}
\alias{removeMergedRegion}
\alias{autoSizeColumn}
\alias{createFreezePane}
\alias{createSplitPane}
\alias{setColumnWidth}
\alias{setPrintArea}
\alias{setZoom}
\title{Functions to do various spreadsheets effects.}
\description{
Functions to do various spreadsheets effects.
}
\usage{

addMergedRegion(sheet, startRow, endRow, startColumn, endColumn)

removeMergedRegion(sheet, ind)

autoSizeColumn(sheet, colIndex)

createFreezePane(sheet, rowSplit, colSplit, startRow=NULL,
  startColumn=NULL)

createSplitPane(sheet, xSplitPos=2000, ySplitPos=2000,
  startRow=1, startColumn=1, position="PANE_LOWER_LEFT")

setColumnWidth(sheet, colIndex, colWidth)

setPrintArea(wb, sheetIndex, startColumn, endColumn, startRow,
  endRow)

setZoom(sheet, numerator=100, denominator=100)
}
\arguments{
  \item{colIndex}{a numeric vector specifiying the columns you want to
    auto size.}

  \item{colSplit}{a numeric value for the column to split.}

  \item{colWidth}{a numeric value to specify the width of the column.
  The units are in 1/256ths of a character width.}

  \item{denominator}{a numeric value representing the denomiator of the
    zoom ratio.}
  
  \item{endColumn}{a numeric value for the ending column.}
  
  \item{endRow}{a numeric value for the ending row.}  

  \item{ind}{a numeric value indicating which merged region you want to
    remove.}

  \item{numerator}{a numeric value representing the numerator
    of the zoom ratio.}
    
  \item{position}{a character.  Valid value are "PANE_LOWER_LEFT",
    "PANE_LOWER_RIGHT", "PANE_UPPER_LEFT", "PANE_UPPER_RIGHT".}
  
  \item{rowSplit}{a numeric value for the row to split.}
    
  \item{sheet}{a \code{\link{Worksheet}} object.}

  \item{sheetIndex}{a numeric value for the worksheet index.}
   
  \item{startColumn}{a numeric value for the starting column.}
  
  \item{startRow}{a numeric value for the starting row.}
  
  \item{xSplitPos}{a numeric value for the horizontal position of split in
    1/20 of a point.}
  
  \item{ySplitPos}{a numeric value for the vertical position of split in
    1/20 of a point.}
  
  \item{wb}{a \code{\link{Workbook}} object.}
  
}

\details{
  Function \code{autoSizeColumn} expands the column width to match the
  column contents thus removing the ###### that you get when cell
  contents are larger than cell width.
  
  You may need other functionality that is not exposed.  Take a
  look at the java docs and the source code of these functions for how
  you can implement it in R.}

\value{
  \code{addMergedRegion} returns a numeric value to label the merged
  region.  You should use this value as the \code{ind} if you want to
  \code{removeMergedRegion}.
}
\author{Adrian Dragulescu}
\examples{

  wb <- createWorkbook()
  sheet1 <- createSheet(wb, "Sheet1")
  rows   <- createRow(sheet1, 1:10)              # 10 rows
  cells  <- createCell(rows, colIndex=1:8)       # 8 columns

  cat("Merge cells \n")
  setCellValue(cells[[1,1]], "A title that spans 3 columns")
  addMergedRegion(sheet1, 1, 1, 1, 3)

  cat("Set zoom 2:1 \n")
  setZoom(sheet1, 200, 100)

  sheet2 <- createSheet(wb, "Sheet2")
  rows  <- createRow(sheet2, 1:10)              # 10 rows
  cells <- createCell(rows, colIndex=1:8)       # 8 columns
  #createFreezePane(sheet2, 1, 1, 1, 1)
  createFreezePane(sheet2, 5, 5, 8, 8)
  
  sheet3 <- createSheet(wb, "Sheet3")
  rows  <- createRow(sheet3, 1:10)              # 10 rows
  cells <- createCell(rows, colIndex=1:8)       # 8 columns
  createSplitPane(sheet3, 2000, 2000, 1, 1, "PANE_LOWER_LEFT")

  # set the column width of first column to 25 characters wide
  setColumnWidth(sheet1, 1, 25)

  # Don't forget to save the workbook ...  
 
}

