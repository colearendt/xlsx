\name{Row}
\alias{Row}
\alias{createRow}
\alias{getRows}
\alias{removeRow}
\title{Functions to manipulate rows of a worksheet.}
\description{
Functions to manipulate rows of a worksheet. 
}
\usage{

createRow(sheet, rowIndex=1:5)

getRows(sheet, rowIndex=NULL)

removeRow(sheet, rows=NULL)

}
\arguments{
  \item{sheet}{a worksheet object as returned by \code{createSheet} or
    by subsetting \code{getSheets}.}
  \item{rowIndex}{a numeric vector specifying the index of rows to
    create.  For \code{getRows}, a \code{NULL} value will return all
    non empty rows.}
  \item{rows}{a list of \code{Row} objects.}
}
\details{
  \code{removeRow} is just a convenience wrapper to remove the rows from
  the sheet (before saving).  Internally it calls \code{lapply}.
}
\value{
  For \code{getRows} a list of java object references each  pointing to
  a row.  The list is named with the row number.
}
\author{Adrian Dragulescu}
\seealso{To extract the cells from a given row, see \code{\link{Cell}}.}
\examples{

file <- system.file("tests", "test_import.xlsx", package = "xlsx")

wb <- loadWorkbook(file)  
sheets <- getSheets(wb)

sheet <- sheets[[2]]  
rows  <- getRows(sheet)  # get all the rows

# see all the available java methods that you can call
.jmethods(rows[[1]])

# for example
rows[[1]]$getRowNum()   # zero based index in Java

removeRow(sheet, rows)  # remove them all

}

