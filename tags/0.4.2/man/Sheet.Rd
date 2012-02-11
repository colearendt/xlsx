\name{Sheet}
\alias{Sheet}
\alias{Worksheet}
\alias{createSheet}
\alias{removeSheet}
\alias{getSheets}
\title{Functions to manipulate worksheets.}
\description{
Functions to manipulate worksheets. 
}
\usage{

createSheet(wb, sheetName="Sheet1")

removeSheet(wb, sheetName="Sheet1")

getSheets(wb)

}
\arguments{
  \item{wb}{a workbook object as returned by \code{createWorksheet} or
    \code{loadWorksheet}.}
  \item{sheetName}{a character specifying the name of the worksheet to
    create, or remove.}
}

\value{
  \code{createSheet} returns the created \code{Sheet} object.
  
  \code{getSheets} returns a list of java object references each
  pointing to an worksheet.  The list is named with the sheet names. 
}
\author{Adrian Dragulescu}
\seealso{To extract rows from a given sheet, see \code{\link{Row}}.}
\examples{

file <- system.file("tests", "test_import.xlsx", package = "xlsx")

wb <- loadWorkbook(file)  
sheets <- getSheets(wb)

sheet  <- sheets[[2]]  # extract the second sheet

# see all the available java methods that you can call
.jmethods(sheet)

# for example
sheet$getLastRowNum()

}

