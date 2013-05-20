\name{NamedRanges}
\alias{NamedRanges}
\alias{Range}
\alias{getRanges}
\alias{readRange}
\alias{createRange}
\title{Functions to manipulate named ranges.}
\description{
Functions to manipulate (contiguous) named ranges. 
}
\usage{

getRanges(wb)

readRange(range, sheet, colClasses="character")

createRange(rangeName, firstCell, lastCell)

}
\arguments{
  \item{wb}{a workbook object as returned by \code{createWorksheet} or
    \code{loadWorksheet}.}
  \item{range}{a range object as returned by \code{getRanges}.}
  \item{sheet}{a sheet object as returned by \code{getSheets}.}
  \item{rangeName}{a character specifying the name of the name to
    create.}
  \item{colClasses}{the type of the columns supported.  Only
	\code{numeric} and \code{character} are supported.  See
	\code{\link{read.xlsx2}} for more details.}
  \item{firstCell}{a cell object corresponding to the top left cell in
	the range.}
  \item{lastCell}{a cell object corresponding to the bottom right cell in
	the range.}
}

\value{
  \code{getRanges} returns the existing ranges as a list.

  \code{readRange} reads the range into a data.frame.
  
  \code{createRange} returns the created range object.
  
}
\details{
These functions are provided for convenience only.  Use directly the
Java API to access additional functionality.
}
\author{Adrian Dragulescu}
\examples{

file <- system.file("tests", "test_import.xlsx", package = "xlsx")

wb <- loadWorkbook(file)
sheet <- getSheets(wb)[["deletedFields"]]  
ranges <- getRanges(wb)

# the call below fails on cran tests for MacOS.  You should see the
# FAQ: http://code.google.com/p/rexcel/wiki/FAQ
#res  <- readRange(ranges[[1]], sheet, colClasses="numeric") # read it 

ranges[[1]]$getNameName()  # get its name

# see all the available java methods that you can call
.jmethods(ranges[[1]])

# create a new named range
firstCell <- sheet$getRow(14L)$getCell(4L)
lastCell  <- sheet$getRow(20L)$getCell(7L)
rangeName <- "Test2"
# same issue on MacOS
#createRange(rangeName, firstCell, lastCell)


}

