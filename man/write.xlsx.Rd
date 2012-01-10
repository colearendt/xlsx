\name{write.xlsx}
\alias{write.xlsx}
\alias{write.xlsx2}
\title{Write a data.frame to an Excel 2007 workbook.}
\description{
Write a \code{data.frame} to an Excel 2007 workbook.
}
\usage{
write.xlsx(x, file, sheetName="Sheet1", 
  col.names=TRUE, row.names=TRUE, append=FALSE)

write.xlsx2(x, file, sheetName="Sheet1",
  col.names=TRUE, row.names=TRUE, append=FALSE, ...)
}
\arguments{
  \item{x}{a \code{data.frame} to write to the workbook.}
  \item{file}{the \emph{absolute} path to the file which the data are to
    be read from.}
  \item{sheetName}{a character string with the sheet name.}
  \item{col.names}{a logical value indicating if the column names of
    \code{x} are to be written along with \code{x} to the file.}
  \item{row.names}{a logical value indicating whether the row names of 
    \code{x} are to be written along with \code{x} to the file.}
  \item{append}{a logical value indicating if \code{x} should be
    appended to an existing file.  If \code{TRUE} the file is read
    from disk.}
  
  \item{\ldots}{other arguments to \code{addDataFrame} in the case of
    \code{read.xlsx2}.} 

}

\details{
  
This function provides a high level API for writing a \code{data.frame}
to an Excel 2007 worksheet.  It calls several low level functions in the
process.  Its goal is to provide the conveniency of
\code{\link[utils]{write.csv}} by borrowing from its signature.

Internally, there is a double loop in \R over all the elements of the
\code{data.frame} so performance for very large \code{data.frame} may be
an issue.  Please report if you experience slow performance.  Dates and
POSIXct classes are formatted separately after the insertion.  This also
adds to processing time.

If \code{x} is not a \code{data.frame} it will be converted to one.

Function \code{write.xlsx2} uses \code{addDataFrame} which speeds up the
execution compared to \code{write.xlsx} by an order of magnitude for
large spreadsheets (with more than 100,000 cells).

}
\author{ Adrian Dragulescu }
\seealso{\code{\link{read.xlsx}} for reading \code{xlsx} documents.
  See also \code{\link{addDataFrame}} for writing a \code{data.frame} to a
  sheet.}
\examples{
\dontrun{

file <- paste(tempdir(), "/usarrests.xlsx", sep="")
res <- write.xlsx(USArrests, file)  

}
}

