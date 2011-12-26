\name{Workbook}
\alias{Workbook}
\alias{createWorkbook}
\alias{loadWorkbook}
\alias{saveWorkbook}
\title{Functions to manipulate Excel 2007 workbooks.}
\description{
Functions to manipulate Excel 2007 workbooks. 
}
\usage{

createWorkbook(type="xlsx")

loadWorkbook(file)

saveWorkbook(wb, file)

}
\arguments{
  \item{type}{a String, either \code{xlsx} for Excel 2007 OOXML format,
	or \code{xls} for Excel 95 binary format.}
  \item{file}{the \emph{absolute} path to the file which the data are to
    be read from.}
  \item{wb}{a workbook object as returned by \code{createWorkbook} or
    \code{loadWorkbook}.}
}

\details{
  \code{createWorkbook} creates an empty workbook object.

  \code{loadWorkbook} loads a workbook from a file.

  \code{saveWorkbook} saves an existing workook to an Excel 2007 file. 
}
\value{
  \code{createWorkbook} returns a java object reference pointing to an 
  empty workbook object.
  
  \code{loadWorkbook} creates a java object reference corresponding to
  the file to load.

}
\author{Adrian Dragulescu}
\seealso{\code{\link{write.xlsx}} for writing a \code{data.frame} to an
  \code{xlsx} file.  \code{\link{read.xlsx}} for reading the content of
  a \code{xlsx} worksheet into a \code{data.frame}.  To extract
  worksheets and manipulate them, see \code{\link{Worksheet}}.}
\examples{

wb <- createWorkbook()

# see all the available java methods that you can call
.jmethods(wb)

# for example
wb$getNumberOfSheets()   # no sheet yet!


}

