\name{"+.CellStyle"}
\alias{+.CellStyle}
\title{CellStyle construction.}
\description{
Create cell styles. 
}
\usage{
\method{+}{CellStyle}(cs1, object)
}
\arguments{
  
  \item{cs1}{a \code{\link{CellStyle}} object.}
			
  \item{object}{an object to add.  The object can be another
    \code{\link{CellStyle}}, a \code{\link{DataFormat}}, a
    \code{\link{Alignment}}, a \code{\link{Border}}, a
    \code{\link{Fill}}, a \code{\link{Font}}, or a
    \code{\link{CellProtection}} object.}
}

\value{
  A CellStyle object.
}
\details{
  The style of the argument object takes precedence over the style of
  argument cs1.  
}
\author{Adrian Dragulescu}
\examples{
\dontrun{
  cs <- CellStyle(wb) +
    Font(wb, heightInPoints=20, isBold=TRUE, isItalic=TRUE,
      name="Courier New", color="orange") + 
    Fill(backgroundColor="lavender", foregroundColor="lavender",
      pattern="SOLID_FOREGROUND") +
    Alignment(h="ALIGN_RIGHT")
   
  setCellStyle(cell.1, cellStyle1)  

  # you need to save the workbook now if you want to see this art
}
}

