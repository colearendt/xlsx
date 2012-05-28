\name{Border}
\alias{Border}
\alias{is.Border}
\title{Create an Border object.}
\description{
  Create an Border object, useful when working with cell styles.
}
\usage{

Border(color="black", position="BOTTOM", pen="BORDER_THIN")

is.Border(x)

}
\arguments{

  \item{color}{a character vector specifiying the font color.  Any
    color names as returned by \code{\link[grDevices]{colors}} can be
    used.  For Excel 95 workbooks, only a subset of colors is
    available, see the constant \code{INDEXED_COLORS_}.}
  
  \item{position}{a character vector specifying the border position.  
    Valid values are "BOTTOM", "LEFT", "TOP", "RIGHT".} 

  \item{pen}{a character vector specifying the pen style.  Valid values
  come from constant \code{BORDER_STYLES_}.} 
    
  \item{x}{An Border object, as returned by \code{Border}.}
  
}

\details{
  The values for the color, position, or pen arguments are replicated to
  the longest of them.  
}
\value{
  \code{Border} returns a list with components from the input
  argument, and a class attribute "Border".  Border objects are
  used when constructing cell styles.

  \code{is.Border} returns \code{TRUE} if the argument is of class
  "Border" and \code{FALSE} otherwise.
  
}
\author{Adrian Dragulescu}
\seealso{\code{\link{CellStyle}} for using the a \code{Border} object.}
\examples{

  border <-  Border(color="red", position=c("TOP", "BOTTOM"),
    pen=c("BORDER_THIN", "BORDER_THICK"))
}


