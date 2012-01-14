\name{Fill}
\alias{Fill}
\alias{is.Fill}
\title{Create an Fill object.}
\description{
  Create an Fill object, useful when working with cell styles.
}
\usage{

Fill(foregroundColor="lightblue", backgroundColor="lightblue",
  pattern="SOLID_FOREGROUND")

is.Fill(x)

}
\arguments{

  \item{foregroundColor}{a character vector specifiying the foreground
    color.  Any color names as returned by
    \code{\link[grDevices]{colors}} can be used.  For Excel 95
    workbooks, only a subset of colors is available, see the constant
    \code{INDEXED_COLORS_}.}
  
  \item{backgroundColor}{a character vector specifiying the foreground
    color.  Any color names as returned by
    \code{\link[grDevices]{colors}} can be used.  For Excel 95
    workbooks, only a subset of colors is available, see the constant
    \code{INDEXED_COLORS_}.}

  \item{pattern}{a character vector specifying the fill pattern style.
  Valid values come from constant \code{FILL_STYLES_}.} 
    
  \item{x}{An Fill object, as returned by \code{Fill}.}
}

%\details{}
\value{
  \code{Fill} returns a list with components from the input
  argument, and a class attribute "Fill".  Fill objects are
  used when constructing cell styles.

  \code{is.Fill} returns \code{TRUE} if the argument is of class
  "Fill" and \code{FALSE} otherwise.
  
}
\author{Adrian Dragulescu}
\seealso{\code{\link{CellStyle}} for using the a \code{Fill} object.}
\examples{
  fill <-  Fill()
}


