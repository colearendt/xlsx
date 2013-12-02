\name{Alignment}
\alias{Alignment}
\alias{is.Alignment}
\title{Create an Alignment object.}
\description{
  Create an Alignment object, useful when working with cell styles.
}
\usage{

Alignment(horizontal=NULL, vertical=NULL, wrapText=FALSE,
  rotation=0, indent=0)

is.Alignment(x)

}
\arguments{
  
  \item{horizontal}{a character value specifying the horizontal
    alignment.  Valid values come from constant \code{HALIGN_STYLES_}.} 

  \item{vertical}{a character value specifying the vertical 
    alignment.  Valid values come from constant \code{VALIGN_STYLES_}.} 
  
  \item{wrapText}{a logical indicating if the text should be wrapped.}

  \item{rotation}{a numerical value indicating the degrees you want to
    rotate the text in the cell.}

  \item{indent}{a numerical value indicating the number of spaces you
    want to indent the text in the cell.}
  
  \item{x}{An Alignment object, as returned by \code{Alignment}.}
  
}

\value{
  \code{Alignment} returns a list with components from the input
  argument, and a class attribute "Alignment".  Alignment objects are
  used when constructing cell styles.

  \code{is.Alignment} returns \code{TRUE} if the argument is of class
  "Alignment" and \code{FALSE} otherwise.
  
}
\author{Adrian Dragulescu}
\seealso{\code{\link{CellStyle}} for using the a \code{Alignment} object.}
\examples{

  # you can just use h for horizontal, since R does the matching for you
  a1 <-  Alignment(h="ALIGN_CENTER", rotation=90) # centered and rotated!
}


