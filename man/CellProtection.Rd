\name{CellProtection}
\alias{CellProtection}
\alias{is.CellProtection}
\title{Create a CellProtection object.}
\description{
  Create a CellProtection object used for cell styles.
}
\usage{

CellProtection(locked=TRUE, hidden=FALSE)

is.CellProtection(x)

}
\arguments{

  \item{locked}{a logical indicating the cell is locked.}

  \item{hidden}{a logical indicating the cell is hidden.}

  \item{x}{A CellProtection object, as returned by \code{CellProtection}. }
  
}

%\details{}
\value{
  
  \code{CellProtection} returns a list with components from the input
    argument, and a class attribute "CellProtection".  CellProtection
    objects are used when constructing cell styles.

  \code{is.CellProtection} returns \code{TRUE} if the argument is of class
    "CellProtection" and \code{FALSE} otherwise.

}
\author{Adrian Dragulescu}
\seealso{\code{\link{CellStyle}} for using the a \code{CellProtection} object.}
\examples{

  font <-  CellProtection(locked=TRUE)
}


