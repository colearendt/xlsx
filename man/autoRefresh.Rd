\name{ForceRefresh}
\alias{forceFormulaRefresh}
\alias{forcePivotTableRefresh}
\title{Force Refresh Pivot Tables and Formulae}
\description{
Functions to force formula calculation or refresh of pivot tables when the Excel file is opened.
}
\usage{

forcePivotTableRefresh(file, output=NULL, verbose=FALSE)

forceFormulaRefresh(file, output=NULL, verbose=FALSE)

}
\arguments{
  \item{file}{the path of the source file where formulae/pivot table needs to be refreshed}
  \item{output}{the path of the output file.  If it is \code{NULL} then the source 
    file will be overwritten.}
  \item{verbose}{a boolean.  Set to \code{TRUE} if you want details.}
}

\details{
  \code{forcePivotTableRefresh} forces pivot tables to be refreshed when the Excel file is opened.

  \code{forceFormulaRefresh} forces formulae to be recalculated when the Excel file is opened.

}
\value{
  \code{forcePivotTableRefresh} does not return any results
  
  \code{forceFormulaRefresh} does not return any results

}
\author{Tom Kwong}
\examples{
\dontrun{
# Patch a file where its pivot tables are not recalculated when the file is opened
forcePivotTableRefresh("/tmp/file.xlsx")
forcePivotTableRefresh("/tmp/file.xlsx", "/tmp/fixed_file.xlsx")

# Patch a file where its formulae are not recalculated when the file is opened
forceFormulaRefresh("/tmp/file.xlsx")
forceFormulaRefresh("/tmp/file.xlsx", "/tmp/fixed_file.xlsx")
}

}

