\name{readColumns}
\alias{readColumns}
\title{ Read a contiguous set of columns from sheet into an R data.frame}
\description{
  Read a contiguous set of columns from sheet into an R data.frame.
  Uses the \code{RInterface} for speed.   
}
\usage{
readColumns(sheet, startColumn, endColumn, startRow,
  endRow=NULL, as.data.frame=TRUE, header=TRUE, colClasses=NA,
  ...)
}
\arguments{
  \item{sheet}{a \code{\link{Worksheet}} object.}

  \item{startColumn}{a numeric value for the starting column.}

  \item{endColumn}{a numeric value for the ending column.}

  \item{startRow}{a numeric value for the starting row.}
  
  \item{endRow}{a numeric value for the ending row.  If \code{NULL} it
    reads all the rows in the sheet.}  
  
  \item{colClasses}{a character vector that represent the class of each
    column.  Recycled as necessary, or if \code{NA} an attempt is made
    to guess the type of each column by reading the first row of data.}
    
  \item{\ldots}{other arguments to \code{data.frame}, for example
    \code{stringsAsFactors}} 
}

\details{
  
Use the \code{readColumns} function when you want to read a rectangular
block of data from an Excel worksheet.  It calls several low level
functions in the process.  

Excel internally stores dates and datetimes as numeric values, and does
not keep track of time zones and DST.  When a numeric column is
formatted as a datetime, it will be converted into \code{POSIXct} class
with a \emph{GMT} timezone.  If you need a \code{Date} column, you need
to specify explicitly using \code{colClasses} argument.

Blannk cells in the middle of your data, will be returned as NA's.

}
\value{
  A data.frame or a list, depending on the \code{as.data.frame}
  argument. 
}
\author{ Adrian Dragulescu }
\seealso{\code{\link{read.xlsx2}} for reading entire sheets.
  See also \code{\link{addDataFrame}} for writing a \code{data.frame} to a
  sheet.}
\examples{
\dontrun{

file <- system.file("tests", "test_import.xlsx", package = "xlsx")
res <- read.xlsx(file, 1)  # read first sheet
head(res)
#          NA. Population Income Illiteracy Life.Exp Murder HS.Grad Frost   Area
# 1    Alabama       3615   3624        2.1    69.05   15.1    41.3    20  50708
# 2     Alaska        365   6315        1.5    69.31   11.3    66.7   152 566432
# 3    Arizona       2212   4530        1.8    70.55    7.8    58.1    15 113417
# 4   Arkansas       2110   3378        1.9    70.66   10.1    39.9    65  51945
# 5 California      21198   5114        1.1    71.71   10.3    62.6    20 156361
# 6   Colorado       2541   4884        0.7    72.06    6.8    63.9   166 103766
# >


# To convert an Excel datetime colum to POSIXct, do something like:
#   as.POSIXct((x-25569)*86400, tz="GMT", origin="1970-01-01")
# For Dates, use a conversion like:
#   as.Date(x-25569, origin="1970-01-01") 

res2 <- read.xlsx2(file, 1)  # all columns returned characters

}
}

