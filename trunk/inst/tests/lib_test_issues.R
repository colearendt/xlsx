#####################################################################
# Test Issue 2
# Columns of data (as formulas) are being read in as NA
# CLOSED
.test.issue2 <- function(DIR="C:/google/")
{
  require(xlsx)
  file <- paste(DIR, "rexcel/trunk/resources/xlxs2Test.xlsx", sep="")
  res <- read.xlsx2(file, sheetName="data", startRow=2, endRow=10,
      colIndex=c(1,3:5,7:9), colClasses=c("character",rep("numeric",6)) )
  #head(res)
  if (!any(is.na(res))) {
    cat(".test.issue2 PASSED\n")
  } else {
    cat(".test.issue2 FAILED\n")
  }

  #res <- read.xlsx2(file, sheetName="data", startRow=2, endRow=10,
  #   colIndex=c(1,4:5,8:9), colClasses=c("character",rep("numeric",4)) )

  invisible()
}


#####################################################################
# Test Issue 7
# #N/A values are imported as FALSE  by read.xlsx
# Let's see if the new version of POI fixes this!
#
.test.issue7 <- function(DIR="C:/google/")
{
  require(xlsx)
  file <- paste(DIR, "rexcel/trunk/resources/issue7.xlsx", sep="")
  res <- read.xlsx(file, sheetIndex=1, rowIndex=2:5, colIndex=2:3)

  wb <- loadWorkbook(file)
  sheet <- getSheets(wb)[[1]]
  rows  <- getRows(sheet)   # get all the rows

  cells <- getCells(rows["4"])   # returns all non empty cells
  cell  <- cells[["4.2"]]
  value <- getCellValue(cell)

  # value should be NA, but it is FALSE!
  if (value == NA) {
    cat(".test.issue7 PASSED\n")
  } else {
    cat(".test.issue7 FAILED\n")
  }
  
  # read.xlsx2 imports correctly!
  invisible()
}



#####################################################################
# Test Issue 9
# Problems with read.xlsx for tables that start in the middle of the
# sheet.  I get an NPE.
#
.test.issue9 <- function(DIR="C:/google/")
{
  require(xlsx)
  file <- system.file("tests", "test_import.xlsx", package="xlsx")
  
  #source(paste(DIR, "rexcel/trunk/R/read.xlsx.R", sep=""))
  try(res <- read.xlsx(file, sheetName="issue9", rowIndex=3:5, colIndex=3:5))

  if (class(res) != "try-error") {
    cat(".test.issue9 PASSED\n")
  } else {
    cat(".test.issue9 FAILED\n")
  }
}


#####################################################################
# Test Issue 11
# Get an NPE when reading .xls files when they are not properly constructed
# and return more rows than they actually exist. 
#
.test.issue11 <- function(DIR="C:/google/")
{
  require(xlsx)
  #file <- system.file("tests", "test_import.xlsx", package="xlsx")
  #file <- "C:/temp/fca3_monthly_ob_v2.xls"
  file <- paste(DIR, "rexcel/trunk/resources/read_xlsx2_example.xlsx", sep="")
  res <- read.xlsx(file, sheetIndex=1, header=FALSE)
  
  #source(paste(DIR, "rexcel/trunk/R/readColumns.R", sep=""))
  #res <- read.xlsx2(file, sheetIndex=1, startRow=3)
  
  
}


#####################################################################
# Test Issue 12
# Get an NPE when reading .xls files when they are not properly constructed
# and return more rows than they actually exist. 
#
.test.issue12 <- function(DIR)
{
  require(xlsx)
  #file <- system.file("tests", "test_import.xlsx", package="xlsx")
  file <- "C:/temp/read_xlsx2_example.xlsx"
  res <- read.xlsx2(file, sheetIndex=1, colIndex=1:3)
}


#####################################################################
# Register and run the specific tests
#
.run_test_issues(DIR)
{
  
  .test.issue12(DIR)
}
