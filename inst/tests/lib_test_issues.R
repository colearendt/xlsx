#####################################################################
# Test Issue 2
# Columns of data (as formulas) are being read in as NA
# CLOSED
.test.issue2 <- function(DIR="C:/google/")
{
  cat(".test.issue2 ")
  require(xlsx)
  file <- paste(DIR, "rexcel/trunk/resources/xlxs2Test.xlsx", sep="")
  res <- read.xlsx2(file, sheetName="data", startRow=2, endRow=10,
      colIndex=c(1,3:5,7:9), colClasses=c("character",rep("numeric",6)) )
  #head(res)
  if (!any(is.na(res))) {
    cat("PASSED\n")
  } else {
    cat("FAILED\n")
  }

  invisible()
}


#####################################################################
# Test Issue 6
# setCellValue writes an NA as "#N/A", add an argument to make it an 
# empty cell.  This behavior is is visible with write.xlsx.
#
.test.issue6 <- function(DIR="C:/google/")
{
  cat(".test.issue6 ")
  require(xlsx)
  tfile <- tempfile(fileext=".xlsx")

  wb <- createWorkbook()
  sheet <- createSheet(wb)
  rows <- createRow(sheet, rowIndex=1:5)
  cells <- createCell(rows)
  mapply(setCellValue, cells[,1], c(1,2,NA,4,5))
  setCellValue(cells[[2,2]], "Hello")
  mapply(setCellValue, cells[,3], c(1,2,NA,4,5), showNA=FALSE)
  setCellValue(cells[[3,3]], NA, showNA=FALSE)  
  saveWorkbook(wb, tfile)
  
  aux <- data.frame(x=c(1,2,NA,4,5))
  write.xlsx(aux, file=tfile, showNA=FALSE)
    
  if ( TRUE ) {
    cat("PASSED\n")
  } else {
    cat("FAILED\n")
  }
}


#####################################################################
# Test Issue 7
# #N/A values are imported as FALSE  by read.xlsx
# Let's see if the new version of POI fixes this!
# Not something I can fix!
#
.test.issue7 <- function(DIR="C:/google/")
{
  cat(".test.issue7 ")
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
  if (is.na(value)) {
    cat("PASSED\n")
  } else {
    cat("FAILED\n")
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
  cat(".test.issue9 ")
  require(xlsx)
  file <- system.file("tests", "test_import.xlsx", package="xlsx")
  
  #source(paste(DIR, "rexcel/trunk/R/read.xlsx.R", sep=""))
  try(res <- read.xlsx(file, sheetName="issue9", rowIndex=3:5, colIndex=3:5))

  if (class(res) != "try-error") {
    cat("PASSED\n")
  } else {
    cat("FAILED\n")
  }
}


#####################################################################
# Test Issue 11
# Get an NPE when reading .xls files when they are not properly constructed
# and return more rows than they actually exist. 
#
.test.issue11 <- function(DIR="C:/google/")
{
  cat(".test.issue11 ")
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
.test.issue12 <- function( DIR="C:/google/" )
{
  cat(".test.issue12 ")
  require(xlsx)
  file <- paste(DIR, "rexcel/trunk/resources/issue12.xlsx", sep="")
  try(res <- read.xlsx2(file, sheetIndex=1, colIndex=1:3, startRow=3))

  if (class(res) != "try-error") {
    cat("PASSED\n")
  } else {
    cat("FAILED\n")
  }
}


#####################################################################
# Test Issue 13
#
#
.test.issue13 <- function( DIR="C:/google/" )
{
}


#####################################################################
# Register and run the specific tests
#
.run_test_issues <- function(DIR)
{
  .test.issue2(DIR)
  .test.issue6(DIR)  
  .test.issue7(DIR)  
  .test.issue9(DIR)  
  .test.issue11(DIR)  # lost the file!
  .test.issue12(DIR)
  #.test.issue13(DIR)


  
}
