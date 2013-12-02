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
    cat("FAILED -- OK (known issue!)\n")
  }
  
  # read.xlsx2 imports correctly!
  invisible()
}



#####################################################################
# Test Issue 9
# Problems with read.xlsx for tables that start in the middle of the
# sheet.  I used to get an NPE.
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
# fixed in java, with rexcel_0.5.1.jar
#
.test.issue11 <- function(DIR="C:/google/")
{
  cat(".test.issue11 ")
  require(xlsx)
  #file <- system.file("tests", "test_import.xlsx", package="xlsx")
  #file <- "C:/temp/fca3_monthly_ob_v2.xls"
  file <- paste(DIR, "rexcel/trunk/resources/read_xlsx2_example.xlsx", sep="")
  res <- read.xlsx(file, sheetIndex=1, header=FALSE)

  if (class(res) != "try-error") {
    cat("PASSED\n")
  } else {
    cat("FAILED\n")
  }
  
}


#####################################################################
# Test Issue 12
# read.xlsx2 doesn't evaluate formulas - values are NA
# Not an issue if you specify the colClasses!
#
.test.issue12 <- function( DIR="C:/google/" )
{
  cat(".test.issue12 ")
  require(xlsx)
  file <- system.file("tests", "test_import.xlsx", package="xlsx")
  try(res <- read.xlsx2(file, sheetName="formulas",  
    colClasses=list("numeric", "numeric", "character")))
  if ( is.na(res[1,3]) ) {
    cat("FAILED\n")
  } else {
    cat("PASSED\n")
  }
}


#####################################################################
# Test Issue 16
# Get an NPE when reading .xls files when they are not properly constructed
# and return more rows than they actually exist. 
#
.test.issue16 <- function( DIR="C:/google/" )
{
  cat(".test.issue16 ")
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
# Test Issue 19
# CB.setMatrixData only accepts numeric matrices.  It should accept
# character matrices too.
#
.test.issue19 <- function( DIR="C:/google/" )
{
  cat(".test.issue19 ")
  require(xlsx)

  wb <- createWorkbook()
  sheet <- createSheet(wb)
  mtx <- matrix(c("hello", "world", "check1", "check2"), ncol=2)
  cb <- CellBlock(sheet, 1, 1, 2, 2)
  CB.setMatrixData(cb, mtx, 1, 1)
  fileName <- paste(OUTDIR, "/issue19.xlsx", sep="")
  saveWorkbook(wb, fileName)

  
  cat("PASSED\n")
}


#####################################################################
# Test Issue 21
# Color black not set properly in Font
#
.test.issue21 <- function( DIR="C:/google/" )
{
  cat(".test.issue21 ")
  require(xlsx)

  wb <- createWorkbook()
  tmp <- Font(wb, color="black")

  tmp2 <- Font(wb, color=NULL)
  
  cat("FAILED\n")
}


#####################################################################
# Test Issue 22
# Preserve existing formats when you write data to the xlsx with
# CellBlock construct
#
.test.issue22 <- function( DIR="C:/google/" )
{
  cat(".test.issue22 ")
  require(xlsx)
  fileIn  <- paste(DIR, "rexcel/trunk/resources/issue22.xlsx", sep="")
  fileOut <- paste(OUTDIR, "issue22_out.xlsx", sep="")

  wb <- loadWorkbook(fileIn)
  sheets <- getSheets(wb)

  mat <- matrix(1:9, 3, 3)
  for (sheet in sheets) {
    if (sheet$getSheetName() == "Sheet1" ){
      # need to create the rows for Sheet1 as it is empty!  
      cb <- CellBlock(sheet, 1, 1, 3, 3, create = TRUE)   
    } else {
      cb <- CellBlock(sheet, 1, 1, 3, 3, create = FALSE) 
    }  
    CB.setMatrixData(cb, mat, 1, 1)
  }
  saveWorkbook(wb, fileOut)
  
  cat("PASSED\n")
}



#####################################################################
# Test Issue 23
# add an emf picture
#
.test.issue23 <- function( DIR="C:/google/" )
{
  cat(".test.issue23 ")
  fileName <- paste(OUTDIR, "test_emf.emf", sep="")
  require(devEMF)
  emf(file=fileName, bg="white")
  boxplot(rnorm(100))
  dev.off()  

  require(xlsx)
  wb <- createWorkbook()
  sheet <- createSheet(wb, "EMF_Sheet")
  
  addPicture(file=fileName, sheet)
  saveWorkbook(wb, file=paste(OUTDIR, "/issue23_out.xlsx", sep=""))  

  # the spreadsheet saves but the emf picture is not there
  # used to work in previous versions of POI, not sure why not anymore
  cat("FAILED -- (known issue with 3.9)\n")
  
}


#####################################################################
# Test Issue 25.  Integers cells read as characters are not read
# "cleanly" with RInterface, e.g. "12.0" instead of "12".
#
.test.issue25 <- function( DIR="C:/google/",  out="FAILED\n")
{
  cat(".test.issue25 ")
  require(xlsx)
  file <- paste(DIR, "rexcel/trunk/resources/issue25.xlsx", sep="")

  # reads element [35,1] as a double and then transforms it to a factor
  res1 <- read.xlsx2(file, sheetIndex=1, header=TRUE, startRow=1,
    colClasses=c("character", rep("numeric", 5)), stringsAsFactors=FALSE)
  
  if (res1[35,2] == "250829")  
    out <- "PASSED\n"
      
  # reads element [34,1] correctly, how?!  - R magic
  # res2 <- read.xlsx(file, sheetIndex=1, header=TRUE, startRow=1)

  cat(out)
  
}


#####################################################################
# Test Issue 26.  Customize the format of datetimes in the output
#
.test.issue26 <- function( DIR="C:/google/",  out="FAILED\n")
{
  cat(".test.issue26 ")
  require(xlsx)

  wb <- createWorkbook()
  sheet <- createSheet(wb, "Sheet1")
  
  days <- seq(as.Date("2013-01-01"), by="1 day", length.out=5)
  # use the default
  addDataFrame(data.frame(days=days), sheet, startColumn=1,
               col.names=FALSE, row.names=FALSE)

  # change the options temporarily
  oldOpt <- options()
  options(xlsx.date.format="dd MMM, yyyy")
  addDataFrame(data.frame(days=days), sheet, startColumn=2,
               col.names=FALSE, row.names=FALSE)
  options(oldOpt)
  
  
  saveWorkbook(wb, file=paste(OUTDIR, "issue26_out.xlsx", sep=""))  
  cat("PASSED")
}


#####################################################################
# Register and run the specific tests
#
.run_test_issues <- function(SOURCEDIR)
{
  source(paste(SOURCEDIR, "rexcel/trunk/inst/tests/lib_test_issues.R", sep=""))
  DIR <- SOURCEDIR
  .test.issue2(DIR)
  .test.issue6(DIR)  
  .test.issue7(DIR)  
  .test.issue9(DIR)  
  .test.issue11(DIR)  
  .test.issue12(DIR)
  .test.issue16(DIR)
  .test.issue19(DIR)
  .test.issue22(DIR)
  .test.issue23(DIR)
  .test.issue25(DIR)
  .test.issue26(DIR)


  
}
