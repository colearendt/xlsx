# Some examples of low level 
#
#

test.lowlevel <- function()
{
  require(xlsx)
  #print(.jclassPath())

  cat("Create an empty workbook ...\n") 
  wb <- .jnew("org/apache/poi/xssf/usermodel/XSSFWorkbook")
  #if (class(wb)=="jobjRef") cat("OK.\n")
  #.jmethods(wb)

  cat("Create a sheet called 'Sheet2' ...\n")
  sheet2 <- .jcall(wb, "Lorg/apache/poi/ss/usermodel/Sheet;",
    "createSheet", "Sheet2")
  #if (.jinstanceof(sheet2, "org.apache.poi.ss.usermodel.Sheet"))
  #  cat("OK.\n")  
  #.jmethods(sheet2)

  cat("Create row 1 ...\n")
  row <- .jcall(sheet2, "Lorg/apache/poi/ss/usermodel/Row;",
    "createRow", as.integer(0))
  #.jmethods(row)

  cat("Create cell [1,1] ...\n")
  cell <- .jcall(row, "Lorg/apache/poi/ss/usermodel/Cell;",
    "createCell", as.integer(0))

  cat("Put a value in cell [1,1] ... ")
  cell$setCellValue(1.23)

  cat("Add cell [1,2] and put a numeric value ... \n")
  .jcall(row, "Lorg/apache/poi/ss/usermodel/Cell;",
    "createCell", as.integer(1))$setCellValue(3.1415)

  cat("Add cell [1,3] and put a stringvalue ... \n")
  .jcall(row, "Lorg/apache/poi/ss/usermodel/Cell;",
    "createCell", as.integer(2))$setCellValue("A string")

  cat("Add cell [1,4] and put a boolean value ... \n")
  .jcall(row, "Lorg/apache/poi/ss/usermodel/Cell;",
    "createCell", as.integer(3))$setCellValue(TRUE)

  file <- "C:/Users/adrian/R/findataweb/temp/xlsx/tests/test_lowlevel.xlsx"
  saveWorkbook(wb, file)
  cat("Wrote file:", file, "\n\n")
}

## # make a file handle and write the xlsx to file
## filename <- "C:/Temp/junk.xlsx"
## fh <- .jnew("java/io/FileOutputStream", filename)
## .jcall(wb, "V", "write", .jcast(fh, "java/io/OutputStream"))
## .jcall(fh, "V", "close")




  ## this works using the magic of J
  #sheet1 <- J(wb, "createSheet", "Sheet1")
  #.jinstanceof(sheet1, "org.apache.poi.ss.usermodel.Sheet")







