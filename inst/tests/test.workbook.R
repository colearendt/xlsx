#
#
#

test.workbook <- function(type="xlsx")
{
  cat("##################################################\n")
  cat("Testing basic workbook functions\n")
  cat("##################################################\n")

  cat("Create an empty workbook ... ") 
  wb <- createWorkbook(type=type)
  cat("OK\n")

  cat("Create a sheet called 'Sheet1' ... ")
  sheet1 <- createSheet(wb, sheetName="Sheet1")
  cat("OK\n")

  cat("Create another sheet called 'Sheet2' ... ")
  sheet2 <- createSheet(wb, sheetName="Sheet2")
  cat("OK\n")
  
  cat("Get sheets ... ")
  sheets <- getSheets(wb)
  stopifnot(length(sheets) == 2)
  cat("OK\n")

  cat("Remove sheet named 'Sheet2' ... ")
  removeSheet(wb, sheetName="Sheet2")
  sheets <- getSheets(wb)  
  stopifnot(length(sheets) == 1)
  cat("OK\n")

  cat("Add rows 6:10 on Sheet1 ... ")
  rows <- createRow(sheet1, 6:10)
  stopifnot(length(rows) == 5)
  cat("OK\n")

  cat("Remove rows 1:10 on Sheet1 of test_import.xlsx ... ")
  filename <- paste("test_import.", type, sep="")
  file <- system.file("tests", filename, package="xlsx")
  wb <- loadWorkbook(file)  
  sheets <- getSheets(wb)
  sheet <- sheets[[1]]  
  rows  <- getRows(sheet)           # get all the rows
  removeRow(sheet, rows[1:10])
  rows  <- getRows(sheet)           # get all the rows
  stopifnot(length(rows) == 41)
  cat("OK\n")
  
}








