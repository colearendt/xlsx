context('read.xlsx')

test_that('#N/A values are not imported as FALSE', {
  ## issue #7
  skip('not presently working')
  
  file <- test_ref("issue7.xlsx")
  res <- read.xlsx(file, sheetIndex=1, rowIndex=2:5, colIndex=2:3)
  
  wb <- loadWorkbook(file)
  sheet <- getSheets(wb)[[1]]
  rows  <- getRows(sheet)   # get all the rows
  
  cells <- getCells(rows["4"])   # returns all non empty cells
  cell  <- cells[["4.2"]]
  value <- getCellValue(cell)
  
  expect_true(is.na(value))
})

test_that('works for tables that start in the middle of the sheet', {
  ##issue #9
  file <- test_ref("test_import.xlsx")
  
  #source(paste(DIR, "rexcel/trunk/R/read.xlsx.R", sep=""))
  try(res <- read.xlsx(file, sheetName="issue9", rowIndex=3:5, colIndex=3:5))
  
  expect_is(class(res),"character")
})

test_that('read improperly constructed .xls files', {
  ## issue #11
  file <- test_ref("read_xlsx2_example.xlsx")
  res <- read.xlsx(file, sheetIndex=1, header=FALSE)
  
  expect_is(res,'data.frame')
})

test_that('read does not fail on empty sheets', {
  ## issue #28
  file <- test_ref("issue25.xlsx")
  
  # reads Sheet2 which is empty
  res <- read.xlsx(file, sheetIndex=2)
  res2 <- read.xlsx2(file, sheetIndex=2)
  
  expect_null(res)
  expect_null(res2)
})

test_that('read.xlsx fails on empty row', {
  ## issue #57
  try(aux <- read.xlsx(test_ref("issue57.xlsx"), sheetIndex=1))  
  
  expect_is(aux,'data.frame')
})

test_that('works in complex pipeline', {
  skip("Not working on mac R 3.3.3")
  test_complex_read.xlsx('xls')
  test_complex_read.xlsx('xlsx')
})

test_that('read password protected workbook succeeds', {
  ## issue #49
  filename <- test_ref('issue49_password=test.xlsx')
  df <- read.xlsx(filename, sheetIndex=1, password='test', stringsAsFactors=FALSE)
  
  expect_identical(df,data.frame(Values=c(1,2,3),stringsAsFactors=FALSE))
})

context('read.xlsx2')

test_that('columns of data (as formulas) are not read in as NA', {
  ## issue #2
  file <- test_ref("xlxs2Test.xlsx")
  res <- read.xlsx2(file, sheetName="data", startRow=2, endRow=10,
                    colIndex=c(1,3:5,7:9), colClasses=c("character",rep("numeric",6)) )
  
  expect_false(any(is.na(res)))
})

test_that('does not evaluate formulas - values are NA', {
  skip('does not presently work')
  ## issues #12
  file <- test_ref("test_import.xlsx")
  try(res <- read.xlsx2(file, sheetName="formulas",  
                        colClasses=list("numeric", "numeric", "character")))
  expect_true(is.na(res[1,3]))
})

test_that('read improperly constructed .xls files', {
  ## issue #16
  file <- test_ref("issue12.xlsx")
  try(res <- read.xlsx2(file, sheetIndex=1, colIndex=1:3, startRow=3))
  
  expect_is(res,'data.frame')
})

test_that('integer cells read as characters are read cleanly', {
  ## issue #25
  file <- test_ref("issue25.xlsx")
  
  # reads element [35,1] as a double and then transforms it to a factor
  res1 <- read.xlsx2(file, sheetIndex=1, header=TRUE, startRow=1,
                     colClasses=c("character", rep("numeric", 5)), stringsAsFactors=FALSE)
  
  expect_equal(res1[35,2],250829)
})

test_that('no argument rowIndex, and maybe it should be', {
  ## issue #58
  skip('need to add rowIndex argument')
  try(aux <- read.xlsx2(test_ref("issue58.xlsx"), sheetIndex=1,
                        rowIndex=1:8))
})

test_that('works in complex pipeline', {
  test_complex_read.xlsx2('xls')
  test_complex_read.xlsx2('xlsx')
})

test_that('read password protected workbook succeeds', {
  ## issue #49
  filename <- test_ref('issue49_password=test.xlsx')
  df <- read.xlsx2(filename, sheetIndex=1, password='test',stringsAsFactors=FALSE)
  
  expect_identical(df,data.frame(Values=c('1','2','3'),stringsAsFactors=FALSE))
})

context('low-level interface')

test_that('works in pipeline', {
  test_basic_import('xlsx')
  test_basic_import('xls')
})
